using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "Add a new board" (NewContributeStrategy.md Phase 2, session 2c, task 9) - asks for the
    // manufacturer, hardware and board that identify a brand-new board, and returns the
    // registration to create a local draft from. Returns null via ShowDialog when cancelled.
    //
    // It only RESOLVES the answer; the caller (Main.NewBoard.cs) performs the create, the same
    // division WorklogAddPhotoWindow and WorklogAttachCaptureWindow already keep. That is what lets
    // the create path stay one sequence in one place rather than being split across a dialog and
    // its caller.
    //
    // The three names are folder segments as well as display names, so every validation rule here
    // exists to stop a board being created that looks fine and is then broken in a way the user
    // cannot diagnose - see NewBoardIdentity.IsValidPathSegment for what each rule prevents.
    //
    // Enter confirms and Escape cancels, the ordinary modal treatment. Deliberately NOT the
    // Tunnel-both-cancel handling DiscardDraftWindow and DeleteWorkbookWindow use: that exists
    // because their submit button performs a permanent delete, and this one creates something.
    //
    // *** CREATING REQUIRES ACCEPTING THE MAINTAINER ROLE (owner request, 2026-09-24). ***
    // "Create board" first shows NewBoardMaintainerWindow, and the registration is handed back
    // ONLY when that is accepted. Declining returns to this form with nothing created and every
    // field still filled in - so a mis-click costs one more click, not the typing. The gate lives
    // here rather than in Main so that no caller can obtain a registration without it: this
    // window is the one way a new board's identity comes into being.
    // ###########################################################################################
    public partial class NewBoardWindow : Window
    {
        // Every "{HardwareName}|{BoardName}" already taken, so a new board cannot collide with one.
        // Seeded by Initialize from DataManager.HardwareBoards, which by then already includes the
        // main workbook's boards, any _UserContribution ones AND other drafts' registrations - so
        // this single check covers all three sources.
        private readonly HashSet<string> thisTakenIdentities = new(StringComparer.OrdinalIgnoreCase);

        // Replaces the maintainer agreement with a canned answer. Null in the app, which shows
        // NewBoardMaintainerWindow; set by headless tests that want to drive the create path
        // without answering a real dialog. The same override-else-real shape as TabWorkbooks'
        // BoardKeyOverrideForTests, so the tests still run the real ConfirmAndCloseAsync around it.
        internal Func<NewBoardRegistration, Task<bool>>? ConfirmMaintainerRoleOverrideForTests { get; set; }

        public NewBoardWindow()
        {
            this.InitializeComponent();

            this.ManufacturerBox.TextChanged += this.OnAnyFieldChanged;
            this.HardwareBox.TextChanged += this.OnAnyFieldChanged;
            this.BoardBox.TextChanged += this.OnAnyFieldChanged;

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);

            this.UpdateValidation();
        }

        // ###########################################################################################
        // Seeds the manufacturer suggestions and the set of names already in use. Takes the entry
        // list rather than reading DataManager.HardwareBoards itself, the same seam every other
        // testable surface in this codebase uses (TabWorkbooks.CurrentBoardDataOverrideForTests,
        // TabDrafts.HardwareBoardsOverrideForTests) - a headless test cannot populate that static
        // without loading a real main workbook.
        // ###########################################################################################
        public void Initialize(IReadOnlyList<HardwareBoardEntry> existingBoards)
        {
            this.thisTakenIdentities.Clear();

            foreach (var entry in existingBoards)
            {
                this.thisTakenIdentities.Add($"{entry.HardwareName}|{entry.BoardName}");
            }

            // Offering the manufacturers already in use is what stops a typo ("Comodore") quietly
            // creating a second top-level folder beside the real one - the names are free text, so
            // nothing else would catch it.
            this.ManufacturerBox.ItemsSource = existingBoards
                .Select(entry => NewBoardIdentity.ExtractManufacturer(entry.ExcelDataFile))
                .Where(manufacturer => !string.IsNullOrWhiteSpace(manufacturer))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(manufacturer => manufacturer, StringComparer.OrdinalIgnoreCase)
                .ToList();

            this.UpdateValidation();
        }

        // ###########################################################################################
        // The registration the current field values describe, or null when they are not valid. One
        // method for both the live validation and the create click, so the button can never be
        // enabled for input the create path would then reject.
        // ###########################################################################################
        internal NewBoardRegistration? BuildRegistrationOrNull(out string validationMessage)
        {
            string manufacturer = NewBoardIdentity.SanitizePathSegment(this.ManufacturerBox.Text);
            string hardware = NewBoardIdentity.SanitizePathSegment(this.HardwareBox.Text);
            string board = NewBoardIdentity.SanitizePathSegment(this.BoardBox.Text);

            // Blank fields are the starting state, not a mistake worth shouting about - the button
            // is simply disabled until all three are filled in.
            if (manufacturer.Length == 0 || hardware.Length == 0 || board.Length == 0)
            {
                validationMessage = string.Empty;
                return null;
            }

            if (!NewBoardIdentity.IsValidPathSegment(manufacturer, out string manufacturerReason))
            {
                validationMessage = $"Manufacturer {manufacturerReason}.";
                return null;
            }

            if (!NewBoardIdentity.IsValidPathSegment(hardware, out string hardwareReason))
            {
                validationMessage = $"Hardware {hardwareReason}.";
                return null;
            }

            if (!NewBoardIdentity.IsValidPathSegment(board, out string boardReason))
            {
                validationMessage = $"Board {boardReason}.";
                return null;
            }

            if (this.thisTakenIdentities.Contains($"{hardware}|{board}"))
            {
                validationMessage = $"[{hardware}] / [{board}] already exists. Pick a different hardware or board name.";
                return null;
            }

            string excelDataFile = NewBoardIdentity.BuildExcelDataFile(manufacturer, hardware, board);
            if (string.IsNullOrWhiteSpace(excelDataFile))
            {
                validationMessage = "Could not build a folder path from those names.";
                return null;
            }

            // A folder already holding a draft means a board was created here before and is still
            // live - creating over it would overwrite that contributor's own work. This can happen
            // when the hardware/board names differ from a previous board's only by the display
            // name, so the identity check above passes while the FOLDER still collides.
            //
            // Asked of the MARKER since Phase 6, which is what makes a folder a draft. A board
            // folder copied in by hand gets one when the board list loads (DraftFolderImport), so
            // one that was there at startup is refused here too.
            if (DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile))
            {
                validationMessage = "A draft already exists in that folder. Pick a different name.";
                return null;
            }

            validationMessage = string.Empty;

            return new NewBoardRegistration
            {
                HardwareName = hardware,
                BoardName = board,
                HardwareNotes = this.NotesBox.Text?.Trim() ?? string.Empty,
                ExcelDataFile = excelDataFile,
                CreatedUtc = DateTime.UtcNow.ToString("o"),
            };
        }

        private void OnAnyFieldChanged(object? sender, EventArgs e) => this.UpdateValidation();

        // ###########################################################################################
        // Re-runs validation against whatever the boxes currently hold. Exposed to the headless
        // tests because a window that is never attached to a visual tree never raises TextChanged,
        // so setting .Text in a test does not reach OnAnyFieldChanged - the same trap TabWorkbooks
        // documents for its own search box. This is the real path either way: it is exactly what
        // the TextChanged handler calls, so a test driving it exercises the shipped logic rather
        // than a parallel copy.
        // ###########################################################################################
        internal void RefreshValidationForTests() => this.UpdateValidation();

        // ###########################################################################################
        // Keeps the preview, the message and the Create button in step with the fields. The button
        // is enabled only when a registration could actually be built, so an invalid create is
        // unreachable rather than merely discouraged.
        // ###########################################################################################
        private void UpdateValidation()
        {
            var registration = this.BuildRegistrationOrNull(out string validationMessage);

            this.CreateButton.IsEnabled = registration != null;

            this.ValidationText.Text = validationMessage;
            this.ValidationText.IsVisible = !string.IsNullOrWhiteSpace(validationMessage);

            this.PreviewText.Text = registration == null
                ? "Fill in the manufacturer, hardware and board above."
                : $"{registration.HardwareName} - {registration.BoardName}\n{registration.ExcelDataFile}";
        }

        // Exposed for the headless tests, which cannot read a TextBlock's enabled state meaningfully
        // without also knowing what the window thinks the identity is.
        internal string PreviewTextForTests => this.PreviewText.Text ?? string.Empty;

        internal string ValidationTextForTests =>
            this.ValidationText.IsVisible ? this.ValidationText.Text ?? string.Empty : string.Empty;

        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close(null);
                e.Handled = true;
                return;
            }

            // Enter confirms, but only from a state that could be confirmed by clicking - the same
            // rule the button itself follows.
            if (e.Key == Key.Enter && this.CreateButton.IsEnabled)
            {
                this.OnCreateClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private async void OnCreateClick(object? sender, RoutedEventArgs e)
        {
            await this.ConfirmAndCloseAsync();
        }

        // ###########################################################################################
        // The create path: validate, ask for the maintainer agreement, and only then close with the
        // registration. Returns what the window closed with - the registration when accepted, null
        // when the input was invalid or the agreement declined (the window then stays open).
        // ###########################################################################################
        internal async Task<NewBoardRegistration?> ConfirmAndCloseAsync()
        {
            var registration = this.BuildRegistrationOrNull(out _);
            if (registration == null)
            {
                // Unreachable through the UI (the button is disabled and Enter checks the same
                // flag), but a create with invalid input must never fall through to the caller.
                return null;
            }

            var confirmMaintainerRole = this.ConfirmMaintainerRoleOverrideForTests ?? this.ShowMaintainerAgreementAsync;

            if (!await confirmMaintainerRole(registration))
            {
                return null;
            }

            this.Close(registration);

            return registration;
        }

        private async Task<bool> ShowMaintainerAgreementAsync(NewBoardRegistration registration)
        {
            var agreement = new NewBoardMaintainerWindow();
            agreement.Initialize($"{registration.HardwareName} - {registration.BoardName}");

            return await agreement.ShowDialog<bool?>(this) == true;
        }
    }
}
