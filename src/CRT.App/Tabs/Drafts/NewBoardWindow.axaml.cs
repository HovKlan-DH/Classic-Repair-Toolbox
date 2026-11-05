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
    // The fields and every rule about them are BoardNameFields' - the Contribute tab's "Board name"
    // shows the same (2026-11-01).
    //
    // Escape cancels. Enter presses "Create board", the DEFAULT button, once the form is filled in - an
    // Enter nobody else handled: in the notes box it starts a new line, on a focused button it presses
    // that one, and on a suggestion of the manufacturer box it picks it (code review, 2026-11-02: the
    // window took Enter itself on the tunnel route, before any of them could have it, as
    // SubmitDraftWindow did until 2026-10-13). A disabled default button does nothing.
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
        // Replaces the maintainer agreement with a canned answer. Null in the app, which shows
        // NewBoardMaintainerWindow; set by headless tests that want to drive the create path
        // without answering a real dialog. The same override-else-real shape as TabWorkbooks'
        // BoardKeyOverrideForTests, so the tests still run the real ConfirmAndCloseAsync around it.
        internal Func<NewBoardRegistration, Task<bool>>? ConfirmMaintainerRoleOverrideForTests { get; set; }

        public NewBoardWindow()
        {
            this.InitializeComponent();

            this.Fields.FieldsChanged += (_, _) => this.UpdateValidation();

            this.AddHandler(KeyDownEvent, this.OnWindowKeyDown, RoutingStrategies.Tunnel);

            this.UpdateValidation();
        }

        // Seeds the fields' suggestions and the set of names already in use (BoardNameFields).
        public void Initialize(IReadOnlyList<HardwareBoardEntry> existingBoards)
        {
            this.Fields.Initialize(existingBoards);
            this.UpdateValidation();
        }

        internal string? ManufacturerOfferedForTests => this.Fields.ManufacturerOfferedForTests;

        internal string? HardwareOfferedForTests => this.Fields.HardwareOfferedForTests;

        internal NewBoardRegistration? BuildRegistrationOrNull(out string validationMessage) =>
            this.Fields.BuildRegistrationOrNull(out validationMessage);

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
            this.Fields.ShowOffers();

            var registration = this.BuildRegistrationOrNull(out string validationMessage);

            this.CreateButton.IsEnabled = registration != null;

            this.ValidationText.Text = validationMessage;
            this.ValidationText.IsVisible = !string.IsNullOrWhiteSpace(validationMessage);

            this.PreviewText.Text = registration == null
                ? "Fill in the manufacturer, hardware and board above."
                : NewBoardWindow.PreviewOf(registration);
        }

        // What "This will create" says (owner request, 2026-11-03): where the board will be seen, and the
        // file it will be - from the top of the data, so "/" leads.
        internal static string PreviewOf(NewBoardRegistration registration) =>
            $"Hardware drop-down will show: {registration.HardwareName}\n" +
            $"Board drop-down will show: {registration.BoardName}\n" +
            $"Path and datafile will be: /{registration.ExcelDataFile}";

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
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private async void OnCreateClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ConfirmAndCloseAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT
                // (code review, 2026-10-13 - AsyncVoidHandlersGuardTests).
                Logger.Critical($"OnCreateClick failed - [{ex}]");
            }
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
                // Unreachable through the UI (the button is disabled, and a disabled default button
                // takes no Enter), but a create with invalid input must never fall through to the caller.
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
