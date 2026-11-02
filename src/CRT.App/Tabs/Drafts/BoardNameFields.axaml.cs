using Avalonia.Controls;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE FIELDS THAT NAME A NEW BOARD - manufacturer, hardware, board and notes, and every rule about
    // them: the one form for "Add a new board" (NewBoardWindow) and "Board name" on the Contribute tab
    // (BoardNamePanel; owner request, 2026-11-01: "how to edit those details again afterwards"), so a
    // name refused when a board is made is refused when it is changed.
    //
    // The three names are folder segments as well as display names, so every validation rule here
    // exists to stop a board being created that looks fine and is then broken in a way the user
    // cannot diagnose - see NewBoardIdentity.IsValidPathSegment for what each rule prevents.
    //
    // *** "DID YOU MEAN COMMODORE?" (CRT 4.0 step 3, slice 3e). *** A manufacturer or machine written
    // another way than CRT has it becomes a second one beside it in the drop-down lists, so a name
    // close to one CRT has is offered under its box - a machine by its name or its folder, "C64"
    // offering "Commodore 64" (SimilarNames). Never in the way: the name typed may be a new one.
    // ###########################################################################################
    public partial class BoardNameFields : UserControl
    {
        // Every "{HardwareName}|{BoardName}" already taken, so a board cannot collide with one - the main
        // workbook's boards, any _UserContribution ones AND other drafts' registrations, all in
        // DataManager.HardwareBoards by then. Without the board being renamed.
        private readonly HashSet<string> thisTakenIdentities = new(StringComparer.OrdinalIgnoreCase);

        // Every other listed board's FOLDER, "Manufacturer/Hardware/Board", with how the board is listed -
        // see BuildRegistrationOrNull.
        private readonly Dictionary<string, string> thisTakenFolders = new(StringComparer.OrdinalIgnoreCase);

        // The manufacturers and machines CRT has, for "Did you mean": a machine by its name and by its
        // folder ("C64" is "Commodore 64").
        private List<KnownName> thisKnownManufacturers = [];
        private List<KnownName> thisKnownHardware = [];

        // The board being renamed - its own folder is no collision. Empty for a new board.
        private string thisOwnIdentity = string.Empty;

        // Told when anything is typed, or a name offered is taken.
        internal event EventHandler? FieldsChanged;

        public BoardNameFields()
        {
            this.InitializeComponent();

            this.ManufacturerBox.TextChanged += this.OnAnyFieldChanged;
            this.HardwareBox.TextChanged += this.OnAnyFieldChanged;
            this.BoardBox.TextChanged += this.OnAnyFieldChanged;
            this.NotesBox.TextChanged += this.OnAnyFieldChanged;

            // "Use Commodore": the name offered goes into its box.
            this.ManufacturerDidYouMeanButton.Click += (_, _) => this.UseOffered(this.ManufacturerDidYouMeanButton, text => this.ManufacturerBox.Text = text);
            this.HardwareDidYouMeanButton.Click += (_, _) => this.UseOffered(this.HardwareDidYouMeanButton, text => this.HardwareBox.Text = text);
        }

        // ###########################################################################################
        // Seeds the manufacturer suggestions and the set of names already in use. Takes the entry list
        // rather than reading DataManager.HardwareBoards itself, the same seam every other testable
        // surface in this codebase uses - a headless test cannot populate that static without loading
        // a real main workbook. `ownIdentity`: the board being renamed, left out of both.
        // ###########################################################################################
        internal void Initialize(IReadOnlyList<HardwareBoardEntry> existingBoards, string ownIdentity = "")
        {
            this.thisOwnIdentity = ownIdentity?.Trim() ?? string.Empty;
            this.thisTakenIdentities.Clear();
            this.thisTakenFolders.Clear();

            List<HardwareBoardEntry> others = existingBoards
                .Where(entry => !string.Equals(entry.ExcelDataFile, this.thisOwnIdentity, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (HardwareBoardEntry entry in others)
            {
                this.thisTakenIdentities.Add($"{entry.HardwareName}|{entry.BoardName}");
                this.thisTakenFolders.TryAdd(BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile), $"[{entry.HardwareName}] / [{entry.BoardName}]");
            }

            // Offering the manufacturers already in use is what stops a typo ("Comodore") quietly
            // creating a second top-level folder beside the real one - the names are free text, so
            // nothing else would catch it.
            List<string> manufacturers = others
                .Select(entry => NewBoardIdentity.ExtractManufacturer(entry.ExcelDataFile))
                .Where(manufacturer => !string.IsNullOrWhiteSpace(manufacturer))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(manufacturer => manufacturer, StringComparer.OrdinalIgnoreCase)
                .ToList();

            this.ManufacturerBox.ItemsSource = manufacturers;

            this.thisKnownManufacturers = manufacturers.Select(manufacturer => new KnownName(manufacturer, manufacturer)).ToList();
            this.thisKnownHardware = others
                .Where(entry => !string.IsNullOrWhiteSpace(entry.HardwareName))
                .SelectMany(entry => new[]
                {
                    new KnownName(entry.HardwareName, entry.HardwareName),
                    new KnownName(NewBoardIdentity.ExtractHardware(entry.ExcelDataFile), entry.HardwareName),
                })
                .Distinct()
                .ToList();

            this.ShowOffers();
        }

        // What the fields hold to begin with - a board's names as they are.
        internal void Use(string manufacturer, string hardware, string board, string notes)
        {
            this.ManufacturerBox.Text = manufacturer;
            this.HardwareBox.Text = hardware;
            this.BoardBox.Text = board;
            this.NotesBox.Text = notes;
            this.ShowOffers();
        }

        // ###########################################################################################
        // A board whose name cannot be changed here: shown, not typed in. Its notes go too - they are
        // sent with a new board's submission, and a published board has none here to show.
        // ###########################################################################################
        internal void ShowReadOnly(bool readOnly, bool withNotes)
        {
            this.ManufacturerBox.IsEnabled = !readOnly;
            this.HardwareBox.IsEnabled = !readOnly;
            this.BoardBox.IsEnabled = !readOnly;
            this.NotesBox.IsEnabled = !readOnly;
            this.NotesHeading.IsVisible = withNotes;
            this.NotesHint.IsVisible = withNotes;
            this.NotesBox.IsVisible = withNotes;

            if (readOnly)
            {
                this.ManufacturerDidYouMeanPanel.IsVisible = false;
                this.HardwareDidYouMeanPanel.IsVisible = false;
            }
        }

        // "Did you mean" under a box: shown while what it holds is close to a name CRT has.
        internal void ShowOffers()
        {
            BoardNameFields.ShowDidYouMean(this.ManufacturerDidYouMeanPanel, this.ManufacturerDidYouMeanText, this.ManufacturerDidYouMeanButton, this.ManufacturerBox.Text, this.thisKnownManufacturers);
            BoardNameFields.ShowDidYouMean(this.HardwareDidYouMeanPanel, this.HardwareDidYouMeanText, this.HardwareDidYouMeanButton, this.HardwareBox.Text, this.thisKnownHardware);
        }

        private static void ShowDidYouMean(StackPanel panel, TextBlock question, Button use, string? typed, IReadOnlyList<KnownName> known)
        {
            string? offered = SimilarNames.DidYouMean(typed, known);

            panel.IsVisible = offered is not null;
            question.Text = offered is null ? string.Empty : SimilarNames.Question(offered);
            use.Content = offered is null ? string.Empty : SimilarNames.UseButton(offered);
            use.Tag = offered;
        }

        private void UseOffered(Button use, Action<string> put)
        {
            if (use.Tag is string offered)
            {
                put(offered);
                this.OnAnyFieldChanged(use, EventArgs.Empty);
            }
        }

        internal string? ManufacturerOfferedForTests => this.ManufacturerDidYouMeanPanel.IsVisible ? this.ManufacturerDidYouMeanButton.Tag as string : null;

        internal string? HardwareOfferedForTests => this.HardwareDidYouMeanPanel.IsVisible ? this.HardwareDidYouMeanButton.Tag as string : null;

        private void OnAnyFieldChanged(object? sender, EventArgs e)
        {
            this.ShowOffers();
            this.FieldsChanged?.Invoke(this, EventArgs.Empty);
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

            // ###########################################################################################
            // *** A LISTED BOARD'S FOLDER IS TAKEN, WHATEVER THE BOARD IS CALLED (code review, 2026-11-02).
            // *** The check above compares what a board is LISTED as ("Commodore 64" / "250407 (long
            // board)") and the draft check below only asks for a marker, so Commodore / C64 / 250407 passed
            // both. A draft made there has no entry of its own - the listed board reads the folder
            // (DataManager.MergeDraftOnlyBoards) - so the published board opened EMPTY on this machine, and
            // Submit sent every row as removed under its id. Capitals ignored, as folder names are on
            // Windows and macOS.
            // ###########################################################################################
            string folder = BoardDescriptorRules.BoardIdFromExcelDataFile(excelDataFile);

            if (this.thisTakenFolders.TryGetValue(folder, out string? takenBy))
            {
                validationMessage = $"[{folder}] is already the folder of {takenBy}. Pick a different hardware or board name.";
                return null;
            }

            // A folder already holding a draft means a board was created here before and is still
            // live - creating over it would overwrite that contributor's own work. This can happen
            // when the hardware/board names differ from a previous board's only by the display
            // name, so the identity check above passes while the FOLDER still collides.
            //
            // Asked of the MARKER since Phase 6, which is what makes a folder a draft. A board
            // folder copied in by hand gets one when the board list loads (DraftFolderImport), so
            // one that was there at startup is refused here too. The board being renamed is no
            // collision with itself.
            if (!string.Equals(excelDataFile, this.thisOwnIdentity, StringComparison.OrdinalIgnoreCase) &&
                DraftBoardSource.HasDraft(DraftManager.DraftsRoot, excelDataFile))
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
    }
}
