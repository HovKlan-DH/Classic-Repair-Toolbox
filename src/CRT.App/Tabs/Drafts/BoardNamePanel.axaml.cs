using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "BOARD NAME" ON THE RIGHT OF THE CONTRIBUTE TAB (owner request, 2026-11-01: "how to edit those
    // details again afterwards - I think this should be the very first item in the list in 'Board
    // contents'"): "Add a new board"'s fields (BoardNameFields) with what the board has in them, and
    // one button. The move itself is the host's (TabDrafts.BoardName.cs).
    //
    // A board whose name cannot be changed - published, or sent for review under it - shows its name
    // and why, with nothing to press (BoardNameWording).
    // ###########################################################################################
    public partial class BoardNamePanel : UserControl
    {
        // What the board has now - Save is on only for something else.
        private string thisManufacturer = string.Empty;
        private string thisHardware = string.Empty;
        private string thisBoard = string.Empty;
        private string thisNotes = string.Empty;
        private bool thisReadOnly;

        // Saves what the fields describe; what stands in its way when it could not, null when done.
        internal Func<NewBoardRegistration, Task<string?>>? SaveRequested { get; set; }

        public BoardNamePanel()
        {
            this.InitializeComponent();
            this.Fields.FieldsChanged += (_, _) => this.ShowSave();
            this.MovesText.Text = BoardNameWording.RenameMoves;
        }

        // ###########################################################################################
        // The page for a board - `whyNot`: why its name cannot be changed here, null when it can.
        // `existing`: every board, for the names already taken.
        // ###########################################################################################
        internal void Show(
            string identity,
            string hardware,
            string board,
            string notes,
            IReadOnlyList<HardwareBoardEntry> existing,
            string? whyNot,
            bool isNewBoard)
        {
            this.thisManufacturer = NewBoardIdentity.ExtractManufacturer(identity);
            this.thisHardware = hardware;
            this.thisBoard = board;
            this.thisNotes = notes;
            this.thisReadOnly = whyNot is not null;

            this.IntroText.Text = BoardNameWording.Intro;
            this.Fields.Initialize(existing, identity);
            this.Fields.Use(this.thisManufacturer, hardware, board, notes);
            this.Fields.ShowReadOnly(this.thisReadOnly, withNotes: isNewBoard);

            this.WhyNotPanel.IsVisible = whyNot is not null;
            this.WhyNotText.Text = whyNot;
            this.SavePanel.IsVisible = !this.thisReadOnly;
            this.MovesText.IsVisible = !this.thisReadOnly;

            this.ShowSaved(null);
            this.ShowSave();
        }

        // Says what came of a save, beside its button.
        internal void ShowSaved(string? text)
        {
            this.SavedText.Text = text;
            this.SavedText.IsVisible = !string.IsNullOrEmpty(text);
        }

        // ###########################################################################################
        // Save is on for a name CRT can take that is not the board's already (BoardNameFields' rules).
        // ###########################################################################################
        private void ShowSave()
        {
            if (this.thisReadOnly)
            {
                this.ValidationText.IsVisible = false;
                return;
            }

            NewBoardRegistration? wanted = this.Fields.BuildRegistrationOrNull(out string message);

            this.SaveButton.IsEnabled = wanted is not null && this.IsAChange(wanted);
            this.ValidationText.Text = message;
            this.ValidationText.IsVisible = message.Length > 0;
        }

        private bool IsAChange(NewBoardRegistration wanted) =>
            !string.Equals(NewBoardIdentity.ExtractManufacturer(wanted.ExcelDataFile), this.thisManufacturer, StringComparison.Ordinal) ||
            !string.Equals(wanted.HardwareName, this.thisHardware, StringComparison.Ordinal) ||
            !string.Equals(wanted.BoardName, this.thisBoard, StringComparison.Ordinal) ||
            !string.Equals(wanted.HardwareNotes, this.thisNotes.Trim(), StringComparison.Ordinal);

        internal bool CanSaveForTests => this.SaveButton.IsEnabled;

        internal void RefreshForTests() => this.ShowSave();

        private async void OnSaveClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.SaveAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Saving the board's name failed - [{ex}]");
            }
        }

        internal async Task SaveAsync()
        {
            if (this.Fields.BuildRegistrationOrNull(out _) is not { } wanted || this.SaveRequested is not { } save)
            {
                return;
            }

            string? problem = await save(wanted);

            this.ValidationText.Text = problem;
            this.ValidationText.IsVisible = !string.IsNullOrEmpty(problem);
        }
    }
}
