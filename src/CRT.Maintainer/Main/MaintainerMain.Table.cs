using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer
{
    // ###########################################################################################
    // THE SUBMISSION'S TABLE, the whole of the submission panel under a short header. It was a
    // window of its own at first (ReviewTableWindow, 2026-09-25), then took the change summary's
    // place in the panel behind a "View in table format" button (2026-09-26: "could it instead open
    // in the existing right-side panel, just alike it does in the CRT app?"), and then the summary
    // was dropped and the table opens with the submission (2026-09-26: "make the table the default
    // first view, as this is the most helpful one").
    //
    // *** THE LOGIC IS ELSEWHERE, AS EVERYWHERE IN THIS APPLICATION. *** The table is CRT.UI's
    // BoardTableEditor in document mode, its rules are CRT.Data's (BoardTableDocument), what a save
    // sends is ReviewTableWording.RowsToSave, and what may change - who, and when - is the
    // server's (AmendSubmissionFlow). Resting on a file cell shows the file (ReviewTableFileSource).
    //
    // *** A TABLE OPENED IS A VERSION SENT. *** A save carries the amendment version the table was
    // opened at, and the server refuses it when somebody else changed the submission since.
    //
    // *** UNSAVED CHANGES ARE NEVER LEFT BEHIND SILENTLY. *** The modal window guaranteed that by
    // being modal. In the panel, every way of leaving the table asks first (Save, Discard, Cancel):
    // choosing another submission, signing out, closing the window. A decision waits
    // until the table is saved - it would decide the SAVED version, not the one on screen. Only a
    // submission leaving the queue, or the session ending, closes it without asking, since nothing
    // could be saved then anyway.
    // ###########################################################################################
    public partial class MaintainerMain
    {
        // The submission whose table is open - null when none is selected - and what the table was
        // opened on: the version a save is checked against, and the rows it edits.
        private ReviewQueueRow? thisTableRow;
        private ReviewTableData? thisTable;

        // The submission detail on screen, for the file preview's submitted-file hashes.
        private ReviewSubmissionDetail? thisShownDetail;

        // ###########################################################################################
        // The sheet last looked at IN EACH SUBMISSION, so coming back to one opens it where it was
        // left (owner requests, 2026-09-26: "navigating between different systems in the left-side
        // menu should remember the tab last visited" - then "per system, so if I am in 'Important
        // signals' in one system, then I can navigate to another system, and then it will show the
        // last sheet/tab for that system"). Each queue entry is its own, as the project owner sees
        // them; one not opened yet starts on its first sheet with a change. Unless "Show changes
        // only" hides the remembered sheet's tab - see BoardTableDocument.SheetToShow. For as long
        // as the application runs.
        // ###########################################################################################
        private readonly Dictionary<long, string> thisSheetBySubmission = [];

        // Set once closing the window has been settled, so the Closing handler lets it go the second
        // time round.
        private bool thisClosingSettled;

        // ###########################################################################################
        // *** LOOKED UP BY NAME, NOT THROUGH A GENERATED FIELD. *** This window loads its markup
        // itself (AvaloniaXamlLoader.Load in InitializeComponent), so the fields Avalonia generates
        // for x:Name are never filled in - reading one here threw in the constructor, and the whole
        // application failed to start (caught by rendering it, 2026-09-26). Every other control in
        // this window is found the same way.
        // ###########################################################################################
        private BoardTableEditor TableEditor => this.FindControl<BoardTableEditor>("SubmissionTable")!;

        internal bool IsTableOpen => this.thisTableRow is not null;

        private void WireTable()
        {
            this.TableEditor.SaveRequested += async (_, _) => await this.SaveTableAsync();
            this.Closing += this.OnWindowClosing;
        }

        // ###########################################################################################
        // Opens the selected submission's table - straight away, as the panel's only view. Nothing
        // to do when it is the submission already open.
        // ###########################################################################################
        private async Task OpenTableAsync(ReviewQueueRow row)
        {
            if (this.thisClient is null || this.thisSession is null || this.thisTableRow?.Id == row.Id)
                return;

            this.thisTableRow = row;
            this.thisTable = null;

            this.TableEditor.FileSource = new ReviewTableFileSource(
                this.thisClient,
                this.thisSession,
                row.Id,
                () => this.thisShownDetail?.SubmittedFiles ?? [],
                () => this.thisTable is { Published: null },
                this.LaunchFileAsync);

            this.ShowTablePanel(open: true);
            this.ShowTableLoadMessage("Loading the table...", isError: false);
            await this.LoadTableAsync(message: null);
        }

        // ###########################################################################################
        // Fetches the submission's rows and the published board, and opens the table on them.
        // `message` is shown under the table's toolbar afterwards - "Saved - ..." after a save.
        //
        // The answer is dropped if the table was closed, or moved to another submission, while it
        // was in flight - the same rule the summary follows.
        // ###########################################################################################
        private async Task LoadTableAsync(string? message)
        {
            if (this.thisClient is null || this.thisSession is null || this.thisTableRow is null)
                return;

            long id = this.thisTableRow.Id;

            ReviewApiResult<ReviewTableData> result = await this.thisClient.GetTableAsync(this.thisSession, id);

            if (this.thisTableRow?.Id != id)
                return;

            if (!result.IsOk)
            {
                this.ShowTableLoadMessage(result.Message);
                return;
            }

            this.ShowTable(result.Value!, message);
        }

        private void ShowTable(ReviewTableData table, string? message)
        {
            this.thisTable = table;
            this.ShowTableLoadMessage(null);

            // ###########################################################################################
            // *** A NEW SYSTEM IS COMPARED WITH THE SUBMISSION ITSELF, AS IT WAS OPENED (owner decision,
            // 2026-09-26). *** Its rows start white, and only what the MAINTAINER inserts, changes or
            // deletes is coloured - "Only if the maintainer inserts something, it should show as
            // added (or removed or changed)" - and "Show changes only" shows nothing until they do.
            //
            // Compared with nothing, the table had marked nothing and hidden the colour key but a
            // lone "0 Flagged" ("where are the others?"). Compared with an EMPTY board, every row was
            // green - tried the same day and turned down: everything is new, so green says nothing.
            //
            // A save makes the edits the submission's content, so the table reopened after it is
            // white again. (The Drafts tab keeps its own "every row is your own" view of a
            // contributor's new system.)
            // ###########################################################################################
            BoardData published = SubmissionRowsBoard.ToBoard(table.Published ?? table.Submitted);
            BoardData submitted = SubmissionRowsBoard.ToBoard(table.Submitted);

            this.TableEditor.Open(
                BoardTableDocument.Create(published, submitted),
                message ?? ReviewTableWording.OpenedMessage(table),
                preferredSheet: this.thisTableRow is { } row && this.thisSheetBySubmission.TryGetValue(row.Id, out string? sheet) ? sheet : null);
        }

        // The table opened on `table` without asking the server - the seam the tests use, since
        // OpenTableAsync fetches it.
        internal void OpenTableForTests(ReviewQueueRow row, ReviewTableData table)
        {
            this.thisTableRow = row;
            this.ShowTablePanel(open: true);
            this.ShowTable(table, message: null);
        }

        internal BoardTableEditor TableEditorForTests => this.TableEditor;

        // ###########################################################################################
        // "Save changes": the table's rows, sent as an amendment at the version it was opened at.
        // The submission is then read again - the table AND the summary, the files it removes and who
        // must still approve, all of which follow the changed content and which the decision bar
        // beside the table acts on. Returns whether it was saved, for the leaving prompt's Save.
        // ###########################################################################################
        private async Task<bool> SaveTableAsync()
        {
            if (this.thisClient is null || this.thisSession is null || this.thisTableRow is null || this.thisTable is null)
                return false;

            BoardTableDocument? document = this.TableEditor.CommitAndGetDocument();

            if (document is null || !document.HasUnsavedChanges)
                return true;

            ReviewQueueRow row = this.thisTableRow;

            this.TableEditor.ShowMessage("Saving...");

            ReviewApiResult<ReviewAmendResult> result = await this.thisClient.AmendAsync(
                this.thisSession,
                row.Id,
                this.thisTable.Version,
                ReviewTableWording.RowsToSave(document, this.thisTable.Submitted));

            if (!result.IsOk)
            {
                this.TableEditor.ShowMessage(ReviewTableWording.NotSaved(result.Message));
                return false;
            }

            await this.LoadTableAsync(ReviewTableWording.Saved(result.Value!));
            await this.LoadSubmissionAsync(row);

            return true;
        }

        // ###########################################################################################
        // Closes the table - on leaving its submission - asking first when it holds unsaved changes.
        // False when the maintainer chose to stay (Cancel), or chose Save and it was refused - the
        // table stays open with the reason on screen.
        // ###########################################################################################
        internal async Task<bool> CloseTableAsync()
        {
            if (!this.IsTableOpen)
                return true;

            if (!await this.MayLeaveTableAsync())
                return false;

            this.CloseTableNow();
            return true;
        }

        // Save, Discard or Cancel - the Drafts tab's prompt, worded for a submission. True when the
        // table may go.
        private async Task<bool> MayLeaveTableAsync()
        {
            if (!this.IsTableOpen || !this.TableEditor.HasUnsavedChanges)
                return true;

            var prompt = new UnsavedTableEditsWindow();
            prompt.Initialize(UnsavedTableEditsPrompt.LeavingSubmission);

            UnsavedTableEditsChoice? choice = await prompt.ShowDialog<UnsavedTableEditsChoice?>(this);

            return choice switch
            {
                UnsavedTableEditsChoice.Save => await this.SaveTableAsync(),
                UnsavedTableEditsChoice.Discard => true,
                _ => false
            };
        }

        // Closes it WITHOUT asking - only where nothing could be saved any more (see the header).
        private void CloseTableNow()
        {
            if (!this.IsTableOpen)
                return;

            if (this.TableEditor.CurrentSheet?.Name is string sheet)
                this.thisSheetBySubmission[this.thisTableRow!.Id] = sheet;

            this.thisTableRow = null;
            this.thisTable = null;

            this.TableEditor.Clear();
            this.TableEditor.FileSource = null;
            this.ShowTableLoadMessage(null);
            this.ShowTablePanel(open: false);
        }

        // ###########################################################################################
        // Puts the queue's selection back on the table's submission, after the maintainer chose to
        // stay in the table rather than move to another one. Suppressed, so it does not count as a
        // new choice.
        // ###########################################################################################
        private void ReselectTableRow()
        {
            if (this.thisTableRow is null)
                return;

            this.thisSuppressQueueSelection = true;

            try
            {
                this.SelectQueueRow(this.thisTableRow.Id);
            }
            finally
            {
                this.thisSuppressQueueSelection = false;
            }
        }

        private void ShowTablePanel(bool open)
        {
            if (this.FindControl<DockPanel>("TablePanel") is DockPanel table)
                table.IsVisible = open;
        }

        // Under the header: "Loading the table...", or why it could not be read.
        private void ShowTableLoadMessage(string? message, bool isError = true)
        {
            if (this.FindControl<TextBlock>("TableLoadMessageText") is not TextBlock text)
                return;

            text.Text = message ?? string.Empty;
            text.IsVisible = !string.IsNullOrWhiteSpace(message);
            text.Foreground = isError ? Avalonia.Media.Brushes.IndianRed : Avalonia.Media.Brushes.Gray;
        }

        // Hands a file the preview saved to the operating system - a PDF opens in the PDF viewer.
        private async Task<bool> LaunchFileAsync(string fullPath)
        {
            try
            {
                return await this.Launcher.LaunchFileInfoAsync(new FileInfo(fullPath));
            }
            catch (Exception)
            {
                // A platform with no handler for the type; the preview says it could not open.
                return false;
            }
        }

        // ###########################################################################################
        // Closing the window with unsaved table changes asks first, like every other way out.
        // ###########################################################################################
        private async void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            if (this.thisClosingSettled || !this.IsTableOpen || !this.TableEditor.HasUnsavedChanges)
                return;

            e.Cancel = true;

            if (await this.MayLeaveTableAsync())
            {
                this.thisClosingSettled = true;
                this.Close();
            }
        }
    }
}
