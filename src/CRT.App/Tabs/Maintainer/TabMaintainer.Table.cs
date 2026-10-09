using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Platform.Storage;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // THE SUBMISSION'S TABLE, the whole of the submission panel under a short header. It was a
    // window of its own at first (ReviewTableWindow, 2026-09-25), then took the change summary's
    // place in the panel behind a "View in table format" button (2026-09-26: "could it instead open
    // in the existing right-side panel, just alike it does in the CRT app?"), and then the summary
    // was dropped and the table opens with the submission (2026-09-26: "make the table the default
    // first view, as this is the most helpful one").
    //
    // *** THE LOGIC IS ELSEWHERE, AS EVERYWHERE IN THIS TAB. *** The table is the shared
    // BoardTableEditor in document mode, its rules are CRT.Data's (BoardTableDocument), what a save
    // sends is ReviewTableWording.RowsToSave, and what may change - who, and when - is the
    // server's (AmendSubmissionFlow). Resting on a file cell shows the file (ReviewTableFileSource).
    //
    // *** A TABLE OPENED IS A VERSION SENT. *** A save carries the amendment version the table was
    // opened at, and the server refuses it when somebody else changed the submission since.
    //
    // *** UNSAVED CHANGES ARE NEVER LEFT BEHIND SILENTLY. *** The modal window guaranteed that by
    // being modal. In the panel, every way of leaving the table asks first (Save, Discard, Cancel):
    // choosing another submission, signing out, quitting CRT. A decision waits
    // until the table is saved - it would decide the SAVED version, not the one on screen. Only a
    // submission leaving the queue, or the session ending, closes it without asking, since nothing
    // could be saved then anyway.
    // ###########################################################################################
    public partial class TabMaintainer
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
        // menu should remember the tab last visited" - then "per board, so if I am in 'Important
        // signals' in one board, then I can navigate to another board, and then it will show the
        // last sheet/tab for that board"). Each queue entry is its own, as the project owner sees
        // them; one not opened yet starts on its first sheet with a change. Unless the picked
        // colour-key pills hide the remembered sheet's tab - see BoardTableDocument.SheetToShow. For as long
        // as the application runs.
        // ###########################################################################################
        private readonly Dictionary<long, string> thisSheetBySubmission = [];

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
        }

        // ###########################################################################################
        // The table's filter - the colour-key pills picked (owner request, 2026-10-02; it was the
        // "Show changes only" check box until then) - as the maintainer last picked it, and every
        // later pick handed to `remember`: the PICK (FilterWantedChanged), never the filter a
        // particular table happens to turn on or off for itself (2026-09-29; CRT Maintainer kept it
        // in a file of its own until then).
        //
        // *** CALLED BY Main, NOT BY THE CONSTRUCTOR. *** Main passes UserSettings; a tab a test
        // builds on its own never reads or writes the shared settings - otherwise one test ticking
        // the box would leave the next test's table filtered (seen on the move: a remembered sheet
        // then had no tab and the table opened elsewhere). The same split the separate application
        // kept with its window's UseSettings, which only its App called.
        // ###########################################################################################
        internal void UseRememberedChoices(BoardTableRowKinds filter, Action<BoardTableRowKinds> remember)
        {
            ArgumentNullException.ThrowIfNull(remember);

            // A board's table on the Boards screen (2026-10-03) opens on the same pick, and its own
            // picks are remembered the same way - one table, as the maintainer sees it, in two places.
            // So a pick in either is handed to the other too (code review, 2026-10-04): each wrote
            // the one setting without telling the other, and the next launch opened both on
            // whichever was picked last.
            //
            // The stable source's read-only table (2026-10-04) is one more place of the same table.
            BoardTableEditor[] tables = [.. this.TableEditorsForSharedChoices];

            foreach (BoardTableEditor table in tables)
            {
                table.Filter = filter;

                table.FilterWantedChanged += (_, _) =>
                {
                    remember(table.FilterWanted);

                    foreach (BoardTableEditor other in tables)
                    {
                        if (!ReferenceEquals(other, table))
                            other.UseFilterWanted(table.FilterWanted);
                    }
                };
            }
        }

        // ###########################################################################################
        // The Boards screen's "Compare sources" as last ticked, and every later tick handed to
        // `remember` (owner request, 2026-10-09: "do keep the checkbox as last state"). Called by
        // Main, like UseRememberedChoices - a tab a test builds never touches UserSettings.
        // ###########################################################################################
        internal void UseRememberedComparison(bool compare, Action<bool> remember) =>
            this.BoardDetail.UseRememberedComparison(compare, remember);

        // The tab's tables - a submission's, a board's, and a board's stable source's - which share
        // the picked pills (UseRememberedChoices).
        internal IReadOnlyList<BoardTableEditor> TableEditorsForSharedChoices =>
            [this.TableEditor, this.BoardDetail.TableEditorForRememberedChoices, this.BoardDetail.StableTableEditorForSharedChoices];

        // ###########################################################################################
        // Opens the selected submission's table - straight away, as the panel's only view. Nothing
        // to do when it is the submission already open.
        // ###########################################################################################
        private async Task OpenTableAsync(ReviewQueueRow row)
        {
            if (this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session ||
                this.thisTableRow?.Id == row.Id)
            {
                return;
            }

            this.BeginTable(row, client, session);
            this.ShowTableLoadMessage("Loading the table...", isError: false);
            await this.LoadTableAsync(message: null);
        }

        // Makes `row` the table's submission, with nothing in it yet - before the rows are read, or
        // with rows read ahead of time (TabMaintainer.Prefetch.cs).
        private void BeginTable(ReviewQueueRow row, ReviewApiClient client, ReviewSession session)
        {
            this.thisTableRow = row;
            this.thisTable = null;

            this.TableEditor.FileSource = new ReviewTableFileSource(
                client,
                session,
                row.Id,
                () => this.thisShownDetail?.SubmittedFiles ?? [],
                () => this.thisTable is { Published: null },
                this.LaunchFileAsync);

            this.ShowTablePanel(open: true);
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
            // *** A NEW BOARD IS COMPARED WITH THE SUBMISSION ITSELF, AS IT WAS OPENED (owner decision,
            // 2026-09-26). *** Its rows start white, and only what the MAINTAINER inserts, changes or
            // deletes is coloured - "Only if the maintainer inserts something, it should show as
            // added (or removed or changed)" - and picking Added, Modified or Deleted shows nothing until they do.
            //
            // Compared with nothing, the table had marked nothing and hidden the colour key but a
            // lone "0 Flagged" ("where are the others?"). Compared with an EMPTY board, every row was
            // green - tried the same day and turned down: everything is new, so green says nothing.
            //
            // A save makes the edits the submission's content, so the table reopened after it is
            // white again. (The Drafts tab keeps its own "every row is your own" view of a
            // contributor's new board.)
            // ###########################################################################################
            BoardData published = SubmissionRowsBoard.ToBoard(table.Published ?? table.Submitted);
            BoardData submitted = SubmissionRowsBoard.ToBoard(table.Submitted);

            // ###########################################################################################
            // What a changed cell's tooltip calls the value it replaced. For a NEW BOARD the
            // baseline above is the SUBMISSION itself, not a published board, so the default
            // "Published value: (empty)" named something that does not exist (owner report,
            // 2026-09-26). It then says "As submitted", the file card's own word for that same side
            // (ReviewTableFiles.SideLabels), so one thing has one name whether the pointer is on a
            // text cell or a file cell.
            //
            // A PUBLISHED board is the server's BETA tree (ReviewEndpoints.GetTableAsync reads
            // DataTreeRoot), so the tooltip says "BETA source value" (owner request, 2026-10-05:
            // "Published value" read as stable, and a value only BETA held yet looked like an
            // error). The file card says "Before (BETA source)" there because it is showing two
            // pictures side by side and has to say which is which, while a tooltip names one value.
            // Same side, two sentences - which is why this picks a word rather than reusing that pair
            // wholesale.
            // ###########################################################################################
            string baselineLabel = table.Published is null
                ? ReviewTableFiles.SideLabels(nothingPublished: true).Published
                : BoardTableDocument.BetaSourceBaselineLabel;

            this.TableEditor.Open(
                BoardTableDocument.Create(published, submitted, baselineLabel),
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
            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;
            int versionBefore = this.thisTable.Version;
            SubmissionRows rows = ReviewTableWording.RowsToSave(document, this.thisTable.Submitted);

            bool saved = false;

            // The whole window waits, as for a decision (2026-09-28) - the save AND the re-read after
            // it, so nothing can be pressed against the table while it is being replaced.
            await BusyOverlay.HoldAsync(this, ReviewTableWording.SavingWait, async () =>
            {
                ReviewApiResult<ReviewAmendResult> result = await ServerWait.CallAsync(
                    this,
                    ReviewTableWording.SavingWait,
                    token => client.AmendAsync(session, row.Id, versionBefore, rows, token));

                // ###########################################################################################
                // *** NO ANSWER IN TWO MINUTES: THE VERSION SAYS WHETHER IT SAVED. *** Every saved
                // amendment moves the submission's version on, so reading the table again tells a
                // save that landed from one that did not - and only a landed one reloads the table,
                // since reloading would otherwise throw the unsaved changes away.
                // ###########################################################################################
                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    ReviewApiResult<ReviewTableData> now = await ServerWait.CallAsync(
                        this, WaitWording.Checking, token => client.GetTableAsync(session, row.Id, token));

                    int? versionNow = now.IsOk ? now.Value!.Version : null;
                    string afterwards = MaintainerWaitWording.SaveAfterTimeout(versionBefore, versionNow);

                    if (versionNow > versionBefore)
                    {
                        saved = true;
                        await this.ReadSubmissionAgainAsync(row, afterwards);
                    }
                    else
                    {
                        this.TableEditor.ShowMessage(afterwards);
                    }

                    return;
                }

                if (!result.IsOk)
                {
                    this.TableEditor.ShowMessage(ReviewTableWording.NotSaved(result.Message));
                    return;
                }

                saved = true;
                await this.ReadSubmissionAgainAsync(row, ReviewTableWording.Saved(result.Value!));
            });

            return saved;
        }

        // The table and the rest of the submission read again after a save, under the overlay the
        // save is holding.
        private Task ReadSubmissionAgainAsync(ReviewQueueRow row, string message) =>
            ServerWait.RunAsync(this, MaintainerWaitWording.ReadingSubmissionAgain, async () =>
            {
                await this.LoadTableAsync(message);
                await this.LoadSubmissionAsync(row);
            });

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
        // table may go. While the tab says "CRT has to be updated", Discard or Cancel: a save would
        // only be turned away (TabMaintainer.UpdateRequired.cs).
        private async Task<bool> MayLeaveTableAsync(Window? owner = null)
        {
            if (!this.IsTableOpen || !this.TableEditor.HasUnsavedChanges)
                return true;

            UnsavedTableEditsPrompt asked = this.thisUpdateRequired
                ? UnsavedTableEditsPrompt.LeavingUpdateRequired
                : UnsavedTableEditsPrompt.LeavingSubmission;

            UnsavedTableEditsChoice? choice;

            if (this.UnsavedTableEditsAnswerForTests is { } answer)
            {
                choice = answer(asked);
            }
            else
            {
                // Owned by CRT's window. With none (the tab not on screen) there is nobody to ask,
                // so the answer is Cancel - the table stays, which is never the harmful choice.
                if ((owner ?? TopLevel.GetTopLevel(this) as Window) is not Window ownerWindow)
                    return false;

                var prompt = new UnsavedTableEditsWindow();
                prompt.Initialize(asked);

                choice = await prompt.ShowDialog<UnsavedTableEditsChoice?>(ownerWindow);
            }

            return choice switch
            {
                UnsavedTableEditsChoice.Save when !this.thisUpdateRequired => await this.SaveTableAsync(),
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
        private Task<bool> LaunchFileAsync(string fullPath) => MaintainerFileLauncher.LaunchAsync(this, fullPath);

        // ###########################################################################################
        // *** QUITTING CRT ASKS TOO, FROM Main (2026-09-29). *** While this was its own window it
        // cancelled its own Closing; a tab has no Closing, so Main.OnWindowClosing asks these two
        // - the same pair TabDrafts hands it - and brings this tab forward first, so the prompt is
        // about a table the maintainer can see. `owner` is Main itself.
        //
        // *** BOTH TABLES (2026-10-03). *** A board's table on the Boards screen can hold a change
        // not sent too. It is asked about after the submission's, on its own screen, for the same
        // reason.
        // ###########################################################################################
        internal bool HasUnsavedTableEdits =>
            (this.IsTableOpen && this.TableEditor.HasUnsavedChanges) || this.BoardDetail.HasUnsavedTableEdits;

        internal async Task<bool> ConfirmLeavingTableAsync(Window? owner = null)
        {
            if (!await this.MayLeaveTableAsync(owner))
                return false;

            if (!this.BoardDetail.HasUnsavedTableEdits)
                return true;

            await this.ShowModeAsync(MaintainerMode.Boards);
            return await this.BoardDetail.MayLeaveTableAsync(owner, canSend: !this.thisUpdateRequired);
        }

        // Answers the unsaved-changes prompt in a headless test, where a dialog cannot be answered.
        internal Func<UnsavedTableEditsPrompt, UnsavedTableEditsChoice>? UnsavedTableEditsAnswerForTests { get; set; }
    }
}
