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
    // A BOARD'S BOARD DATA (owner request, 2026-10-03: "Board data (and I should be able to do the
    // same edits)") - BETA's board in the submission's own table editor, in document mode.
    //
    // *** A SAVE GOES STRAIGHT TO BETA (owner decision, 2026-10-03: "it should go directly to the
    // next queue, 'BETA > Stable', so it can directly be tested in BETA"). *** "Save changes" asks the
    // server what the change would remove from BETA (BoardEditFlow.CheckAsync), asks the maintainer
    // for a REASON beside that list (PublishBoardChangeWindow - "do ask for a change reason when
    // clicking the 'Save changes' button, so this can go along with the change, just like any normal
    // contribution"), and sends it: the server makes it a submission from this account and approves
    // it at once. The table then shows BETA as it is now, and the BETA > Stable list is read again
    // (AfterPublished). A change the server made but could not publish waits under Contributor
    // Submissions, which the screen then opens (AfterSent) - nothing is lost.
    //
    // *** COLOURED AGAINST BETA AS OPENED. *** Every row starts white and only the maintainer's own
    // inserts, edits and deletions are coloured - the rule a new board's submission table keeps.
    // With "Compare sources" ticked it is coloured against the stable source instead
    // (BoardDetailView.Compare.cs).
    //
    // *** A TABLE OPENED IS A BOARD SENT BACK. *** The edit carries the fingerprint of BETA's board
    // the table was read from, and the server refuses it when BETA changed since - publishing it would
    // otherwise put that change back as it was.
    //
    // *** READ-ONLY FOR A BOARD THIS ACCOUNT MAY NOT CHANGE - OR NOT NOW. *** Every maintainer may
    // look at every board; only its own maintainers and the administrator may change it, and nobody
    // while it waits under BETA > Stable (owner decision, 2026-10-03: "it should simply disallow it,
    // even if this is coming from a maintainer"). The server's MayEdit says which, and for those the
    // table is the editor's read-only mode, with the server's reason above it.
    //
    // *** UNSAVED CHANGES ARE NEVER LEFT BEHIND SILENTLY. *** Choosing another board, signing out
    // and quitting CRT ask first (MayLeaveTableAsync, through TabMaintainer); switching view or
    // screen hides the table and keeps it.
    //
    // *** A TABLE BETA HAS MOVED PAST IS READ AGAIN (owner report, 2026-10-04). *** A fix approved
    // under Contributor Submissions, and promoted, went on showing here the warning it fixed until
    // CRT was restarted. The board's row now carries BETA's content hash, and a table read before
    // it moved is read again quietly (BoardDetailView.Sections.cs, CatchUpWithBetaAsync) - also when only
    // whether this account may change it moved, as a promotion does.
    // ###########################################################################################
    public partial class BoardDetailView
    {
        // What the table was opened on: BETA's rows, their fingerprint, and whether a save may be sent.
        private BoardTableAnswer? thisTable;

        // The board as it was when the table was READ - held against the board as it is now, so a
        // board BETA has moved past is read again (QueueRefreshRules.BoardTableChanged; owner report,
        // 2026-10-04). Null after a change of this table's own was published: the next reading of
        // the board is that state (CatchUpWithBetaAsync).
        private BoardOverviewEntry? thisTableReadAt;

        // Who maintained the board when the table was read, and as its detail says now (account
        // ids) - a swap of one maintainer for another keeps the count the row carries, but changes
        // whether this account may edit (code review, 2026-10-04). Null while not known.
        private IReadOnlyCollection<long>? thisTableReadAtMaintainers;
        private IReadOnlyCollection<long>? thisShownMaintainers;

        // The sheet last looked at in each board, so coming back to one opens it where it was left -
        // the submission table's own rule (TabMaintainer.Table.cs). For as long as CRT runs.
        private readonly Dictionary<string, string> thisSheetByBoard = new(StringComparer.Ordinal);

        private BoardTableEditor BoardTable => this.FindControl<BoardTableEditor>("BoardDataTable")!;

        // ###########################################################################################
        // Told once a change is in BETA - TabMaintainer reads the BETA > Stable list and the boards
        // again, so the board shows as waiting at once. Null (a test) does nothing more.
        // ###########################################################################################
        public Func<Task>? AfterPublished { get; set; }

        // ###########################################################################################
        // Told the submission a change became when the server made it but could NOT publish it -
        // TabMaintainer opens it under Contributor Submissions, where it waits. Null does nothing more.
        // ###########################################################################################
        public Func<long, Task>? AfterSent { get; set; }

        private void WireTable()
        {
            this.BoardTable.SaveRequested += async (_, _) => await this.SendTableAsync();
        }

        internal BoardTableEditor BoardTableForTests => this.BoardTable;

        // The table is open on this board's data.
        private bool HoldsTableFor(string? boardId) =>
            this.thisTable is not null && string.Equals(this.thisTable.BoardId, boardId, StringComparison.Ordinal);

        internal bool HasUnsavedTableEdits => this.thisTable is not null && this.BoardTable.HasUnsavedChanges;

        // ###########################################################################################
        // Reads BETA's data for the board on screen and opens the table on it. The answer is
        // dropped when another board was chosen while it was in flight. The caller holds the wait.
        // ###########################################################################################
        private async Task LoadTableAsync()
        {
            if (this.ShownBoard is not { } board)
                return;

            BoardTableAnswer? beta = await this.ReadTableForAsync(board);

            if (!this.IsShowing(board))
                return;

            // A stable table compared with BETA follows it (ShowTables).
            this.ShowTables(beta, betaMessage: null, stable: null);

            if (beta is not null)
                this.NoteTableReadAt(board);
        }

        // ###########################################################################################
        // BETA's table of `board` as the server answers it, or null - with why said above the table.
        // A board BETA no longer holds (pushed back out of it) is closed rather than left on screen
        // as if it were still there; any other failure leaves the table as it was. Null too for an
        // answer about a board no longer chosen, which the caller drops.
        // ###########################################################################################
        private async Task<BoardTableAnswer?> ReadTableForAsync(BoardOverviewEntry board)
        {
            if (!this.CanReadTable)
                return null;

            this.ShowTableLoadMessage(null);

            ReviewApiResult<BoardTableAnswer> result = await this.ReadTableAsync(board.BoardId);

            if (!this.IsShowing(board))
                return null;

            if (!result.IsOk)
            {
                if (result.Failure == ReviewApiFailure.NotFound)
                    this.CloseTableNow();

                // "Not in BETA" is an answer about the board, not a failure - shown in plain grey.
                this.ShowTableLoadMessage(result.Message, isError: result.Failure != ReviewApiFailure.NotFound);
                return null;
            }

            return result.Value!;
        }

        // The table just read was read off `board` as it is now - what TableIsBehind holds it against.
        private void NoteTableReadAt(BoardOverviewEntry board)
        {
            this.thisTableReadAt = board;
            this.thisTableReadAtMaintainers = this.thisShownMaintainers;
        }

        private bool IsShowing(BoardOverviewEntry board) =>
            string.Equals(this.ShownBoard?.BoardId, board.BoardId, StringComparison.Ordinal);

        private bool CanReadTable =>
            this.ReadTableOverrideForTests is not null || (this.thisClient is not null && this.thisSession is not null);

        private Task<ReviewApiResult<BoardTableAnswer>> ReadTableAsync(string boardId) =>
            this.ReadTableOverrideForTests is { } read
                ? read(boardId)
                : this.thisClient!.GetBoardTableAsync(this.thisSession!, boardId);

        // Reads a board's table without a server - for tests.
        internal Func<string, Task<ReviewApiResult<BoardTableAnswer>>>? ReadTableOverrideForTests { get; set; }

        // BETA's table opened on `table`, `message` first in the line above it - and a stable table
        // compared with BETA built again against it (ShowTables).
        private void OpenTable(BoardTableAnswer table, string? message) =>
            this.ShowTables(table, message, stable: null);

        // Holds `table` as BETA's. The sheet on screen is kept for the board it showed, so the same
        // board opened again (after a publish, or compared anew) stays on it.
        private void TakeTable(BoardTableAnswer table)
        {
            if (this.thisTable is not null && this.BoardTable.CurrentSheet?.Name is string current)
                this.thisSheetByBoard[this.thisTable.BoardId] = current;

            this.thisTable = table;
        }

        // ###########################################################################################
        // Builds BETA's table as held, against what it is compared with now, with `message` first in
        // the line above it. Only ShowTables and ApplyComparison call it - see ShowTables for why.
        // ###########################################################################################
        private void BuildTable(string? message)
        {
            if (this.thisTable is not { } table)
                return;

            this.TableBuildsForTests++;

            BoardTableEditor editor = this.BoardTable;
            this.ShowTableLoadMessage(null);

            // "Compare sources" ticked (BoardDetailView.Compare.cs): built against the stable
            // source's board, so everything BETA holds differently is marked.
            BoardTableAnswer? stable = this.StableToCompareWith(table.BoardId);
            this.thisTableComparedWith = stable;

            if (this.thisClient is ReviewApiClient client)
            {
                editor.FileSource = new PublishedTableFileSource(client, table.BetaDataUrl, this.LaunchFileAsync)
                {
                    Baseline = BoardDetailView.StableFilesBaseline(stable)
                };
            }

            // Otherwise coloured against BETA AS OPENED - every row white until the maintainer
            // changes it.
            BoardData beta = SubmissionRowsBoard.ToBoard(table.Rows);
            BoardData comparedWith = SubmissionRowsBoard.ToBoard((stable ?? table).Rows);

            editor.Clear();
            editor.IsReadOnly = !table.MayEdit;
            editor.Open(
                BoardTableDocument.Create(
                    comparedWith,
                    beta,
                    stable is null ? BoardSections.BaselineLabel : BoardTableDocument.StableSourceBaselineLabel),
                null,
                preferredSheet: this.thisSheetByBoard.TryGetValue(table.BoardId, out string? sheet) ? sheet : null);

            // What the table is - under the BETA / Stable switch, above the table (owner request,
            // 2026-10-04: it sat under the search box, and moved the table about between the two).
            // Why it cannot be changed is the panel above the views (ShowReadOnlyNotice), never
            // this line - which says what a read-only table is compared with, when it is (code
            // review, 2026-10-09).
            this.ShowTableNote(BoardSections.TableNote(table, comparedWithStable: stable is not null, said: message));
        }

        // ###########################################################################################
        // The panel saying why BETA's table cannot be changed - or nothing. It sits directly above
        // the view switch, said whichever view is open, as the queue's "why Approve is off" does
        // (owner request, 2026-10-09: "this kind of UI should be unified").
        //
        // *** FROM WHICHEVER WAS READ LAST: THE BOARD'S DETAIL OR BETA'S TABLE (code review,
        // 2026-10-09). *** Said only once the table was read, the panel was missing from a board
        // opened on another view until Board data had been visited - the same board looked
        // different by the views visited. The detail now carries the table's answer (ShowDetail);
        // a server older than that sends none, and the table still says it when read.
        // ###########################################################################################
        private void ShowReadOnlyNotice(string? reason)
        {
            if (this.FindControl<TextBlock>("ReadOnlyNoticeText") is TextBlock text)
                text.Text = reason ?? string.Empty;

            this.SetShown("ReadOnlyNotice", !string.IsNullOrWhiteSpace(reason));
        }

        // The read-only panel's words as shown, or empty - for tests.
        internal string ReadOnlyNoticeForTests =>
            this.FindControl<Border>("ReadOnlyNotice") is { IsVisible: true } && this.FindControl<TextBlock>("ReadOnlyNoticeText") is TextBlock text
                ? text.Text ?? string.Empty
                : string.Empty;

        // The table opened on `table` without asking the server - the seam the tests use. `readAt` is
        // the board as it was when the table was read; null takes the next reading as that.
        internal void OpenTableForTests(BoardTableAnswer table, BoardOverviewEntry? readAt = null)
        {
            this.OpenTable(table, message: null);
            this.thisTableReadAt = readAt;
            this.thisTableReadAtMaintainers = this.thisShownMaintainers;
        }

        // ###########################################################################################
        // "Save changes": the table's change, PUBLISHED TO BETA after its check and its reason (the
        // header says why each). Returns whether the change has left the table - published, or made
        // into a submission that waits - for the leaving prompt's Save.
        // ###########################################################################################
        internal async Task<bool> SendTableAsync()
        {
            if (this.thisTable is not { MayEdit: true } table ||
                (this.SendOverrideForTests is null && (this.thisClient is null || this.thisSession is null)))
            {
                return false;
            }

            BoardTableDocument? document = this.BoardTable.CommitAndGetDocument();

            if (document is null || !document.HasUnsavedChanges)
                return true;

            var request = new BoardEditRequest(table.BoardId, table.Fingerprint, null, BoardSections.RowsToSend(document, table.Rows));

            // ---- What it would remove - asked first, so a change that cannot go is refused before
            // the maintainer is asked for a reason. Nothing is made by asking.
            ReviewApiResult<BoardEditCheckAnswer> check = this.CheckOverrideForTests is { } checkOverride
                ? await checkOverride(request)
                : await ServerWait.CallAsync(
                    this, BoardSections.CheckingWait, token => this.thisClient!.CheckBoardEditAsync(this.thisSession!, request, token));

            if (!check.IsOk)
            {
                this.ShowTableNote(BoardSections.NotPublished(check.Message));
                return false;
            }

            IReadOnlyList<string> removals = check.Value!.Removals;

            // ---- The reason, beside what it removes. Cancel keeps the change in the table.
            string? reason = await this.AskReasonAsync(table.BoardId, removals);

            if (reason is null)
                return false;

            // Kept until the change has left the table, so a refused publish asks again with it
            // filled in (code review, 2026-10-04: it was typed again after every refusal).
            this.thisUnsentReason = (table.BoardId, reason);

            request = request with { Summary = reason, ExpectedRemovals = removals };

            // ---- Published, and the table read again, as one wait.
            ReviewApiResult<BoardEditResult>? result = null;
            string? publishedLine = null;

            await BusyOverlay.HoldAsync(this, BoardSections.PublishingWait, async () =>
            {
                result = this.SendOverrideForTests is { } send
                    ? await send(request)
                    : await ServerWait.CallAsync(
                        this, BoardSections.PublishingWait, token => this.thisClient!.SendBoardEditAsync(this.thisSession!, request, token));

                if (result.IsOk && result.Value!.Published)
                {
                    publishedLine = BoardSections.Published(result.Value);
                    await this.ShowBetaAfterPublishAsync(table.BoardId, publishedLine);
                }
                else if (result.IsOk)
                {
                    await this.ReopenAfterSentAsync(table, BoardSections.SavedNotPublished(result.Value!));
                }
            });

            // ###########################################################################################
            // *** NO ANSWER IN TWO MINUTES: THE BOARD'S SUBMISSIONS SAY WHAT HAPPENED. *** A change that
            // arrived is the newest submission of this board with this reason - in BETA, or waiting in
            // the queue. Only one that arrived lets the table go: otherwise the change is still only here.
            // ###########################################################################################
            if (result!.Failure == ReviewApiFailure.TimedOut)
                return await this.LookAgainAfterTimeoutAsync(table, reason);

            if (!result.IsOk)
            {
                this.ShowTableNote(BoardSections.NotPublished(result.Message));
                return false;
            }

            this.thisUnsentReason = null;

            if (publishedLine is not null)
            {
                if (this.AfterPublished is { } afterPublished)
                    await afterPublished();

                return true;
            }

            // Made into a submission, not published: the table went back to BETA as it is - the
            // change now lives in its submission (ReopenAfterSentAsync) - and that submission is
            // opened where it waits.
            if (this.AfterSent is { } afterSent)
                await afterSent(result.Value!.SubmissionId);

            return true;
        }

        // ###########################################################################################
        // *** A CHANGE MADE INTO A SUBMISSION THAT WAITS: THE TABLE IS READ AGAIN (code review,
        // 2026-10-04). *** The server refuses another change of this board while that submission
        // waits, so the table opens on the server's own answer - read-only, with its reason. Reopening
        // the answer the table was read from left it editable, a second change was refused only at
        // Save, and nothing on the board's row (BoardTableChanged) moves to catch it. Not read
        // again, it opens as it was read but not to be edited, saying why.
        // ###########################################################################################
        private async Task ReopenAfterSentAsync(BoardTableAnswer table, string line)
        {
            ReviewApiResult<BoardTableAnswer> fresh = this.CanReadTable
                ? await this.ReadTableAsync(table.BoardId)
                : ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, WaitWording.NoAnswer);

            if (!string.Equals(this.ShownBoard?.BoardId, table.BoardId, StringComparison.Ordinal))
                return;

            if (fresh.IsOk)
                this.OpenTable(fresh.Value!, line);
            else
                this.OpenTable(table with { MayEdit = false, MayNotEditReason = null }, line);

            // Taken as the state read at the board's next reading, as after a publish.
            this.thisTableReadAt = null;
            this.thisTableReadAtMaintainers = null;
        }

        // ###########################################################################################
        // After a publish: BETA's board as it is now - read-only with the server's reason while the
        // board waits under BETA > Stable - with what the publish did said first. The Files view is
        // read again when next shown, since the publish can have removed files. A table that cannot
        // be read again is CLOSED rather than left on the board as it was before the change, which
        // would look editable and be refused.
        // ###########################################################################################
        private async Task ShowBetaAfterPublishAsync(string boardId, string publishedLine)
        {
            this.ClearFiles();

            ReviewApiResult<BoardTableAnswer> fresh = this.CanReadTable
                ? await this.ReadTableAsync(boardId)
                : ReviewApiResult<BoardTableAnswer>.Failed(ReviewApiFailure.Unreachable, WaitWording.NoAnswer);

            if (!string.Equals(this.ShownBoard?.BoardId, boardId, StringComparison.Ordinal))
                return;

            if (fresh.IsOk)
            {
                this.OpenTable(fresh.Value!, publishedLine);

                // BETA as this publish left it, which the board's next reading will say - taken as
                // the state read, rather than read again over the line saying what the publish did.
                this.thisTableReadAt = null;
                this.thisTableReadAtMaintainers = null;
                return;
            }

            this.CloseTableNow();
            this.ShowTableLoadMessage($"{publishedLine} {BoardSections.NotReadAgain(fresh.Message)}", isError: false);
        }

        // ###########################################################################################
        // The two-minute limit passed with no answer: the board's detail is read, and its newest
        // submission with this reason says how far the change got (BoardSections.FindSent).
        // ###########################################################################################
        private async Task<bool> LookAgainAfterTimeoutAsync(BoardTableAnswer table, string reason)
        {
            BoardSubmissionEntry? found = null;
            bool read = false;

            if (this.thisClient is ReviewApiClient client && this.thisSession is ReviewSession session)
            {
                ReviewApiResult<BoardDetailAnswer> detail = await ServerWait.CallAsync(
                    this, WaitWording.Checking, token => client.GetBoardDetailAsync(session, table.BoardId, token));

                read = detail.IsOk;
                found = read ? BoardSections.FindSent(detail.Value!, reason) : null;
            }

            string line = MaintainerWaitWording.BoardEditAfterTimeout(read, found);

            if (found is null)
            {
                this.ShowTableNote(line);
                return false;
            }

            this.thisUnsentReason = null;

            if (BoardSections.ReachedBeta(found))
            {
                await BusyOverlay.HoldAsync(this, BoardSections.PublishingWait, () => this.ShowBetaAfterPublishAsync(table.BoardId, line));

                if (this.AfterPublished is { } afterPublished)
                    await afterPublished();

                return true;
            }

            await BusyOverlay.HoldAsync(this, BoardSections.PublishingWait, () => this.ReopenAfterSentAsync(table, line));

            if (this.AfterSent is { } afterSent)
                await afterSent(found.Id);

            return true;
        }

        // ###########################################################################################
        // The reason for the change, from PublishBoardChangeWindow - null for Cancel. With nobody to
        // ask (the tab not on screen) the answer is Cancel: the change stays in the table.
        // ###########################################################################################
        private async Task<string?> AskReasonAsync(string boardId, IReadOnlyList<string> removals)
        {
            // The reason typed for this board's change last time, when that publish was refused.
            string? typed = this.thisUnsentReason is { } unsent && string.Equals(unsent.BoardId, boardId, StringComparison.Ordinal)
                ? unsent.Reason
                : null;

            this.ReasonOfferedForTests = typed;

            if (this.ReasonAnswerForTests is { } answer)
                return answer(removals);

            if (TopLevel.GetTopLevel(this) is not Window owner)
                return null;

            var dialog = new PublishBoardChangeWindow();
            dialog.Initialize(boardId, removals, typed);

            return await dialog.ShowDialog<string?>(owner);
        }

        // The reason typed for a change whose publish was refused, by board - offered again by the
        // next Save. Forgotten once the change leaves the table, or the table closes.
        private (string BoardId, string Reason)? thisUnsentReason;

        // Answers the reason dialog in a headless test, where a dialog cannot be answered: given what
        // the check said the change removes, the reason typed - or null for Cancel.
        internal Func<IReadOnlyList<string>, string?>? ReasonAnswerForTests { get; set; }

        // The reason the dialog was last opened with - for tests.
        internal string? ReasonOfferedForTests { get; private set; }

        // Checks a change without a server - for tests.
        internal Func<BoardEditRequest, Task<ReviewApiResult<BoardEditCheckAnswer>>>? CheckOverrideForTests { get; set; }

        // Sends a change without a server - for tests.
        internal Func<BoardEditRequest, Task<ReviewApiResult<BoardEditResult>>>? SendOverrideForTests { get; set; }

        // ###########################################################################################
        // Save, Discard or Cancel - the submission table's prompt, worded for a change that goes
        // straight to BETA. Save asks for the reason as the button does. True when the table may go.
        // With nobody to ask (the tab not on screen) the answer is Cancel: the table stays, which is
        // never the harmful choice.
        //
        // `canSend` false - the tab says "CRT has to be updated" (code review, 2026-10-09) - offers
        // Discard or Cancel only: a save would be turned away out of sight.
        // ###########################################################################################
        internal async Task<bool> MayLeaveTableAsync(Window? owner = null, bool canSend = true)
        {
            if (!this.HasUnsavedTableEdits)
                return true;

            UnsavedTableEditsPrompt asked = canSend
                ? UnsavedTableEditsPrompt.LeavingBoard
                : UnsavedTableEditsPrompt.LeavingUpdateRequired;

            UnsavedTableEditsChoice? choice;

            if (this.UnsavedTableEditsAnswerForTests is { } answer)
            {
                choice = answer(asked);
            }
            else
            {
                if ((owner ?? TopLevel.GetTopLevel(this) as Window) is not Window ownerWindow)
                    return false;

                var prompt = new UnsavedTableEditsWindow();
                prompt.Initialize(asked);

                choice = await prompt.ShowDialog<UnsavedTableEditsChoice?>(ownerWindow);
            }

            return choice switch
            {
                UnsavedTableEditsChoice.Save when canSend => await this.SendTableAsync(),
                UnsavedTableEditsChoice.Discard => true,
                _ => false
            };
        }

        // Answers the unsaved-changes prompt in a headless test, where a dialog cannot be answered.
        internal Func<UnsavedTableEditsPrompt, UnsavedTableEditsChoice>? UnsavedTableEditsAnswerForTests { get; set; }

        // Closes the table WITHOUT asking - another board chosen (asked about already), or signed out.
        private void CloseTableNow()
        {
            if (this.thisTable is not null && this.BoardTable.CurrentSheet?.Name is string sheet)
                this.thisSheetByBoard[this.thisTable.BoardId] = sheet;

            this.thisTable = null;
            this.thisTableReadAt = null;
            this.thisTableReadAtMaintainers = null;
            this.thisTableComparedWith = null;
            this.thisUnsentReason = null;

            BoardTableEditor editor = this.BoardTable;
            editor.Clear();
            editor.FileSource = null;
            editor.IsReadOnly = false;

            this.ShowTableLoadMessage(null);
            this.ShowReadOnlyNotice(null);
        }

        // Above the table: why it could not be read - or, for a board BETA holds nothing of, so.
        private void ShowTableLoadMessage(string? message, bool isError = true) =>
            WindowMessage.Show(this.FindControl<TextBlock>("TableLoadMessageText"), message, isError);

        // Above the table, under the BETA / Stable switch: what the table is, and what a save did -
        // the same line, in its ordinary colour (owner request, 2026-10-04).
        private void ShowTableNote(string? message) => this.ShowTableLoadMessage(message, isError: false);

        // The line above BETA's table as shown, or empty - for tests.
        internal string TableNoteForTests =>
            this.FindControl<TextBlock>("TableLoadMessageText") is { IsVisible: true } line ? TabMaintainer.TextOf(line) : string.Empty;

        // Hands a file the preview saved to the operating system - a PDF opens in the PDF viewer.
        private Task<bool> LaunchFileAsync(string fullPath) => MaintainerFileLauncher.LaunchAsync(this, fullPath);

        // The filter the maintainer last picked in a table - handed over by TabMaintainer with the
        // submission table's, so both tables open on the same pick.
        internal BoardTableEditor TableEditorForRememberedChoices => this.BoardTable;
    }
}
