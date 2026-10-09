using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE BETA / STABLE SWITCH ON BOARD DATA AND FILES (owner request, 2026-10-04: "Shouldn't there be
    // somewhere a possibility to see what we actually do have in BETA or stable for this?").
    //
    // BETA's table and file tree are BoardDetailView.Table.cs' and .Files.cs' - the table the one a change
    // is made in. The stable source has a table and a file tree OF ITS OWN beside them, here: the
    // table read-only (the server says why - only a publish from BETA changes it), the tree opening
    // each file from the stable source's public address. So switching hides one and shows the other,
    // and a change in BETA's table is never touched - the tab's own "hide, never close" rule.
    //
    // WHICH IS SHOWN is BoardSections.ShowsStable: the one wanted, when the board is there; else
    // the stable source when only it holds the board; else BETA, which then says nothing is there.
    // The one wanted stays for the next board, as the chosen view does.
    //
    // READ when first shown for a board, like BETA's; read again when the stable source moved under
    // it - only a promotion moves it (QueueRefreshRules.StableBoardChanged) - quietly when on screen,
    // and otherwise dropped, to be read when next shown.
    // ###########################################################################################
    public partial class BoardDetailView
    {
        // The maintainer's pick: the stable source, or BETA (the default).
        private bool thisStableWanted;

        // The stable source's board, and the board as it was when it was read.
        private BoardTableAnswer? thisStableTable;
        private BoardOverviewEntry? thisStableTableReadAt;

        // The board the stable files are held for, and as it was when they were read.
        private string? thisStableFilesFor;
        private BoardOverviewEntry? thisStableFilesReadAt;

        // The sheet last looked at in each board's stable table, for as long as CRT runs.
        private readonly Dictionary<string, string> thisStableSheetByBoard = new(StringComparer.Ordinal);

        private BoardTableEditor StableBoardTable => this.FindControl<BoardTableEditor>("StableTable")!;

        private FileTreeView StableFiles => this.FindControl<FileTreeView>("StableFileTree")!;

        // Whether the stable source's half is the one on screen for the board shown.
        private bool ShowsStable => BoardSections.ShowsStable(this.thisStableWanted, this.ShownBoard);

        internal bool ShowsStableForTests => this.ShowsStable;

        internal BoardTableEditor StableTableForTests => this.StableBoardTable;

        internal FileTreeView StableFileTreeForTests => this.StableFiles;

        // The stable table, for the picked pills the Maintainer tab's tables share
        // (TabMaintainer.UseRememberedChoices).
        internal BoardTableEditor StableTableEditorForSharedChoices => this.StableBoardTable;

        private async void OnTreeClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button button)
                return;

            await this.ShowTreeAsync(string.Equals(button.Name, "StableTreeButton", StringComparison.Ordinal));
        }

        // ###########################################################################################
        // The maintainer's pick, then the view on show read when it needs the server - under the
        // window's "please wait", since the maintainer pressed for it.
        // ###########################################################################################
        internal async Task ShowTreeAsync(bool stable)
        {
            this.thisStableWanted = stable;
            this.ApplyTree();

            if (this.ShownBoard is not { } board || !this.NeedsReading(this.thisSection))
                return;

            string wait = this.thisSection == BoardSection.Files
                ? MaintainerWaitWording.ReadingBoardFiles(board.BoardId)
                : MaintainerWaitWording.ReadingBoardTable(board.BoardId);

            if (!await ServerWait.RunAsync(this, wait, this.LoadShownSectionAsync))
                this.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        // ###########################################################################################
        // The switch and the two halves, from the pick and the board on screen: the switch shows
        // with Board data and Files only, a half the board is not in cannot be chosen, and the half
        // shown is the one ShowsStable says.
        // ###########################################################################################
        private void ApplyTree()
        {
            BoardOverviewEntry? board = this.ShownBoard;
            bool stable = this.ShowsStable;

            this.SetShown(
                "TreeSwitchBar",
                board is not null && this.thisSection is BoardSection.BoardData or BoardSection.Files);

            if (this.FindControl<Button>("BetaTreeButton") is Button beta)
            {
                beta.Classes.Set("Selected", !stable);
                beta.IsEnabled = BoardSections.CanChooseBeta(board);
            }

            if (this.FindControl<Button>("StableTreeButton") is Button stableButton)
            {
                stableButton.Classes.Set("Selected", stable);
                stableButton.IsEnabled = BoardSections.CanChooseStable(board);
            }

            this.SetShown("BetaBoardPart", !stable);
            this.SetShown("StableBoardPart", stable);
            this.SetShown("BetaFilesPart", !stable);
            this.SetShown("StableFilesPart", stable);

            // "Compare sources" beside the switch - with Board data only (BoardDetailView.Compare.cs).
            this.ApplyCompareBox();
        }

        private bool HoldsStableTableFor(string? boardId) =>
            this.thisStableTable is not null && string.Equals(this.thisStableTable.BoardId, boardId, StringComparison.Ordinal);

        private bool HoldsStableFilesFor(string? boardId) =>
            this.thisStableFilesFor is not null && string.Equals(this.thisStableFilesFor, boardId, StringComparison.Ordinal);

        private bool StableTableIsBehind() =>
            this.thisStableTableReadAt is { } readAt && this.ShownBoard is { } now && QueueRefreshRules.StableBoardChanged(readAt, now);

        private bool StableFilesAreBehind() =>
            this.thisStableFilesReadAt is { } readAt && this.ShownBoard is { } now && QueueRefreshRules.StableBoardChanged(readAt, now);

        // Whether the stable half of the view on show has still to be read for the board on screen.
        private bool StableNeedsReading(BoardSection section)
        {
            string? boardId = this.ShownBoard?.BoardId;

            return section switch
            {
                BoardSection.BoardData => !this.HoldsStableTableFor(boardId) || this.StableTableIsBehind(),
                BoardSection.Files => !this.HoldsStableFilesFor(boardId) || this.StableFilesAreBehind(),
                _ => false
            };
        }

        // ###########################################################################################
        // Reads the stable source's data for the board on screen and opens its read-only table. An
        // answer for a board no longer chosen is dropped. The caller holds the wait.
        // ###########################################################################################
        private async Task LoadStableTableAsync()
        {
            if (this.ShownBoard is not { } board)
                return;

            BoardTableAnswer? stable = await this.ReadStableTableForAsync(board);

            if (!this.IsShowing(board))
                return;

            // BETA's table compared with the stable source follows it - compared with nothing when
            // the stable table was closed (ShowTables).
            this.ShowTables(beta: null, betaMessage: null, stable);

            if (stable is not null)
                this.thisStableTableReadAt = board;
        }

        // ###########################################################################################
        // The stable source's table of `board` as the server answers it, or null - with why said
        // above it. A board the stable source does not hold, and an older server's answer that is
        // BETA's, close the stable table. Null too for an answer about a board no longer chosen.
        // ###########################################################################################
        private async Task<BoardTableAnswer?> ReadStableTableForAsync(BoardOverviewEntry board)
        {
            if (!this.CanReadStable)
                return null;

            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), null, isError: false);

            ReviewApiResult<BoardTableAnswer> result = await this.ReadStableTableAsync(board.BoardId);

            if (!this.IsShowing(board))
                return null;

            if (!result.IsOk)
            {
                if (result.Failure == ReviewApiFailure.NotFound)
                    this.CloseStableTable();

                // "Not in the stable source" is an answer about the board, not a failure.
                WindowMessage.Show(
                    this.FindControl<TextBlock>("StableTableMessageText"),
                    result.Message,
                    isError: result.Failure != ReviewApiFailure.NotFound);

                return null;
            }

            // An older server answered with BETA's board: never drawn as the stable source's.
            if (!BoardSections.IsStableAnswer(result.Value!))
            {
                this.CloseStableTable();
                WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), BoardSections.StableNeedsNewerServer, isError: true);
                return null;
            }

            return result.Value!;
        }

        private bool CanReadStable =>
            this.ReadStableTableOverrideForTests is not null || (this.thisClient is not null && this.thisSession is not null);

        private Task<ReviewApiResult<BoardTableAnswer>> ReadStableTableAsync(string boardId) =>
            this.ReadStableTableOverrideForTests is { } read
                ? read(boardId)
                : this.thisClient!.GetBoardTableAsync(this.thisSession!, boardId, DataTreeNames.Production);

        // Reads a board's stable table without a server - for tests.
        internal Func<string, Task<ReviewApiResult<BoardTableAnswer>>>? ReadStableTableOverrideForTests { get; set; }

        // The stable table opened on `table` - and BETA's compared with it built again (ShowTables).
        private void OpenStableTable(BoardTableAnswer table) =>
            this.ShowTables(beta: null, betaMessage: null, table);

        // Holds `table` as the stable source's, keeping the sheet on screen for the board it showed.
        private void TakeStableTable(BoardTableAnswer table)
        {
            if (this.thisStableTable is not null && this.StableBoardTable.CurrentSheet?.Name is string current)
                this.thisStableSheetByBoard[this.thisStableTable.BoardId] = current;

            this.thisStableTable = table;
        }

        // Builds the stable table as held, against what it is compared with now. Only ShowTables and
        // ApplyComparison call it.
        private void BuildStableTable()
        {
            if (this.thisStableTable is not { } table)
                return;

            this.StableTableBuildsForTests++;

            BoardTableEditor editor = this.StableBoardTable;

            // "Compare sources" ticked (BoardDetailView.Compare.cs): built against BETA's board as
            // read, so everything the stable source holds differently is marked.
            BoardTableAnswer? beta = this.BetaToCompareWith(table.BoardId);
            this.thisStableComparedWith = beta;

            if (this.thisClient is ReviewApiClient client)
            {
                editor.FileSource = new PublishedTableFileSource(client, table.ProductionDataUrl, this.LaunchFileAsync, BoardSections.StableTreeName)
                {
                    Baseline = BoardDetailView.BetaFilesBaseline(beta)
                };
            }

            BoardData stable = SubmissionRowsBoard.ToBoard(table.Rows);
            BoardData comparedWith = SubmissionRowsBoard.ToBoard((beta ?? table).Rows);

            // Never editable: the server says so, and the table is told so whatever it says.
            editor.Clear();
            editor.IsReadOnly = true;
            editor.Open(
                BoardTableDocument.Create(
                    comparedWith,
                    stable,
                    beta is null ? BoardSections.StableBaselineLabel : BoardTableDocument.BetaSourceBaselineLabel),
                null,
                preferredSheet: this.thisStableSheetByBoard.TryGetValue(table.BoardId, out string? sheet) ? sheet : null);

            // Why it cannot be changed - under the switch, where BETA's table says what it is.
            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), table.MayNotEditReason, isError: false);
        }

        // The line above the stable table as shown, or empty - for tests.
        internal string StableTableNoteForTests =>
            this.FindControl<TextBlock>("StableTableMessageText") is { IsVisible: true } line ? TabMaintainer.TextOf(line) : string.Empty;

        // The stable table opened on `table` without asking the server - for tests.
        internal void OpenStableTableForTests(BoardTableAnswer table, BoardOverviewEntry? readAt = null)
        {
            this.OpenStableTable(table);
            this.thisStableTableReadAt = readAt;
        }

        private void CloseStableTable()
        {
            if (this.thisStableTable is not null && this.StableBoardTable.CurrentSheet?.Name is string sheet)
                this.thisStableSheetByBoard[this.thisStableTable.BoardId] = sheet;

            this.thisStableTable = null;
            this.thisStableTableReadAt = null;
            this.thisStableComparedWith = null;

            BoardTableEditor editor = this.StableBoardTable;
            editor.Clear();
            editor.FileSource = null;

            WindowMessage.Show(this.FindControl<TextBlock>("StableTableMessageText"), null, isError: false);
        }

        // ###########################################################################################
        // Reads the stable source's files for the board on screen and shows them, opened from its
        // public address. The caller holds the wait.
        // ###########################################################################################
        private async Task LoadStableFilesAsync()
        {
            if (this.thisClient is not ReviewApiClient client ||
                this.thisSession is not ReviewSession session ||
                this.ShownBoard is not { } board)
            {
                return;
            }

            ReviewApiResult<BoardFilesAnswer> answer = await client.GetBoardFilesAsync(session, board.BoardId, DataTreeNames.Production);

            if (!string.Equals(this.ShownBoard?.BoardId, board.BoardId, StringComparison.Ordinal))
                return;

            if (!answer.IsOk)
            {
                this.ShowStableFiles(null, null, board.BoardId);
                this.StableFiles.ShowMessage(answer.Message, isError: answer.Failure != ReviewApiFailure.NotFound);
                return;
            }

            // An older server answered with BETA's files: never drawn as the stable source's.
            if (!BoardSections.IsStableAnswer(answer.Value!))
            {
                this.ShowStableFiles(null, null, board.BoardId);
                this.StableFiles.ShowMessage(BoardSections.StableNeedsNewerServer, isError: true);
                return;
            }

            this.ShowStableFiles(answer.Value!, client, board.BoardId);
            this.thisStableFilesFor = board.BoardId;
            this.thisStableFilesReadAt = board;
        }

        private void ShowStableFiles(BoardFilesAnswer? files, ReviewApiClient? client, string boardId)
        {
            FileTreeView tree = this.StableFiles;

            tree.Files = files is null || client is null
                ? null
                : new FileTreeFiles(this, client, this.thisSession, submissionId: null, betaDataUrl: null, files.ProductionDataUrl);

            tree.ShowListing(files?.Files, openFolder: boardId);

            if (this.FindControl<TextBlock>("StableFilesHeadingText") is TextBlock heading)
                heading.Text = files is null ? string.Empty : BoardSections.StableFilesHeading;
        }

        // Shows a stable list without a server - for tests.
        internal void ShowStableFilesForTests(BoardFilesAnswer files, BoardOverviewEntry? readAt = null)
        {
            this.ShowStableFiles(files, client: null, files.BoardId);
            this.thisStableFilesFor = files.BoardId;
            this.thisStableFilesReadAt = readAt;
        }

        private void ClearStableFiles()
        {
            this.thisStableFilesFor = null;
            this.thisStableFilesReadAt = null;

            FileTreeView tree = this.StableFiles;
            tree.Files = null;
            tree.ShowListing(null, openFolder: null);

            if (this.FindControl<TextBlock>("StableFilesHeadingText") is TextBlock heading)
                heading.Text = string.Empty;
        }

        // ###########################################################################################
        // The stable source moved under what is held (a promotion): read again when on screen,
        // otherwise dropped, to be read when next shown.
        // ###########################################################################################
        private async Task CatchUpWithStableAsync(BoardOverviewEntry now)
        {
            bool onScreen = this.ShowsStable;

            if (this.HoldsStableTableFor(now.BoardId) &&
                this.thisStableTableReadAt is { } tableReadAt &&
                QueueRefreshRules.StableBoardChanged(tableReadAt, now))
            {
                // On screen - or what BETA's table on screen is compared with (BoardDetailView.Compare.cs).
                if ((onScreen || this.ComparesSources) && this.thisSection == BoardSection.BoardData)
                    await this.LoadStableTableAsync();
                else
                    this.CloseStableTable();
            }

            if (this.HoldsStableFilesFor(now.BoardId) &&
                this.thisStableFilesReadAt is { } filesReadAt &&
                QueueRefreshRules.StableBoardChanged(filesReadAt, now))
            {
                if (onScreen && this.thisSection == BoardSection.Files)
                    await this.LoadStableFilesAsync();
                else
                    this.ClearStableFiles();
            }
        }

        // Nothing of the stable source's board or files - another board chosen, or signed out.
        private void ForgetStableContent()
        {
            this.CloseStableTable();
            this.ClearStableFiles();
        }
    }
}
