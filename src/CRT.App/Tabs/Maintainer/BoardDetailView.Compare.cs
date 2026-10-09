using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "COMPARE SOURCES" ON BOARD DATA (owner request, 2026-10-09: "a "Compare sources" checkbox shown
    // right after the two radio buttons, "BETA" and "Stable" and if checked, then the below table
    // should show the changes as-if this was a normal board submission ... if I have selected the
    // "BETA" source, then when there is a change on a cell, then I see this text: "Stable source
    // value:" and likewise if I have chosen "Stable" then it would show "BETA source value:"").
    //
    // Ticked, BETA's table (BoardDetailView.Table.cs) is coloured against the stable source's board,
    // and the stable source's table (BoardDetailView.Stable.cs) against BETA's - each table built
    // with the OTHER source as what it is compared with, as a submission's table is built with the
    // published board. So both are read when Board data is shown, whichever is on screen. Unticked,
    // each is coloured against itself as opened, as before. The words are BoardSections'.
    //
    // *** NEVER UNDER A CHANGE NOT SAVED. *** Comparing builds BETA's table again, which would throw
    // a change typed into it away - so the box is off while BETA's table holds one (the table's
    // UnsavedChangesChanged), saying why. Saving builds BETA's table again anyway, compared as ticked.
    // A cell still being typed in counts too (code review, 2026-10-09): the box commits it first, as
    // every button of the table does, and nothing builds BETA's table under it quietly.
    //
    // *** EVERY TABLE READ GOES THROUGH ShowTables, *** so each is built once, and the other follows.
    //
    // What a save SENDS does not move with it: a save sends BETA's live rows applied to BETA as it
    // was read (BoardSections.RowsToSend) - never the stable source's rows drawn as deleted.
    //
    // REMEMBERED between launches (owner request, same day): Main hands the tick in through
    // TabMaintainer.UseRememberedComparison, and every click back out.
    // ###########################################################################################
    public partial class BoardDetailView
    {
        // The maintainer's tick, and where a new one is remembered.
        private bool thisCompareWanted;
        private Action<bool>? thisRememberCompare;

        // What each table was built against: the OTHER source's answer, or null - itself as opened.
        private BoardTableAnswer? thisTableComparedWith;
        private BoardTableAnswer? thisStableComparedWith;

        // Whether the two tables are compared - ticked, and the board in both sources.
        private bool ComparesSources => this.thisCompareWanted && BoardSections.CanCompareSources(this.ShownBoard);

        private CheckBox? CompareBox => this.FindControl<CheckBox>("CompareSourcesBox");

        private void WireCompare()
        {
            this.BoardTable.UnsavedChangesChanged += (_, _) => this.ApplyCompareBox();

            // A comparison held back while a cell was typed in is made once it is done - unless the
            // edit left a change, which keeps the table as built until saved.
            this.BoardTable.CellEditEnded += (_, _) => Dispatcher.UIThread.Post(this.ApplyComparison, DispatcherPriority.Background);

            this.ApplyCompareBox();
        }

        // The tick as last remembered - see the header.
        internal void UseRememberedComparison(bool compare, Action<bool> remember)
        {
            ArgumentNullException.ThrowIfNull(remember);

            this.thisCompareWanted = compare;
            this.thisRememberCompare = remember;
            this.ApplyCompareBox();
        }

        private async void OnCompareSourcesClick(object? sender, RoutedEventArgs e) =>
            await this.CompareSourcesAsync(this.CompareBox?.IsChecked == true);

        // ###########################################################################################
        // The maintainer's tick: remembered, the other source's table read when it is wanted and not
        // held - under the window's "please wait", since the maintainer pressed for it - and both
        // tables built again against what they are now compared with.
        // ###########################################################################################
        internal async Task CompareSourcesAsync(bool compare)
        {
            // A cell still in its editor is a change like any other (code review, 2026-10-09): the
            // grid does not commit it when the box takes the click, and building the table again
            // threw what was typed away without a word.
            this.BoardTable.CommitCellEdit();

            // The box is off then; a click that got here anyway changes nothing.
            if (this.HasUnsavedTableEdits)
            {
                this.ApplyCompareBox();
                return;
            }

            this.thisCompareWanted = compare;
            this.thisRememberCompare?.Invoke(compare);

            if (this.ShownBoard is { } board && this.thisSection == BoardSection.BoardData && this.NeedsReading(BoardSection.BoardData))
            {
                if (!await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingBoardTable(board.BoardId), this.LoadShownSectionAsync))
                    this.ShowMessage(WaitWording.NoAnswer, isError: true);
            }

            this.ApplyComparison();
        }

        // ###########################################################################################
        // *** TABLES READ ARE TAKEN FIRST AND BUILT AFTER (code review, 2026-10-09). *** Each table is
        // built against the other as it is when built. Built as each answer arrived, the stable
        // table of a compared board was built twice - against BETA as it was, then again once BETA's
        // came - and a BETA table opened on its own (after a publish) left the stable table compared
        // with BETA as it was before. So every table read or opened comes here: both answers are
        // taken, each is built ONCE against the other as it now is, and a table not read is built
        // again only when what it is compared with moved (ApplyComparison).
        //
        // `betaMessage` goes first in the line above BETA's table - what a publish just did, say.
        // ###########################################################################################
        private void ShowTables(BoardTableAnswer? beta, string? betaMessage, BoardTableAnswer? stable)
        {
            if (beta is not null)
                this.TakeTable(beta);

            if (stable is not null)
                this.TakeStableTable(stable);

            if (stable is not null)
                this.BuildStableTable();

            if (beta is not null)
            {
                this.BuildTable(betaMessage);

                // Why it cannot be changed, as just read - the board's detail says it too.
                this.ShowReadOnlyNotice(BoardSections.ReadOnlyReason(beta));
            }

            this.ApplyComparison();
        }

        // ###########################################################################################
        // Builds each table held again when what it is compared with is not what it should be now.
        // Never BETA's under a change not saved, nor under a cell still being typed in - it keeps
        // what it was built against until saved, or until the cell is done (WireCompare).
        // ###########################################################################################
        private void ApplyComparison()
        {
            if (this.thisTable is { } beta &&
                !ReferenceEquals(this.thisTableComparedWith, this.StableToCompareWith(beta.BoardId)) &&
                !this.HasUnsavedTableEdits &&
                !this.BoardTable.IsEditingCell)
            {
                this.TakeTable(beta);
                this.BuildTable(message: null);
            }

            if (this.thisStableTable is { } stable &&
                !ReferenceEquals(this.thisStableComparedWith, this.BetaToCompareWith(stable.BoardId)))
            {
                this.TakeStableTable(stable);
                this.BuildStableTable();
            }

            this.ApplyCompareBox();
        }

        // What BETA's table of `boardId` is to be compared with: the stable source's table, when
        // compared and held - else nothing.
        private BoardTableAnswer? StableToCompareWith(string boardId) =>
            this.ComparesSources && this.HoldsStableTableFor(boardId) ? this.thisStableTable : null;

        // And the stable source's: BETA's table as read.
        private BoardTableAnswer? BetaToCompareWith(string boardId) =>
            this.ComparesSources && this.HoldsTableFor(boardId) ? this.thisTable : null;

        // ###########################################################################################
        // The box: with Board data only, ticked as wanted, and off - saying why - when the board is
        // not in both sources or BETA's table holds a change not saved.
        // ###########################################################################################
        private void ApplyCompareBox()
        {
            if (this.CompareBox is not CheckBox box)
                return;

            string? unavailable = BoardSections.CompareUnavailable(this.ShownBoard, this.HasUnsavedTableEdits);

            box.IsVisible = this.thisSection == BoardSection.BoardData;
            box.IsChecked = this.thisCompareWanted;
            box.IsEnabled = unavailable is null;
            ToolTip.SetTip(box, unavailable ?? BoardSections.CompareSourcesTip);
        }

        // ###########################################################################################
        // The file cards of a compared table read the side it is compared with from the OTHER
        // source's address, each picture headed by its source (PublishedTableFileSource.Baseline).
        // ###########################################################################################
        private static PublishedTableBaseline? StableFilesBaseline(BoardTableAnswer? stable) =>
            stable is null
                ? null
                : new PublishedTableBaseline(stable.ProductionDataUrl, BoardSections.StableTreeName, BoardSections.StableSourceSide, BoardSections.BetaSourceSide);

        private static PublishedTableBaseline? BetaFilesBaseline(BoardTableAnswer? beta) =>
            beta is null
                ? null
                : new PublishedTableBaseline(beta.BetaDataUrl, BoardSections.BetaTreeName, BoardSections.BetaSourceSide, BoardSections.StableSourceSide);

        // What each table was built against, and the box - for tests.
        internal BoardTableAnswer? TableComparedWithForTests => this.thisTableComparedWith;

        // How often each table has been built - for tests, which hold ShowTables to once each.
        internal int TableBuildsForTests { get; private set; }

        internal int StableTableBuildsForTests { get; private set; }

        internal BoardTableAnswer? StableComparedWithForTests => this.thisStableComparedWith;

        internal CheckBox CompareBoxForTests => this.CompareBox!;
    }
}
