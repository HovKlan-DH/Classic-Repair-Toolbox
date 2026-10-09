using Avalonia.Controls;
using Avalonia.Threading;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// A board's table on the Boards screen as the WHOLE TAB handles it (2026-10-03): choosing another
// board asks about a change not published yet - Cancel keeps the board and its change - the tab
// counts that change as unsaved (quitting CRT and the tab's visibility read it), and a change the
// server made into a submission but could not publish opens it under Contributor Submissions, where
// it waits. (A published change goes straight to BETA - BoardDetailViewSectionsTests.)
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerBoardTableTests
{
    private const string First = "Commodore/C64/250407";
    private const string Second = "Commodore/C128/310378";

    private static BoardOverviewEntry Board(string boardId) =>
        new(boardId, "Commodore", boardId.Split('/')[1], boardId.Split('/')[2], true, true, false, true, null, null, null, 1);

    private static ReviewQueueRow Row(long id, string board) =>
        new(id, board, "pending", $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, true);

    // Two boards listed, the first on screen with its table open on BETA's board and one change in it.
    private static (TabMaintainer Main, ListBox List) WithChangedTable()
    {
        var main = new TabMaintainer();

        main.ApplyBoardsListAsync(
            new BoardOverviewAnswer([TabMaintainerBoardTableTests.Board(First), TabMaintainerBoardTableTests.Board(Second)]),
            background: true).GetAwaiter().GetResult();

        BoardDetailView view = main.BoardDetailForTests;
        view.ShowDetailForTests(new BoardDetailAnswer(TabMaintainerBoardTableTests.Board(First), [], [], []));

        var rows = new SubmissionRows { Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA" }] };
        view.OpenTableForTests(new BoardTableAnswer(First, "f", rows, MayEdit: true));

        BoardTableSheet components = view.BoardTableForTests.CommitAndGetDocument()!.FindSheet(BoardWorkbookSchema.SheetComponents)!;
        int friendly = BoardWorkbookSchema.Components.ColumnOrder.ToList().IndexOf(BoardWorkbookSchema.ColFriendlyName);
        components.Rows.Single().Cells[friendly].Text = "CPU 6510";

        return (main, main.FindControl<ListBox>("BoardsList")!);
    }

    private static ListBoxItem ItemOf(ListBox list, string boardId) =>
        ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => ((BoardOverviewEntry)item.Tag!).BoardId == boardId);

    [Fact]
    public void A_change_not_sent_in_a_boards_table_counts_as_unsaved_for_the_whole_tab()
    {
        UiTest.Run(() =>
        {
            (TabMaintainer main, _) = TabMaintainerBoardTableTests.WithChangedTable();

            Assert.True(main.BoardDetailForTests.HasUnsavedTableEdits);
            Assert.True(main.HasUnsavedTableEdits);
        });
    }

    // Cancel: the board stays chosen, on screen, with its change.
    [Fact]
    public async Task Choosing_another_board_asks_first_and_Cancel_keeps_the_board_and_its_change()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer main, ListBox list) = TabMaintainerBoardTableTests.WithChangedTable();
            var asked = new List<UnsavedTableEditsPrompt>();

            main.BoardDetailForTests.UnsavedTableEditsAnswerForTests = prompt =>
            {
                asked.Add(prompt);
                return UnsavedTableEditsChoice.Cancel;
            };

            list.SelectedItem = TabMaintainerBoardTableTests.ItemOf(list, Second);
            Dispatcher.UIThread.RunJobs();
            await Task.Yield();

            Assert.Equal([UnsavedTableEditsPrompt.LeavingBoard], asked);
            Assert.Equal(First, main.SelectedBoardForTests?.BoardId);
            Assert.Equal(First, main.BoardDetailForTests.ShownBoard?.BoardId);
            Assert.True(main.BoardDetailForTests.HasUnsavedTableEdits);
        });
    }

    // Discard: the other board is shown, and nothing of the first one's table stays.
    [Fact]
    public async Task Choosing_another_board_and_discarding_shows_it_with_no_table_left()
    {
        await UiTest.RunAsync(async () =>
        {
            (TabMaintainer main, ListBox list) = TabMaintainerBoardTableTests.WithChangedTable();

            main.BoardDetailForTests.UnsavedTableEditsAnswerForTests = _ => UnsavedTableEditsChoice.Discard;

            list.SelectedItem = TabMaintainerBoardTableTests.ItemOf(list, Second);
            Dispatcher.UIThread.RunJobs();
            await Task.Yield();

            Assert.Equal(Second, main.BoardDetailForTests.ShownBoard?.BoardId);
            Assert.False(main.BoardDetailForTests.HasUnsavedTableEdits);
            Assert.False(main.HasUnsavedTableEdits);
        });
    }

    // ###########################################################################################
    // A change made into a submission but NOT published: it is opened under Contributor Submissions -
    // even with another submission chosen there already, which the queue's own rule would otherwise keep.
    // ###########################################################################################
    [Fact]
    public async Task A_change_made_but_not_published_opens_its_submission_under_Contributor_Submissions()
    {
        await UiTest.RunAsync(async () =>
        {
            var main = new TabMaintainer();

            main.ApplyQueueResponse(new ReviewQueueResponse(
                CanPublish: true,
                Submissions: [TabMaintainerBoardTableTests.Row(41, Second), TabMaintainerBoardTableTests.Row(57, First)],
                IsAdministrator: true));

            await main.ShowModeAsync(MaintainerMode.Review);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(41, main.SelectedQueueRowForTests?.Id);

            await main.ShowModeAsync(MaintainerMode.Boards);
            await main.OpenSentSubmissionAsync(57);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(57, main.SelectedQueueRowForTests?.Id);
        });
    }
}
