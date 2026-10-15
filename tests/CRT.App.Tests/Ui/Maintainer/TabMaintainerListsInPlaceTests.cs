using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// *** THE MAINTAINER TAB'S LISTS ARE CHANGED IN PLACE (code review, 2026-10-13). *** The queue,
// the BETA list and the Boards list were each built again from nothing on every minute check: a
// click whose press landed on a row just replaced was lost (the Drafts tab's "I need to click it
// twice", 2026-10-10), and putting the selection back scrolled the list to it. A row still there
// now keeps its item - the SAME object - with its content brought up to date; the selected one is
// never made again. TabMaintainer.QueueItems.cs (ShowQueueItems, ShowItemsInPlace).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerListsInPlaceTests
{
    private static readonly DateTimeOffset Sent = new(2026, 10, 12, 9, 0, 0, TimeSpan.Zero);

    private static ReviewQueueRow Row(long id, string boardId = "Commodore/C64/250407", string summary = "Corrected U8.") =>
        new(id, boardId, "pending", summary, "c@example.com", TabMaintainerListsInPlaceTests.Sent, false, false, true);

    private static ListBoxItem ItemOf(TabMaintainer tab, long id) =>
        tab.QueueItemsForTests.Single(item => (item.Tag as ReviewQueueRow)?.Id == id);

    [Fact]
    public void A_queue_check_keeps_every_unchanged_row_and_heading_as_the_same_object()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            ReviewQueueRow first = Row(41);
            ReviewQueueRow second = Row(42);
            ReviewQueueRow other = Row(43, boardId: "Commodore/C128/310378");

            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [first, second, other], true));

            ListBox list = tab.FindControl<ListBox>("QueueList")!;
            ListBoxItem firstItem = ItemOf(tab, 41);
            ListBoxItem secondItem = ItemOf(tab, 42);
            ListBoxItem heading = tab.QueueItemsForTests[0];
            object? source = list.ItemsSource;

            // The minute check: the same answer, but a fresh one - another submission arrived, and 43 left.
            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(41), Row(42), Row(44)], true), background: true);

            Assert.Same(source, list.ItemsSource);
            Assert.Same(firstItem, ItemOf(tab, 41));
            Assert.Same(secondItem, ItemOf(tab, 42));
            Assert.Same(heading, tab.QueueItemsForTests[0]);
            Assert.DoesNotContain(tab.QueueItemsForTests, item => (item.Tag as ReviewQueueRow)?.Id == 43);
            Assert.Contains(tab.QueueItemsForTests, item => (item.Tag as ReviewQueueRow)?.Id == 44);

            // A heading per board, headings never selectable.
            Assert.Equal(4, tab.QueueItemsForTests.Count);
            Assert.All(tab.QueueItemsForTests.Where(item => item.Tag is null), item => Assert.False(item.IsEnabled));
        });
    }

    // A row that reads differently now keeps its item too - the content changes in place - and the
    // selected one stays selected, without being set again.
    [Fact]
    public void A_changed_row_keeps_its_item_and_the_selection_stays_on_it()
    {
        UiTest.Run(() =>
        {
            var tab = new TabMaintainer();
            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(41), Row(42)], true));

            ListBox list = tab.FindControl<ListBox>("QueueList")!;
            ListBoxItem selected = ItemOf(tab, 42);
            list.SelectedItem = selected;

            int selectionChanges = 0;
            list.SelectionChanged += (_, _) => selectionChanges++;

            tab.ApplyQueueResponse(new ReviewQueueResponse(true, [Row(41), Row(42, summary: "Corrected U8 and U9.")], true), background: true);

            Assert.Same(selected, ItemOf(tab, 42));
            Assert.Same(selected, list.SelectedItem);
            Assert.Equal(0, selectionChanges);
            Assert.Contains("Corrected U8 and U9.", tab.QueueTextsForTests());
            Assert.Equal("Corrected U8 and U9.", ((ReviewQueueRow)selected.Tag!).Summary);
        });
    }

    [Fact]
    public async Task The_BETA_list_keeps_an_unchanged_board_as_the_same_object()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();

            ProductionBoardRow Board(string id, string? betaRevision = null) =>
                new(id, "Commodore", id.Split('/')[1], id.Split('/')[2], betaRevision, "hash", null, null);

            await tab.ApplyBetaListAsync(new ProductionListResponse(true, [Board("Commodore/C64/250407"), Board("Commodore/C128/310378")]), background: true);

            ListBox list = tab.FindControl<ListBox>("BetaList")!;
            ListBoxItem c64 = ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => (item.Tag as ProductionBoardRow)?.BoardId == "Commodore/C64/250407");

            await tab.ApplyBetaListAsync(new ProductionListResponse(true, [Board("Commodore/C64/250407"), Board("Commodore/C128/310378", "2026-October-12")]), background: true);

            ListBoxItem again = ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => (item.Tag as ProductionBoardRow)?.BoardId == "Commodore/C64/250407");
            Assert.Same(c64, again);
            Assert.Contains(tab.BetaTextsForTests(), text => text.Contains("2026-October-12", StringComparison.Ordinal));
        });
    }

    [Fact]
    public async Task The_Boards_list_keeps_an_unchanged_board_as_the_same_object()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();

            BoardOverviewEntry Board(string id, int maintainers = 1) =>
                new(id, "Commodore", id.Split('/')[1], id.Split('/')[2], true, true, false, true, null, null, null, maintainers);

            await tab.ApplyBoardsListAsync(new BoardOverviewAnswer([Board("Commodore/C64/250407"), Board("Commodore/C128/310378")]), background: true);

            ListBox list = tab.FindControl<ListBox>("BoardsList")!;
            ListBoxItem c64 = ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => (item.Tag as BoardOverviewEntry)?.BoardId == "Commodore/C64/250407");

            await tab.ApplyBoardsListAsync(new BoardOverviewAnswer([Board("Commodore/C64/250407"), Board("Commodore/C128/310378", maintainers: 2)]), background: true);

            ListBoxItem again = ((IEnumerable<ListBoxItem>)list.ItemsSource!).Single(item => (item.Tag as BoardOverviewEntry)?.BoardId == "Commodore/C64/250407");
            Assert.Same(c64, again);
            Assert.Equal(2, ((IEnumerable<ListBoxItem>)list.ItemsSource!).Count());
        });
    }

    // ###########################################################################################
    // A failure in CRT itself (HandlerGuard) is said above the list of the screen on show - the
    // Account screen too, which had no line of its own until 2026-10-13.
    // ###########################################################################################
    [Fact]
    public async Task A_failure_on_the_Account_screen_is_said_above_its_list()
    {
        await UiTest.RunAsync(async () =>
        {
            var tab = new TabMaintainer();
            typeof(TabMaintainer).GetMethod("ShowQueuePanel", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(tab, null);

            await tab.ShowModeAsync(MaintainerMode.Account);
            tab.ShowHandlerFailureForTests(MaintainerWaitWording.Unexpected("boom"));

            TextBlock line = tab.FindControl<TextBlock>("AccountListMessageText")!;
            Assert.True(line.IsVisible);
            Assert.Equal(MaintainerWaitWording.Unexpected("boom"), line.Text);

            typeof(TabMaintainer).GetMethod("StopQueueChecks", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!.Invoke(tab, null);
        });
    }
}
