using System.Reflection;
using Avalonia.Controls;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// "Contributor Submissions" and "Beta > Prod" open on something (owner request, 2026-09-30: "it
// should select either whatever the user viewed last time (if set) or it should show the first
// entry, instead of the maintainer needing to press it") - TabMaintainer.OpenOnEntry.cs.
//
// No server: the lists are filled through the same Apply methods a server answer goes through, and
// choosing an entry asks nobody (no client), so what is pinned is WHICH entry is chosen.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class TabMaintainerOpenOnEntryTests
{
    private static readonly BindingFlags Any = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;

    private static ReviewQueueRow Row(long id, string board) =>
        new(id, board, "pending", $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, true);

    private static ProductionBoardRow Beta(string board) =>
        new(board, "Commodore", board.Split('/')[1], board.Split('/')[2], "2026-September-25", "hash", null, null, true);

    // Two boards, so "the first" is the first under the first heading, not merely the lowest id.
    private static TabMaintainer WithQueue()
    {
        var main = new TabMaintainer();

        main.ApplyQueueResponse(new ReviewQueueResponse(true,
        [
            Row(7, "Commodore/C64/250407"),
            Row(3, "Commodore/C128/310378"),
            Row(9, "Commodore/C128/310378")
        ], true));

        return main;
    }

    private static long? FirstShown(TabMaintainer main) =>
        main.QueueTextsForTests().Count == 0 ? null : main.SelectedQueueRowForTests?.Id;

    [Fact]
    public void Choosing_Contributor_Submissions_opens_the_first_submission_in_the_list()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = WithQueue();
            List<long> order = typeof(TabMaintainer).GetMethod("QueueIdsInListOrder", Any)!
                .Invoke(main, [main.FindControl<ListBox>("QueueList")!]) is List<long> ids ? ids : [];

            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();

            Assert.Equal(order[0], main.SelectedQueueRowForTests?.Id);
        });
    }

    // The one looked at last - and looking at one is what remembers it, through Main's callback.
    [Fact]
    public void Choosing_Contributor_Submissions_opens_the_submission_looked_at_last()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = WithQueue();
            long? remembered = null;
            main.UseRememberedSelections(9, null, id => remembered = id, _ => { });

            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();

            Assert.Equal(9, main.SelectedQueueRowForTests?.Id);
            Assert.Equal(9, remembered);
        });
    }

    // Remembered but gone - decided since, say: the first instead, never nothing.
    [Fact]
    public void A_remembered_submission_no_longer_in_the_queue_opens_the_first_instead()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = WithQueue();
            main.UseRememberedSelections(12345, null, _ => { }, _ => { });

            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();

            Assert.NotNull(main.SelectedQueueRowForTests);
            Assert.NotEqual(12345, main.SelectedQueueRowForTests!.Id);
        });
    }

    // ###########################################################################################
    // *** NEVER OVER SOMETHING OPEN. *** A submission decided elsewhere leaves the list but stays
    // open, with its table; choosing another on the way back to the screen would ask about that
    // table's unsaved changes out of the blue. So nothing is chosen while one is open.
    // ###########################################################################################
    [Fact]
    public void Nothing_is_chosen_while_a_submission_decided_elsewhere_is_still_open()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = WithQueue();
            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();
            long opened = main.SelectedQueueRowForTests!.Id;

            // The queue's own check finds it gone: decided by somebody else.
            main.ApplyQueueResponse(
                new ReviewQueueResponse(true, [Row(opened == 7 ? 3 : 7, "Commodore/C64/250407")], true),
                background: true);

            Assert.Null(main.SelectedQueueRowForTests);

            main.SelectOnEntryForTests();

            Assert.Null(main.SelectedQueueRowForTests);
        });
    }

    [Fact]
    public void Choosing_Beta_to_Prod_opens_the_board_looked_at_last_or_the_first()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            string? remembered = null;
            main.UseRememberedSelections(null, "Commodore/C128/310378", _ => { }, id => remembered = id);

            main.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                Beta("Commodore/C64/250407"),
                Beta("Commodore/C128/310378")
            ]), background: false).GetAwaiter().GetResult();

            main.ShowModeAsync(MaintainerMode.Beta).GetAwaiter().GetResult();

            Assert.Equal("Commodore/C128/310378", main.SelectedBetaRowForTests?.BoardId);
            Assert.Equal("Commodore/C128/310378", remembered);

            var fresh = new TabMaintainer();

            fresh.ApplyBetaListAsync(new ProductionListResponse(true,
            [
                Beta("Commodore/C64/250407"),
                Beta("Commodore/C128/310378")
            ]), background: false).GetAwaiter().GetResult();

            fresh.ShowModeAsync(MaintainerMode.Beta).GetAwaiter().GetResult();

            Assert.Equal("Commodore/C64/250407", fresh.SelectedBetaRowForTests?.BoardId);
        });
    }

    // The other two screens are not asked for this: choosing Boards opens nothing by itself.
    [Fact]
    public void Choosing_Boards_opens_nothing_by_itself()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = WithQueue();

            main.ShowModeAsync(MaintainerMode.Boards).GetAwaiter().GetResult();

            Assert.Null(main.SelectedQueueRowForTests);
            Assert.Null(main.SelectedBoardForTests);
        });
    }

    // ###########################################################################################
    // *** WITH NOTHING WAITING, THE TAB OPENS ON BOARDS (owner request, 2026-10-04: "if there is no
    // queue awaiting, when opening the "Maintainer" tab, then go to "Systems" and show the last
    // selected system"). *** Opening is OpenForTests - what showing the tab or signing in on it
    // does. Signed in without a server (UseSessionForTests, no client), the three lists filled
    // through the same Apply methods an answer goes through.
    // ###########################################################################################

    private static BoardOverviewEntry Board(string board) =>
        new(board, "Commodore", board.Split('/')[1], board.Split('/')[2], true, true, false, true, null, null, null, 1);

    private static ReviewQueueRow Waiting(long id, string board, bool awaitsYou) =>
        Row(id, board, "pending", awaitsYou);

    private static ReviewQueueRow Row(long id, string board, string state, bool awaitsYou) =>
        new(id, board, state, $"Submission {id}.", "c@example.com", DateTimeOffset.UtcNow.AddDays(-1), false, false, awaitsYou);

    // Signed in, with the queue and the BETA list read - what makes the badge's number a real answer.
    private static TabMaintainer SignedIn(ReviewQueueRow[] queue, ProductionBoardRow[]? beta = null, string? rememberedBoard = null, Action<string?>? rememberBoard = null)
    {
        var main = new TabMaintainer();
        main.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));
        main.UseRememberedSelections(null, null, _ => { }, _ => { }, rememberedBoard, rememberBoard ?? (_ => { }));

        main.ApplyQueueResponse(new ReviewQueueResponse(true, queue, true));
        main.ApplyBetaListAsync(new ProductionListResponse(true, beta ?? []), background: false).GetAwaiter().GetResult();

        return main;
    }

    private static void WithBoards(TabMaintainer main) =>
        main.ApplyBoardsListAsync(
            new BoardOverviewAnswer([Board("Commodore/C64/250407"), Board("Commodore/C128/310378")]),
            background: true).GetAwaiter().GetResult();

    [Fact]
    public void Opening_with_nothing_waiting_goes_to_Boards_on_the_board_looked_at_last()
    {
        UiTest.Run(() =>
        {
            string? remembered = null;
            TabMaintainer main = SignedIn([], rememberedBoard: "Commodore/C128/310378", rememberBoard: id => remembered = id);
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Boards, main.ShownMode);
            Assert.Equal("Commodore/C128/310378", main.SelectedBoardForTests?.BoardId);
            Assert.Equal("Commodore/C128/310378", remembered);
        });
    }

    // Nothing remembered - or remembered and gone: the first board, as the queues open on their first.
    [Fact]
    public void Opening_on_Boards_with_no_board_remembered_shows_the_first()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([], rememberedBoard: "Commodore/VIC-20/250403");
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Boards, main.ShownMode);
            Assert.Equal("Commodore/C64/250407", main.SelectedBoardForTests?.BoardId);
        });
    }

    // Choosing a board is what remembers it - for the next opening, after a restart too.
    [Fact]
    public void Choosing_a_board_remembers_it()
    {
        UiTest.Run(() =>
        {
            string? remembered = null;
            TabMaintainer main = SignedIn([], rememberBoard: id => remembered = id);
            main.ShowModeAsync(MaintainerMode.Boards).GetAwaiter().GetResult();
            WithBoards(main);

            ListBox boards = main.FindControl<ListBox>("BoardsList")!;
            boards.SelectedItem = ((IEnumerable<ListBoxItem>)boards.ItemsSource!).Last();

            Assert.Equal("Commodore/C128/310378", remembered);
        });
    }

    [Fact]
    public void Opening_with_a_submission_waiting_for_you_stays_on_the_queue_and_opens_it()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([Waiting(7, "Commodore/C64/250407", awaitsYou: true)]);
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(7, main.SelectedQueueRowForTests?.Id);
            Assert.Null(main.SelectedBoardForTests);
        });
    }

    // A board in BETA waiting for this account is waiting too - the tab's badge counts both queues.
    [Fact]
    public void Opening_with_a_BETA_board_waiting_for_you_does_not_go_to_Boards()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([], [Beta("Commodore/C64/250407")]);
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Null(main.SelectedBoardForTests);
        });
    }

    // "Awaiting" in the badge's sense: a submission waiting only for the OTHER approver is not.
    [Fact]
    public void A_submission_waiting_only_for_the_other_approver_does_not_keep_the_tab_on_the_queue()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([Waiting(7, "Commodore/C64/250407", awaitsYou: false)]);
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Boards, main.ShownMode);
            Assert.Null(main.SelectedQueueRowForTests);
            Assert.NotNull(main.SelectedBoardForTests);
        });
    }

    // ###########################################################################################
    // *** NEVER AWAY FROM SOMETHING OPEN. *** A submission the maintainer opened, which waits only for
    // the other approver, is where it was on coming back to the tab.
    // ###########################################################################################
    [Fact]
    public void Opening_again_with_a_submission_open_stays_on_it()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([Waiting(7, "Commodore/C64/250407", awaitsYou: false)]);
            WithBoards(main);
            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();

            Assert.Equal(7, main.SelectedQueueRowForTests?.Id);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(7, main.SelectedQueueRowForTests?.Id);
        });
    }

    // ###########################################################################################
    // *** NOTHING OPENS UNTIL IT IS KNOWN WHETHER ANYTHING WAITS. *** With the BETA list not read yet,
    // the queue's first submission is NOT opened (it would then hold the tab on the queue); once the
    // list is there, the opening goes to Boards.
    // ###########################################################################################
    [Fact]
    public void Nothing_opens_before_both_lists_are_known_and_then_the_tab_goes_to_Boards()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            main.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));
            main.ApplyQueueResponse(new ReviewQueueResponse(true, [Waiting(7, "Commodore/C64/250407", awaitsYou: false)], true));
            WithBoards(main);

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Null(main.SelectedQueueRowForTests);

            main.ApplyBetaListAsync(new ProductionListResponse(true, []), background: false).GetAwaiter().GetResult();
            main.SelectOnEntryForTests();

            Assert.Equal(MaintainerMode.Boards, main.ShownMode);
            Assert.Equal("Commodore/C64/250407", main.SelectedBoardForTests?.BoardId);
        });
    }

    // But one list showing work is enough: a submission waiting for this account opens at once, with
    // the BETA list not read - or not readable - yet.
    [Fact]
    public void A_submission_waiting_for_you_opens_without_waiting_for_the_BETA_list()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            main.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));
            main.ApplyQueueResponse(new ReviewQueueResponse(true, [Waiting(7, "Commodore/C64/250407", awaitsYou: true)], true));

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(7, main.SelectedQueueRowForTests?.Id);
        });
    }

    // The first look after a launch reads the Boards list only once the screen is Boards: the
    // board is chosen when it arrives.
    [Fact]
    public void The_board_is_chosen_once_the_Boards_list_arrives()
    {
        UiTest.Run(() =>
        {
            TabMaintainer main = SignedIn([], rememberedBoard: "Commodore/C128/310378");

            main.OpenForTests();

            Assert.Equal(MaintainerMode.Boards, main.ShownMode);
            Assert.Null(main.SelectedBoardForTests);

            WithBoards(main);
            main.SelectOnEntryForTests();

            Assert.Equal("Commodore/C128/310378", main.SelectedBoardForTests?.BoardId);
        });
    }

    // A screen chosen while the opening was undecided is the maintainer's choice, and stays.
    [Fact]
    public void A_screen_chosen_before_the_opening_was_decided_is_kept()
    {
        UiTest.Run(() =>
        {
            var main = new TabMaintainer();
            main.UseSessionForTests(new ReviewSession("token", DateTimeOffset.UtcNow.AddDays(30), 7, "dh@example.com", "Dennis"));
            main.ApplyQueueResponse(new ReviewQueueResponse(true, [Waiting(7, "Commodore/C64/250407", awaitsYou: false)], true));

            main.OpenForTests();
            main.ShowModeAsync(MaintainerMode.Review).GetAwaiter().GetResult();

            main.ApplyBetaListAsync(new ProductionListResponse(true, []), background: false).GetAwaiter().GetResult();
            main.SelectOnEntryForTests();

            Assert.Equal(MaintainerMode.Review, main.ShownMode);
            Assert.Equal(7, main.SelectedQueueRowForTests?.Id);
        });
    }
}
