using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// QueueRefreshRules - what the minute check reads beside the queue, and when a board's panel is
// read again (code review, 2026-09-27). The check used to read every list on every screen, and
// the Boards overview is the costly one on the server: both data trees walked and a month of
// board views counted, every minute, for every open maintainer window.
// ###########################################################################################
public sealed class QueueRefreshRulesTests
{
    // The minute check reads the overview only on the Boards screen, once it has been read at all.
    [Theory]
    [InlineData(MaintainerMode.Review)]
    [InlineData(MaintainerMode.Beta)]
    [InlineData(MaintainerMode.Account)]
    public void The_minute_check_away_from_the_Boards_screen_does_not_read_the_overview(MaintainerMode shown)
    {
        Assert.False(QueueRefreshRules.ReadsBoardsOverview(minuteCheck: true, shown, overviewKnown: true));
    }

    [Fact]
    public void The_minute_check_on_the_Boards_screen_reads_the_overview()
    {
        Assert.True(QueueRefreshRules.ReadsBoardsOverview(minuteCheck: true, MaintainerMode.Boards, overviewKnown: true));
    }

    // Its badge counts the boards, so an overview never read is read whatever the screen.
    [Fact]
    public void An_overview_never_read_is_read_by_the_minute_check_too()
    {
        Assert.True(QueueRefreshRules.ReadsBoardsOverview(minuteCheck: true, MaintainerMode.Review, overviewKnown: false));
    }

    // Signing in, a decision, a screen's own button: everything, as before.
    [Theory]
    [InlineData(MaintainerMode.Review)]
    [InlineData(MaintainerMode.Boards)]
    public void Any_other_check_reads_the_overview(MaintainerMode shown)
    {
        Assert.True(QueueRefreshRules.ReadsBoardsOverview(minuteCheck: false, shown, overviewKnown: true));
    }

    // ###########################################################################################
    // The view count moves with every board opened in CRT; alone, it is not a reason to read the
    // whole panel again. Anything else is.
    // ###########################################################################################
    [Fact]
    public void A_board_whose_only_change_is_its_view_count_has_not_changed()
    {
        BoardOverviewEntry before = QueueRefreshRulesTests.Board() with { ViewsLast30Days = 48 };

        Assert.False(QueueRefreshRules.BoardChanged(before, before with { ViewsLast30Days = 49 }));
        Assert.False(QueueRefreshRules.BoardChanged(before, before with { ViewsLast30Days = null }));
    }

    [Fact]
    public void A_board_whose_state_changed_has_changed()
    {
        BoardOverviewEntry before = QueueRefreshRulesTests.Board() with { ViewsLast30Days = 48 };

        Assert.True(QueueRefreshRules.BoardChanged(before, before with { IsAwaitingProduction = true }));
        Assert.True(QueueRefreshRules.BoardChanged(before, before with { MaintainerCount = 3, ViewsLast30Days = 49 }));
    }

    // ###########################################################################################
    // *** A BOARD'S TABLE AND FILES FOLLOW BETA (owner report, 2026-10-04). *** A fix approved
    // and promoted went on showing, on the Boards screen, the warning it fixed - the table was read
    // once and kept until another board was chosen. BETA's content hash, which every publish to
    // BETA and every push-back moves, says when they are out of date.
    // ###########################################################################################
    [Fact]
    public void BETA_has_moved_when_its_content_hash_has()
    {
        BoardOverviewEntry readAt = QueueRefreshRulesTests.Board() with { BetaContentHash = "aaa" };

        Assert.False(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { ViewsLast30Days = 12, MaintainerCount = 5 }));
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { BetaContentHash = "bbb" }));

        // A published revision on the same day moves the hash and not the revision - still moved.
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { BetaContentHash = "bbb", BetaRevision = readAt.BetaRevision }));

        // A board published for the first time (no hash before, one now).
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt with { BetaContentHash = null }, readAt));

        // With a hash on both sides, nothing else counts: a promotion leaves BETA as it is.
        Assert.False(QueueRefreshRules.BetaBoardChanged(readAt with { IsAwaitingProduction = true }, readAt));
    }

    // An older server sends no hash: then the facts that move with a publish to BETA and a push-back.
    [Fact]
    public void Without_a_content_hash_BETA_has_moved_when_its_revision_or_waiting_moved()
    {
        BoardOverviewEntry readAt = QueueRefreshRulesTests.Board();

        Assert.False(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { MaintainerCount = 7 }));
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { BetaRevision = "2026-October-4" }));
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { IsAwaitingProduction = true }));
        Assert.True(QueueRefreshRules.BetaBoardChanged(readAt, readAt with { InBeta = false }));
    }

    // ###########################################################################################
    // THE STABLE SOURCE'S BOARD MOVED (2026-10-04): only a promotion moves it, and a promotion sets
    // the stable revision and when it was published there - or the board appears there first.
    // Anything else about the board leaves a stable table and files as they are.
    // ###########################################################################################
    [Fact]
    public void The_stable_board_moved_only_when_a_promotion_moved_it()
    {
        var readAt = new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true,
            "2026-October-4", "2026-September-25", new DateTimeOffset(2026, 9, 25, 12, 0, 0, TimeSpan.Zero), 1, BetaContentHash: "a");

        Assert.False(QueueRefreshRules.StableBoardChanged(readAt, readAt with { BetaContentHash = "b", IsAwaitingProduction = true, MaintainerCount = 2 }));
        Assert.True(QueueRefreshRules.StableBoardChanged(readAt, readAt with { ProductionRevision = "2026-October-4" }));
        Assert.True(QueueRefreshRules.StableBoardChanged(readAt, readAt with { ProductionPublishedUtc = readAt.ProductionPublishedUtc!.Value.AddDays(9) }));
        Assert.True(QueueRefreshRules.StableBoardChanged(readAt with { InProduction = false }, readAt));
    }

    // ###########################################################################################
    // The TABLE also carries whether this account may change it - off while the board waits under
    // BETA > Stable, for a board closed to contributions, and by who maintains it. A promotion
    // moves only that: the table read while it waited stayed read-only after it no longer did.
    // ###########################################################################################
    [Fact]
    public void A_boards_table_is_out_of_date_when_its_board_or_who_may_change_it_moved()
    {
        BoardOverviewEntry readAt = QueueRefreshRulesTests.Board() with { BetaContentHash = "aaa", IsAwaitingProduction = true };

        Assert.False(QueueRefreshRules.BoardTableChanged(readAt, readAt with { ViewsLast30Days = 3 }));
        Assert.False(QueueRefreshRules.BoardTableChanged(readAt, readAt with { ProductionRevision = "2026-October-4" }));

        Assert.True(QueueRefreshRules.BoardTableChanged(readAt, readAt with { BetaContentHash = "bbb" }));
        Assert.True(QueueRefreshRules.BoardTableChanged(readAt, readAt with { IsAwaitingProduction = false }));
        Assert.True(QueueRefreshRules.BoardTableChanged(readAt, readAt with { IsAccepting = false }));
        Assert.True(QueueRefreshRules.BoardTableChanged(readAt, readAt with { MaintainerCount = 3 }));
    }

    // ###########################################################################################
    // One maintainer swapped for another keeps the count - and the account swapped in or out had a
    // table saying the opposite of what the server would do (code review, 2026-10-04). Given who
    // maintains it at both readings, the swap counts; the same people in another order do not.
    // ###########################################################################################
    [Fact]
    public void A_maintainer_swapped_for_another_puts_a_boards_table_out_of_date()
    {
        BoardOverviewEntry readAt = QueueRefreshRulesTests.Board() with { BetaContentHash = "aaa", MaintainerCount = 2 };

        Assert.True(QueueRefreshRules.BoardTableChanged(readAt, readAt, [1, 2], [1, 3]));
        Assert.False(QueueRefreshRules.BoardTableChanged(readAt, readAt, [1, 2], [2, 1]));

        // Without the lists, the count stands in - and cannot see the swap.
        Assert.False(QueueRefreshRules.BoardTableChanged(readAt, readAt, null, [1, 3]));
    }

    private static BoardOverviewEntry Board() =>
        new("Commodore/C64/250407", "Commodore", "C64", "250407",
            InBeta: true, InProduction: true, IsAwaitingProduction: false, IsAccepting: true,
            BetaRevision: "2026-September-27", ProductionRevision: "2026-May-14",
            ProductionPublishedUtc: new DateTimeOffset(2026, 5, 14, 0, 0, 0, TimeSpan.Zero),
            MaintainerCount: 2);

    // ###########################################################################################
    // When the queue checks itself: never asked, or asked long enough ago. (Lived in the separate
    // application's MaintainerSettingsTests until 2026-09-29, when that file went with its settings
    // class; the rule it pins is this class's.)
    // ###########################################################################################
    [Fact]
    public void A_queue_check_is_due_when_never_asked_or_asked_long_enough_ago()
    {
        var now = new DateTimeOffset(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        Assert.True(QueueRefreshRules.IsDue(null, now, QueueRefreshRules.ActivationGap));
        Assert.False(QueueRefreshRules.IsDue(now.AddSeconds(-5), now, QueueRefreshRules.ActivationGap));
        Assert.True(QueueRefreshRules.IsDue(now.AddSeconds(-15), now, QueueRefreshRules.ActivationGap));
        Assert.Equal(TimeSpan.FromMinutes(1), QueueRefreshRules.PollInterval);
    }

    // ###########################################################################################
    // *** OFF SCREEN, ONLY THE BADGE'S TWO LISTS - AND ONLY WHILE THE BADGE CAN BE SEEN (owner
    // request, 2026-09-30). *** The tab's badge in CRT's row of tabs must stay current while the
    // maintainer works elsewhere; the rest (the Boards overview walks both data trees) waits
    // until the tab is looked at, and a minimised CRT - or the tab turned off - asks nothing.
    // ###########################################################################################
    [Theory]
    [InlineData(true, true, QueueCheck.Everything)]
    [InlineData(true, false, QueueCheck.Everything)]
    [InlineData(false, true, QueueCheck.BadgesOnly)]
    [InlineData(false, false, QueueCheck.Nothing)]
    public void The_minute_check_reads_everything_on_screen_and_only_the_badges_lists_off_it(
        bool onScreenAndInFront, bool badgeCanBeSeen, QueueCheck expected)
    {
        Assert.Equal(expected, QueueRefreshRules.MinuteCheck(onScreenAndInFront, badgeCanBeSeen));
    }
}
