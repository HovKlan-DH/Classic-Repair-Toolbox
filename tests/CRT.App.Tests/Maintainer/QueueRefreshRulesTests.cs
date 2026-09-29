using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// QueueRefreshRules - what the minute check reads beside the queue, and when a system's panel is
// read again (code review, 2026-09-27). The check used to read every list on every screen, and
// the Systems overview is the costly one on the server: both data trees walked and a month of
// board views counted, every minute, for every open maintainer window.
// ###########################################################################################
public sealed class QueueRefreshRulesTests
{
    // The minute check reads the overview only on the Systems screen, once it has been read at all.
    [Theory]
    [InlineData(MaintainerMode.Review)]
    [InlineData(MaintainerMode.Beta)]
    [InlineData(MaintainerMode.Admin)]
    public void The_minute_check_away_from_the_Systems_screen_does_not_read_the_overview(MaintainerMode shown)
    {
        Assert.False(QueueRefreshRules.ReadsSystemsOverview(minuteCheck: true, shown, overviewKnown: true));
    }

    [Fact]
    public void The_minute_check_on_the_Systems_screen_reads_the_overview()
    {
        Assert.True(QueueRefreshRules.ReadsSystemsOverview(minuteCheck: true, MaintainerMode.Systems, overviewKnown: true));
    }

    // Its badge counts the systems, so an overview never read is read whatever the screen.
    [Fact]
    public void An_overview_never_read_is_read_by_the_minute_check_too()
    {
        Assert.True(QueueRefreshRules.ReadsSystemsOverview(minuteCheck: true, MaintainerMode.Review, overviewKnown: false));
    }

    // Signing in, a decision, a screen's own button: everything, as before.
    [Theory]
    [InlineData(MaintainerMode.Review)]
    [InlineData(MaintainerMode.Systems)]
    public void Any_other_check_reads_the_overview(MaintainerMode shown)
    {
        Assert.True(QueueRefreshRules.ReadsSystemsOverview(minuteCheck: false, shown, overviewKnown: true));
    }

    // ###########################################################################################
    // The view count moves with every board opened in CRT; alone, it is not a reason to read the
    // whole panel again. Anything else is.
    // ###########################################################################################
    [Fact]
    public void A_system_whose_only_change_is_its_view_count_has_not_changed()
    {
        SystemOverviewEntry before = QueueRefreshRulesTests.System() with { ViewsLast30Days = 48 };

        Assert.False(QueueRefreshRules.SystemChanged(before, before with { ViewsLast30Days = 49 }));
        Assert.False(QueueRefreshRules.SystemChanged(before, before with { ViewsLast30Days = null }));
    }

    [Fact]
    public void A_system_whose_state_changed_has_changed()
    {
        SystemOverviewEntry before = QueueRefreshRulesTests.System() with { ViewsLast30Days = 48 };

        Assert.True(QueueRefreshRules.SystemChanged(before, before with { IsAwaitingProduction = true }));
        Assert.True(QueueRefreshRules.SystemChanged(before, before with { MaintainerCount = 3, ViewsLast30Days = 49 }));
    }

    private static SystemOverviewEntry System() =>
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
}
