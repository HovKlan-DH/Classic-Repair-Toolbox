using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// NoProgressWatchdog - a download gives up when NOTHING arrives for the limit, not when the whole
// file takes longer than a fixed budget (code review, 2026-10-13: under a 30-second total budget no
// file above ~20 MB ever arrived on a slower line). Time is moved on by hand (ManualTimeProvider).
public sealed class NoProgressWatchdogTests
{
    private static readonly TimeSpan Limit = TimeSpan.FromSeconds(30);

    [Fact]
    public void Nothing_arriving_for_the_limit_cancels_and_says_it_stalled()
    {
        var time = new ManualTimeProvider();
        using var watchdog = new NoProgressWatchdog(Limit, time: time);

        time.Advance(TimeSpan.FromSeconds(29));
        Assert.False(watchdog.Token.IsCancellationRequested);

        time.Advance(TimeSpan.FromSeconds(1));
        Assert.True(watchdog.Token.IsCancellationRequested);
        Assert.True(watchdog.Stalled);
    }

    // The whole point: a long download that keeps moving is never cut off, however long it takes.
    [Fact]
    public void Steady_progress_keeps_it_going_far_past_the_limit()
    {
        var time = new ManualTimeProvider();
        using var watchdog = new NoProgressWatchdog(Limit, time: time);

        for (int chunk = 0; chunk < 100; chunk++)
        {
            time.Advance(TimeSpan.FromSeconds(20));
            watchdog.Progress();
        }

        Assert.False(watchdog.Token.IsCancellationRequested);

        time.Advance(Limit);
        Assert.True(watchdog.Stalled);
    }

    // The caller giving up is not a stall.
    [Fact]
    public void The_outer_token_cancels_without_counting_as_a_stall()
    {
        var time = new ManualTimeProvider();
        using var outer = new CancellationTokenSource();
        using var watchdog = new NoProgressWatchdog(Limit, outer.Token, time);

        outer.Cancel();

        Assert.True(watchdog.Token.IsCancellationRequested);
        Assert.False(watchdog.Stalled);
    }

    [Fact]
    public void Progress_after_a_stall_does_not_bring_it_back()
    {
        var time = new ManualTimeProvider();
        using var watchdog = new NoProgressWatchdog(Limit, time: time);

        time.Advance(Limit);
        watchdog.Progress();

        Assert.True(watchdog.Token.IsCancellationRequested);
    }
}
