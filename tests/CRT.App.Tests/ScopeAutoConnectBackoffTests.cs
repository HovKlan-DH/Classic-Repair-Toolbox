using Handlers.Oscilloscope;
using System;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// AUTOMATIC CONNECT WAITS LONGER AFTER EACH FAILED TRY (owner report, 2026-11-07: an Identify command the scope left
// unanswered was sent on a new connection every few seconds, "it seems to repeat forever(?)", until the scope answered
// nothing at all) - a second while all is well, doubling while it fails, half a minute at most.
// ###########################################################################################
public sealed class ScopeAutoConnectBackoffTests
{
    [Theory]
    [InlineData(-1, 1)]
    [InlineData(0, 1)]
    [InlineData(1, 2)]
    [InlineData(2, 4)]
    [InlineData(3, 8)]
    [InlineData(4, 16)]
    [InlineData(5, 30)]
    [InlineData(6, 30)]
    [InlineData(1000, 30)]
    public void The_wait_doubles_with_each_failed_try_and_stops_at_half_a_minute(int failures, int seconds)
    {
        Assert.Equal(TimeSpan.FromSeconds(seconds), ScopeAutoConnectBackoff.DelayAfter(failures));
    }

    // ###########################################################################################
    // *** A RUN OF FAILED TRIES ENDS WITH ANY TURN THAT IS NO FAILED TRY (code review, 2026-11-08). *** Counted down only
    // by a try that went through, five failures outlived a session connected by hand: the first try after that session
    // ended waited half a minute instead of a second, and the line saying why it failed - written for the first failure
    // of a run - was never written.
    // ###########################################################################################
    [Fact]
    public void A_run_of_failed_tries_ends_with_any_turn_that_is_no_failed_try()
    {
        int failures = 0;

        for (int turn = 0; turn < 5; turn++)
        {
            failures = ScopeAutoConnectBackoff.FailuresAfter(failures, ScopeAutoConnectBackoff.Turn.Failed);
        }

        Assert.Equal(5, failures);
        Assert.Equal(TimeSpan.FromSeconds(30), ScopeAutoConnectBackoff.DelayAfter(failures));

        // A session connected by hand, a hold, connect switched off, no scope to try: no try, and the run is over.
        failures = ScopeAutoConnectBackoff.FailuresAfter(failures, ScopeAutoConnectBackoff.Turn.NoTry);
        Assert.Equal(0, failures);

        // The session gone, the first try that fails is the first of a new run: a second's wait - and said, once.
        failures = ScopeAutoConnectBackoff.FailuresAfter(failures, ScopeAutoConnectBackoff.Turn.Failed);
        Assert.Equal(1, failures);
        Assert.Equal(TimeSpan.FromSeconds(2), ScopeAutoConnectBackoff.DelayAfter(failures));

        Assert.Equal(0, ScopeAutoConnectBackoff.FailuresAfter(7, ScopeAutoConnectBackoff.Turn.Connected));
        Assert.Equal(1, ScopeAutoConnectBackoff.FailuresAfter(-3, ScopeAutoConnectBackoff.Turn.Failed));
    }

    [Fact]
    public void A_failed_run_says_why_and_that_CRT_keeps_trying()
    {
        string line = ScopeAutoConnectBackoff.FailedLine("the scope took the connection but did not answer the Identify command \"*IaDN?\" in time");

        Assert.StartsWith("Automatic connect failed - the scope took the connection", line, StringComparison.Ordinal);
        Assert.Contains("\"*IaDN?\"", line, StringComparison.Ordinal);
        Assert.Contains("keeps trying", line, StringComparison.Ordinal);
        Assert.Contains("30 seconds", line, StringComparison.Ordinal);
    }
}
