using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// WaitLimit - when a wait the user is watching ends (owner decision, 2026-09-28: "It should await
// up to 2 minutes, until it will auto-close, but of course it must be solid in validating if it did
// finish").
//
// THE CLOCK IS THE TEST'S. Every limit here is a task the test completes by hand ("two minutes
// have now passed"), so nothing waits on real time and nothing races a timer.
// ###########################################################################################
public sealed class WaitLimitTests
{
    // Every limit WaitLimit starts, in order, each completable by the test and each remembering
    // whether WaitLimit stopped caring about it (its token cancelled - progress restarted the clock).
    private sealed class ManualClock
    {
        private readonly object thisGate = new();
        private readonly List<(TaskCompletionSource Passed, CancellationToken Token)> thisLimits = [];

        public Task Start(CancellationToken token)
        {
            var passed = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);

            lock (this.thisGate)
                this.thisLimits.Add((passed, token));

            return passed.Task;
        }

        public int Started
        {
            get
            {
                lock (this.thisGate)
                    return this.thisLimits.Count;
            }
        }

        public (TaskCompletionSource Passed, CancellationToken Token) this[int index]
        {
            get
            {
                lock (this.thisGate)
                    return this.thisLimits[index];
            }
        }

        // WaitLimit's loop runs its continuations off the reporting thread, so the next limit
        // appears a moment after the report - waited for here, never by sleeping a fixed time.
        public async Task WaitForStartedAsync(int count)
        {
            DateTime giveUp = DateTime.UtcNow.AddSeconds(30);

            while (this.Started < count)
            {
                Assert.True(DateTime.UtcNow < giveUp, $"Expected {count} limits to be started, saw {this.Started}.");
                await Task.Delay(5);
            }
        }
    }

    [Fact]
    public async Task Work_that_finishes_gives_its_value_and_is_not_a_timeout()
    {
        var clock = new ManualClock();

        WaitResult<int> result = await WaitLimit.RunAsync(_ => Task.FromResult(42), clock.Start);

        Assert.False(result.IsTimedOut);
        Assert.Equal(42, result.Value);
    }

    // ###########################################################################################
    // *** THE LIMIT PASSING ENDS THE WAIT, AND SAYS SO. *** The work's token is cancelled - a
    // request stops waiting - and the answer is TimedOut, not a failure: the server may have
    // finished regardless, which only the caller can go and check.
    // ###########################################################################################
    [Fact]
    public async Task Work_that_never_answers_times_out_when_the_limit_passes_and_its_token_is_cancelled()
    {
        var clock = new ManualClock();
        CancellationToken seen = default;
        var never = new TaskCompletionSource<int>();

        Task<WaitResult<int>> waiting = WaitLimit.RunAsync(context =>
        {
            seen = context.Token;
            return never.Task;
        }, clock.Start);

        await clock.WaitForStartedAsync(1);
        Assert.False(waiting.IsCompleted);

        clock[0].Passed.SetResult();
        WaitResult<int> result = await waiting;

        Assert.True(result.IsTimedOut);
        Assert.Equal(0, result.Value);
        Assert.True(seen.IsCancellationRequested);
    }

    // ###########################################################################################
    // *** TWO MINUTES WITH NOTHING HAPPENING, NOT TWO MINUTES IN ALL. *** A large upload on a slow
    // line runs past two minutes while visibly moving; each report starts the clock again, so the
    // limit that was running when it reported no longer counts.
    // ###########################################################################################
    [Fact]
    public async Task A_report_starts_the_clock_again_so_the_old_limit_passing_ends_nothing()
    {
        var clock = new ManualClock();
        var finish = new TaskCompletionSource<string>();
        WaitContext? context = null;

        Task<WaitResult<string>> waiting = WaitLimit.RunAsync(given =>
        {
            context = given;
            return finish.Task;
        }, clock.Start);

        await clock.WaitForStartedAsync(1);

        context!.Report("Uploading 1 of 3");
        await clock.WaitForStartedAsync(2);

        // The first clock was abandoned; its passing now is the old two minutes, not a new one.
        Assert.True(clock[0].Token.IsCancellationRequested);
        clock[0].Passed.TrySetResult();
        await Task.Delay(20);
        Assert.False(waiting.IsCompleted);

        finish.SetResult("done");
        WaitResult<string> result = await waiting;

        Assert.False(result.IsTimedOut);
        Assert.Equal("done", result.Value);
    }

    [Fact]
    public async Task After_a_report_the_new_limit_passing_still_times_the_work_out()
    {
        var clock = new ManualClock();
        WaitContext? context = null;

        Task<WaitResult<bool>> waiting = WaitLimit.RunAsync(given =>
        {
            context = given;
            return new TaskCompletionSource().Task;
        }, clock.Start);

        await clock.WaitForStartedAsync(1);
        context!.Report("Still going");
        await clock.WaitForStartedAsync(2);

        clock[1].Passed.SetResult();

        Assert.True((await waiting).IsTimedOut);
    }

    [Fact]
    public async Task Each_report_hands_on_its_sentence_and_position()
    {
        var clock = new ManualClock();
        var reports = new List<WaitReport>();

        await WaitLimit.RunAsync(context =>
        {
            context.Report("Packaging... 50%", 0.5);
            context.Report("Sending...");
            return Task.CompletedTask;
        }, clock.Start, report => reports.Add(report));

        Assert.Equal(
            [new WaitReport("Packaging... 50%", 0.5), new WaitReport("Sending...", null)],
            reports);
    }

    // An exception is the caller's, whether thrown before the work's first await or after it.
    [Fact]
    public async Task An_exception_from_the_work_reaches_the_caller_unchanged()
    {
        var clock = new ManualClock();

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            WaitLimit.RunAsync<int>(_ => throw new InvalidOperationException("before any await"), clock.Start));

        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            WaitLimit.RunAsync<int>(async _ =>
            {
                await Task.Yield();
                throw new InvalidOperationException("after an await");
            }, clock.Start));
    }

    // ###########################################################################################
    // *** WORK THAT CANNOT BE STOPPED CAN STILL BE WATCHED. *** A PDF being written does not listen
    // to the token. After the timeout, Work is that same work, so the caller can say when it did
    // finish - the "validating if it did finish" half for local work.
    // ###########################################################################################
    [Fact]
    public async Task After_a_timeout_the_work_can_still_be_watched_to_its_end()
    {
        var clock = new ManualClock();
        var slow = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);

        Task<WaitResult<string>> waiting = WaitLimit.RunAsync(_ => slow.Task, clock.Start);

        await clock.WaitForStartedAsync(1);
        clock[0].Passed.SetResult();

        WaitResult<string> result = await waiting;
        Assert.True(result.IsTimedOut);
        Assert.False(result.Work.IsCompleted);

        slow.SetResult("written");
        Assert.Equal("written", await result.Work);
    }

    [Fact]
    public async Task Work_with_no_value_finishes_as_true()
    {
        WaitResult<bool> result = await WaitLimit.RunAsync(_ => Task.CompletedTask, new ManualClock().Start);

        Assert.False(result.IsTimedOut);
        Assert.True(result.Value);
    }

    [Fact]
    public void The_limit_is_two_minutes()
    {
        Assert.Equal(TimeSpan.FromMinutes(2), WaitLimit.Maximum);
    }
}
