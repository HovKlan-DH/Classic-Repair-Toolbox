using System;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // How long anything the user waits for may take, and what happens when it takes longer (owner
    // decision, 2026-09-28: "It should await up to 2 minutes, until it will auto-close, but of
    // course it must be solid in validating if it did finish").
    //
    // No control here - the BusyOverlay shows the wait, this decides when it ends - so every rule
    // is a unit test with no display and no clock (WaitLimitTests).
    //
    // *** TWO MINUTES WITH NOTHING HAPPENING, NOT TWO MINUTES IN ALL. *** Work that reports
    // progress (WaitContext.Report) starts the clock again each time: a large upload on a slow line
    // legitimately runs for longer than two minutes while visibly moving, and cutting it off at a
    // fixed total would abort work that was going fine. Work that reports nothing - one request
    // waiting for one answer - gets two minutes in all.
    //
    // *** A TIMEOUT IS NOT A FAILURE, and the caller is told which it was. *** When the limit
    // passes, the work's token is cancelled (a request stops waiting) and the result says
    // TimedOut. The SERVER may well have finished regardless - a publish carries on after the
    // client stops listening - so a caller that changed something must look again before it says
    // what happened. That look is the "validating if it did finish" half, and it is the caller's,
    // because only the caller knows what "finished" means for its own operation.
    // ###########################################################################################
    public static class WaitLimit
    {
        // The limit itself. One value for every wait, so no screen waits longer than another.
        public static readonly TimeSpan Maximum = TimeSpan.FromMinutes(2);

        // ###########################################################################################
        // Runs `work` until it finishes, or until `limitPassed` completes with no progress reported
        // since the last time it was started.
        //
        // `limitPassed` is the clock: given a token (cancelled when progress restarts it), it
        // returns a task that completes when the limit has passed. Null is the real one, a delay of
        // Maximum. It is a parameter so a test can say "two minutes have now passed" exactly when
        // it wants, rather than waiting or racing a timer.
        //
        // `onReport` receives every sentence the work reports, for the overlay to show.
        //
        // An exception from the work - other than the cancellation this method itself caused - is
        // the caller's, thrown from here exactly as the work threw it.
        // ###########################################################################################
        public static async Task<WaitResult<T>> RunAsync<T>(
            Func<WaitContext, Task<T>> work,
            Func<CancellationToken, Task>? limitPassed = null,
            Action<WaitReport>? onReport = null)
        {
            ArgumentNullException.ThrowIfNull(work);

            limitPassed ??= token => Task.Delay(WaitLimit.Maximum, token);

            using var cancellation = new CancellationTokenSource();
            var context = new WaitContext(cancellation.Token, onReport);

            Task<T> task;

            try
            {
                task = work(context);
            }
            catch (Exception exception)
            {
                // Thrown before its first await - the same to the caller as thrown after it.
                task = Task.FromException<T>(exception);
            }

            while (true)
            {
                using var clock = new CancellationTokenSource();

                Task limit = limitPassed(clock.Token);
                Task activity = context.NextActivity;

                // The work first: when it and the limit are both complete, finishing wins.
                Task first = await Task.WhenAny(task, limit, activity).ConfigureAwait(true);

                clock.Cancel();

                if (first == task)
                    return WaitResult<T>.Finished(await task.ConfigureAwait(true), task);

                if (first == activity && !task.IsCompleted)
                    continue;

                if (task.IsCompleted)
                    return WaitResult<T>.Finished(await task.ConfigureAwait(true), task);

                // The limit passed with nothing happening. Stop the work where it listens, and never
                // leave its eventual exception unobserved - it is ours, not the caller's.
                cancellation.Cancel();
                WaitLimit.Observe(task);

                return WaitResult<T>.TimedOut(task);
            }
        }

        // The same for work that returns nothing.
        public static Task<WaitResult<bool>> RunAsync(
            Func<WaitContext, Task> work,
            Func<CancellationToken, Task>? limitPassed = null,
            Action<WaitReport>? onReport = null)
        {
            ArgumentNullException.ThrowIfNull(work);

            return WaitLimit.RunAsync<bool>(
                async context =>
                {
                    await work(context).ConfigureAwait(true);
                    return true;
                },
                limitPassed,
                onReport);
        }

        private static void Observe(Task task) =>
            _ = task.ContinueWith(
                finished => _ = finished.Exception,
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
    }

    // ###########################################################################################
    // What the work is handed: the token that is cancelled when the limit passes, and a way to say
    // what it is doing now - which also starts the two minutes again (see WaitLimit's header).
    //
    // Report may be called from any thread: the Feedback tab reports its upload's percentage from
    // the thread pool.
    // ###########################################################################################
    public sealed class WaitContext
    {
        private readonly Action<WaitReport>? thisOnReport;
        private readonly object thisGate = new();
        private TaskCompletionSource thisActivity = WaitContext.NewSignal();

        public WaitContext(CancellationToken token, Action<WaitReport>? onReport = null)
        {
            this.Token = token;
            this.thisOnReport = onReport;
        }

        public CancellationToken Token { get; }

        // A new sentence, and optionally how far along it is (0 to 1; null leaves the bar moving
        // without a position). Either way, it counts as something happening.
        public void Report(string? message, double? fraction = null)
        {
            TaskCompletionSource signal;

            lock (this.thisGate)
            {
                signal = this.thisActivity;
                this.thisActivity = WaitContext.NewSignal();
            }

            this.thisOnReport?.Invoke(new WaitReport(message, fraction));
            signal.TrySetResult();
        }

        // The same reports as an IProgress<string>, for work that already takes one (the Feedback
        // tab's packing and upload).
        public IProgress<string> AsProgress() => new Reporter(this);

        private sealed class Reporter(WaitContext context) : IProgress<string>
        {
            public void Report(string value) => context.Report(value);
        }

        // Completes at the next Report.
        internal Task NextActivity
        {
            get
            {
                lock (this.thisGate)
                    return this.thisActivity.Task;
            }
        }

        // Continuations run off the reporting thread, so a Report from inside a lock of the
        // caller's can never run WaitLimit's loop inline.
        private static TaskCompletionSource NewSignal() =>
            new(TaskCreationOptions.RunContinuationsAsynchronously);
    }

    public sealed record WaitReport(string? Message, double? Fraction);

    // ###########################################################################################
    // How a wait ended: the work's value, or TimedOut - in which case Value is default and Work is
    // the work itself, which a caller whose work cannot be stopped (a PDF being written) can keep
    // watching to say when it did finish.
    // ###########################################################################################
    public sealed class WaitResult<T>
    {
        private WaitResult(bool timedOut, T? value, Task<T> work)
        {
            this.IsTimedOut = timedOut;
            this.Value = value;
            this.Work = work;
        }

        public bool IsTimedOut { get; }

        public T? Value { get; }

        public Task<T> Work { get; }

        internal static WaitResult<T> Finished(T value, Task<T> work) => new(false, value, work);

        internal static WaitResult<T> TimedOut(Task<T> work) => new(true, default, work);
    }
}
