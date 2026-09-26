using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Headless;

namespace CRT.Maintainer.Tests.Ui;

// ###########################################################################################
// Runs a test body on Avalonia's UI thread inside a headless session - a copy of CRT.App.Tests'
// harness (2026-09-26), as this project's csproj asked for once it needed to build a control. The
// first need was the maintainer window itself: it failed to START when the table moved into it,
// and no test built it.
//
// Avalonia ships an xunit adapter ([AvaloniaFact]), but it was first kept out because it needed
// xunit v3 while this suite was on xunit 2, and now that the suite IS on v3 (4.x), the 12.1.3
// adapter is built against the 3.2 extensibility API that 4.0 changed. The session API
// underneath the adapter is public, so the tests use it directly and depend on no test
// framework at all.
//
// Everything touching a control must go through Run: Avalonia requires a dispatcher and
// will throw if a visual is created on an arbitrary thread.
// ###########################################################################################
public static class UiTest
{
    // One session per assembly. Starting it is expensive (it spins up the app, its styles
    // and its resource dictionaries), so it is created once and reused by every UI test.
    private static readonly HeadlessUnitTestSession Session =
        HeadlessUnitTestSession.GetOrStartForAssembly(typeof(UiTest).Assembly);

    public static void Run(Action body)
    {
        Session.Dispatch(body, CancellationToken.None).GetAwaiter().GetResult();
    }

    // ###########################################################################################
    // The same thing for a body that awaits.
    //
    // Use this rather than calling GetAwaiter().GetResult() on the body inside Run. Run's body
    // executes ON the dispatcher thread, so blocking it there blocks the dispatcher itself: any
    // code under test that awaits a Dispatcher.UIThread.InvokeAsync round-trip - would deadlock, because the continuation it waits for can only
    // run on the thread the block is holding. Dispatch's async overload keeps pumping instead, so
    // those round-trips complete.
    //
    // THE TEST MUST NOT CONTINUE ON THE DISPATCHER THREAD AFTERWARDS, which a plain await of
    // Dispatch does under xunit v3. The session completes that task from its own dispatcher thread,
    // with no RunContinuationsAsynchronously, so the continuation runs inline there - and so does
    // everything after it, xunit included, which goes on to run the NEXT test on that thread. The
    // next UiTest.Run then queues its body to the thread it is itself blocking, and the whole run
    // hangs, idle. xunit 2 masked this by running every test under its own SynchronizationContext,
    // which the await captured; v3 has none. So the continuation is moved to the thread pool
    // explicitly: ContinueWith without ExecuteSynchronously queues to TaskScheduler.Default, never
    // inline, whether the body succeeded or threw. The second await then rethrows the body's
    // exception, if any, on that pool thread. Pinned by CRT.App.Tests' UiTestTests.
    // ###########################################################################################
    public static async Task RunAsync(Func<Task> body)
    {
        Task dispatched = Session.Dispatch(
            async () =>
            {
                await body();
                return true;
            },
            CancellationToken.None
        );

        await dispatched.ContinueWith(
            static _ => { },
            CancellationToken.None,
            TaskContinuationOptions.None,
            TaskScheduler.Default);

        await dispatched;
    }
}

// UI tests share one dispatcher thread, so they are kept in a single xunit collection
// rather than being run in parallel against each other.
[CollectionDefinition("HeadlessUi")]
public class HeadlessUiCollection
{
}
