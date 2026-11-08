using System;
using System.Threading;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Headless;
using Avalonia.Styling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// Runs a test body on Avalonia's UI thread inside a headless session.
//
// Avalonia ships an xunit adapter ([AvaloniaFact]), but it was first kept out because it needed
// xunit v3 while this suite was on xunit 2, and now that the suite IS on v3 (4.x), the 12.1.3
// adapter is built against the 3.2 extensibility API that 4.0 changed. The session API
// underneath the adapter is public, so the tests use it directly and depend on no test
// framework at all.
//
// Everything touching a control must go through Run: Avalonia requires a dispatcher and
// will throw if a visual is created on an arbitrary thread.
//
// *** A TEST THAT CHANGES THE APPLICATION'S THEME FAILS, AND THE THEME IS PUT BACK (2026-10-04). ***
// The Application is built once per assembly (TestAppBuilder.cs), so a test that turned it dark
// left every later test dark - and a test comparing a drawn colour with the light theme then
// failed on CI, wherever the random order put it, never at the test that did it
// (SmallTabsTests' theme drop-down, through App.ApplyConfiguredTheme). Checked around every body,
// so the culprit is the test that fails.
// ###########################################################################################
public static class UiTest
{
    // One session per assembly. Starting it is expensive (it spins up the app, its styles
    // and its resource dictionaries), so it is created once and reused by every UI test.
    private static readonly HeadlessUnitTestSession Session =
        HeadlessUnitTestSession.GetOrStartForAssembly(typeof(UiTest).Assembly);

    public static void Run(Action body)
    {
        Session.Dispatch(
            () =>
            {
                ThemeVariant? before = UiTest.Theme;
                bool finished = false;

                try
                {
                    body();
                    finished = true;
                }
                finally
                {
                    UiTest.PutTheThemeBack(before, failTheTest: finished);
                }
            },
            CancellationToken.None).GetAwaiter().GetResult();
    }

    // ###########################################################################################
    // The same thing for a body that awaits.
    //
    // Use this rather than calling GetAwaiter().GetResult() on the body inside Run. Run's body
    // executes ON the dispatcher thread, so blocking it there blocks the dispatcher itself: any
    // code under test that awaits a Dispatcher.UIThread.InvokeAsync round-trip - which several
    // paths in TabOscilloscope do - would deadlock, because the continuation it waits for can only
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
    // exception, if any, on that pool thread. Pinned by UiTestTests.
    // ###########################################################################################
    public static async Task RunAsync(Func<Task> body)
    {
        Task dispatched = Session.Dispatch(
            async () =>
            {
                ThemeVariant? before = UiTest.Theme;
                bool finished = false;

                try
                {
                    await body();
                    finished = true;
                }
                finally
                {
                    UiTest.PutTheThemeBack(before, failTheTest: finished);
                }

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

    private static ThemeVariant? Theme => Application.Current?.RequestedThemeVariant;

    // The theme as it was before the body; a change fails the test, unless the body already failed
    // for its own reason - that failure is the one to see.
    private static void PutTheThemeBack(ThemeVariant? before, bool failTheTest)
    {
        ThemeVariant? after = UiTest.Theme;

        if (Equals(after, before) || Application.Current is not Application app)
            return;

        app.RequestedThemeVariant = before;

        if (failTheTest)
        {
            Assert.Fail(
                $"The test changed the shared Application's theme ({before} to {after}). Every later test would run in it - " +
                "use the control's seam instead (TabConfiguration.ApplyThemeOverrideForTests, for one).");
        }
    }
}

// UI tests share one dispatcher thread, so they are kept in a single xunit collection
// rather than being run in parallel against each other.
[CollectionDefinition("HeadlessUi")]
public class HeadlessUiCollection
{
}
