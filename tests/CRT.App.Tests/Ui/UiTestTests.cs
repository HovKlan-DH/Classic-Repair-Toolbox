using Avalonia.Controls;

namespace ClassicRepairToolbox.Tests.Ui;

// The headless harness itself - specifically what happens AFTER a UiTest.RunAsync body finishes.
//
// WHY THIS EXISTS: moving to xunit v3 made the whole CRT.App.Tests run hang forever, idle, on an
// ordinary synchronous UI test. The session completes the task RunAsync awaits from its OWN
// dispatcher thread, with no RunContinuationsAsynchronously, so whatever runs after that await runs
// right there, inline, unless something redirects it. xunit 2 did: it ran every test under a
// SynchronizationContext of its own, which the await captured. xunit v3 does not, so the rest of
// the test - and then xunit itself, carrying on to the NEXT test - ran on the session's one
// dispatcher thread. The next UiTest.Run queued its body to that thread and blocked it waiting for
// the result: a thread waiting on itself. Which test hung depended only on which synchronous UI
// test xunit v3's randomised order put straight after an async one.
[Collection("HeadlessUi")]
public class UiTestTests
{
    // The cause, asserted directly: after RunAsync returns, the test is no longer on the thread
    // that ran the body. Fails against the version that simply awaited Session.Dispatch.
    [Fact]
    public async Task After_RunAsync_the_test_continues_off_the_headless_dispatcher_thread()
    {
        int dispatcherThreadId = -1;

        await UiTest.RunAsync(async () =>
        {
            dispatcherThreadId = Environment.CurrentManagedThreadId;
            await Task.Yield();
        });

        Assert.NotEqual(-1, dispatcherThreadId);
        Assert.NotEqual(dispatcherThreadId, Environment.CurrentManagedThreadId);
    }

    // The symptom, asserted with a timeout rather than left to hang: a synchronous UI body can run
    // straight after an async one. Run from Task.Run, so that on a broken harness the thread that
    // blocks is a pool thread, and this test fails after the timeout instead of stopping the suite.
    [Fact]
    public async Task A_synchronous_UI_body_runs_straight_after_an_async_one()
    {
        await UiTest.RunAsync(async () =>
        {
            _ = new TextBlock();
            await Task.Yield();
        });

        bool ran = false;

        // BLOCKING on purpose, which is what xUnit1031 objects to: it is what the next synchronous
        // UI test does. An await here would hand a hijacked dispatcher thread back to the session,
        // which would then run the queued body - and this test would pass on the broken harness.
#pragma warning disable xUnit1031
        bool finished = Task.Run(() => UiTest.Run(() => ran = new TextBlock() != null))
            .Wait(TimeSpan.FromSeconds(30));
#pragma warning restore xUnit1031

        Assert.True(finished, "UiTest.Run did not complete within 30 seconds of a UiTest.RunAsync body.");
        Assert.True(ran);
    }
}
