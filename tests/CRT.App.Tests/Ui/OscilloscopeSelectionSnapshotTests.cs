using Avalonia.Controls;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The oscilloscope's background workers read the selection the UI thread PUBLISHED, and never
// block on the UI thread for it (code review, 2026-10-13). The auto-connect loop took a snapshot
// every second by blocking a pool thread on the dispatcher until the UI thread read the controls.
//
// The test takes the snapshot on a pool thread WHILE the UI thread is blocked waiting for it - so
// the old code could only time out here, never answer.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class OscilloscopeSelectionSnapshotTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public OscilloscopeSelectionSnapshotTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void A_worker_takes_the_published_selection_without_waiting_for_the_UI_thread()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOscilloscope();
            tab.GetControl<TextBox>("HostTextBox").Text = "10.0.0.5";
            tab.GetControl<TextBox>("PortTextBox").Text = "5555";

            // TextChanged is raised from the dispatcher's queue; publishing follows it.
            Dispatcher.UIThread.RunJobs();

            var fromWorker = Task.Run(tab.TakeSelectionSnapshotForTests);

            Assert.True(fromWorker.Wait(TimeSpan.FromSeconds(10)), "the worker waited for the UI thread");
            Assert.Equal(("10.0.0.5", 5555), fromWorker.Result);
        });
    }

    // On the UI thread the snapshot reads the controls themselves - and publishes what it read.
    [Fact]
    public void A_snapshot_on_the_UI_thread_reads_the_controls_as_they_are()
    {
        UiTest.Run(() =>
        {
            var tab = new TabOscilloscope();
            tab.GetControl<TextBox>("HostTextBox").Text = "scope.local";

            Assert.Equal("scope.local", tab.TakeSelectionSnapshotForTests().Host);

            var fromWorker = Task.Run(tab.TakeSelectionSnapshotForTests);
            Assert.True(fromWorker.Wait(TimeSpan.FromSeconds(10)));
            Assert.Equal("scope.local", fromWorker.Result.Host);
        });
    }
}
