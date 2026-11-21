using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The oscilloscope as the component pane's baselines use it (IBaselineScope) - no network: its screen and what it says it
// is set to are what the test gives it, and every setting and key sent to it is noted.
// ###########################################################################################
internal sealed class FakeBaselineScope : IBaselineScope
{
    // A 1 x 1 PNG - enough for a thumbnail to be decoded from.
    internal static readonly byte[] Png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    // The screen a capture brings - null or empty for one that brings nothing.
    public byte[]? Screen { get; set; } = FakeBaselineScope.Png;

    // What the oscilloscope says it is set to.
    public ScopeSettingsReading Settings { get; set; } = new(null, null, null);

    public int Captures { get; private set; }

    public List<ComponentImageEntry> Applied { get; } = [];

    public List<int> TriggerSteps { get; } = [];

    public List<int> TimeDivSteps { get; } = [];

    // Set, a capture waits for it - a test does something while one is being taken.
    public TaskCompletionSource<byte[]?>? CaptureGate { get; set; }

    public Task<byte[]?> CaptureAsync(CancellationToken cancellationToken)
    {
        this.Captures++;
        return this.CaptureGate?.Task ?? Task.FromResult(this.Screen);
    }

    public Task<ScopeSettingsReading> ReadSettingsAsync(CancellationToken cancellationToken) => Task.FromResult(this.Settings);

    // Set, setting the oscilloscope for a pin waits for it - a capture waits for that in turn, with no "please wait" up yet.
    public TaskCompletionSource? ApplyGate { get; set; }

    public Task ApplyAsync(ComponentImageEntry baseline, CancellationToken cancellationToken)
    {
        this.Applied.Add(baseline);
        return this.ApplyGate?.Task ?? Task.CompletedTask;
    }

    public void StepTriggerLevel(int direction) => this.TriggerSteps.Add(direction);

    public void StepTimeDiv(int offset) => this.TimeDivSteps.Add(offset);

    public List<double> VoltsDivSet { get; } = [];

    public List<ScopeCommandPalette> Commands { get; } = [];

    public void SetVoltsDiv(double volts) => this.VoltsDivSet.Add(volts);

    public Task RunCommandAsync(ScopeCommandPalette command, CancellationToken cancellationToken)
    {
        this.Commands.Add(command);
        return Task.CompletedTask;
    }
}
