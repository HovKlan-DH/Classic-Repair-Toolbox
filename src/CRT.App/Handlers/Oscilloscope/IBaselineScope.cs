using Handlers.DataHandling;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // THE OSCILLOSCOPE AS THE COMPONENT PANE'S BASELINES USE IT - the "Oscilloscope" tab's session (TabOscilloscope.Baselines.cs),
    // a fake in a test. "Capture again" on a baseline (ComponentPane.BaselineRun.cs) and "Capture from oscilloscope pin-by-pin"
    // (BaselineCaptureWindow).
    // ###########################################################################################
    internal interface IBaselineScope
    {
        // The screen as PNG bytes - null, the reason in the "Oscilloscope" tab's output.
        Task<byte[]?> CaptureAsync(CancellationToken cancellationToken);

        // What it is set to - each value null when its model cannot ask it or it did not answer.
        Task<ScopeSettingsReading> ReadSettingsAsync(CancellationToken cancellationToken);

        // Set as `baseline` was measured: its T/DIV, V/DIV and trigger level, those it has - sent every time it is asked, though they
        // were the last sent, as the knobs may have been turned by hand since (code review, 2026-11-21).
        Task ApplyAsync(ComponentImageEntry baseline, CancellationToken cancellationToken);

        // The component info window's keys (Controlling-oscilloscope-with-keyboard): the trigger level a step up (+1) or down
        // (-1), the time per division a step through its model's list - the popup's "+" key sends -1, its "-" key +1 - the
        // volts per division set to a value of its model's list, and a command of the model's run (clear statistics, single, run).
        void StepTriggerLevel(int direction);

        void StepTimeDiv(int offset);

        void SetVoltsDiv(double volts);

        Task RunCommandAsync(ScopeCommandPalette command, CancellationToken cancellationToken);
    }
}
