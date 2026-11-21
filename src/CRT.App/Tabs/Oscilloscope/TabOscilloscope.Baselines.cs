using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE OSCILLOSCOPE FOR THE COMPONENT PANE'S BASELINES (IBaselineScope; owner requests, 2026-11-18 and 2026-11-20 - "Capture from
    // oscilloscope pin-by-pin"): the screen captured, the oscilloscope set as a baseline was measured, the popup's keys - each the tab's own way, over
    // its one session - and, new, WHAT IT IS SET TO READ BACK (owner decision, the same day: the baseline stores the settings
    // its image was taken at). See TabOscilloscope.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class TabOscilloscope : IBaselineScope
    {
        Task<byte[]?> IBaselineScope.CaptureAsync(CancellationToken cancellationToken) =>
            this.CaptureOscilloscopePngAsync(cancellationToken);

        Task<ScopeSettingsReading> IBaselineScope.ReadSettingsAsync(CancellationToken cancellationToken) =>
            this.ReadScopeSettingsAsync(cancellationToken);

        // Every time it is asked, though these very settings were sent last: knobs turned by hand since are no change the tab can
        // see (code review, 2026-11-21).
        Task IBaselineScope.ApplyAsync(ComponentImageEntry baseline, CancellationToken cancellationToken) =>
            this.ApplyComponentImageOscilloscopeSettingsAsync(baseline, cancellationToken, evenIfSentLast: true);

        void IBaselineScope.StepTriggerLevel(int direction) => this.QueueTriggerLevelKeyboardStep(direction);

        void IBaselineScope.StepTimeDiv(int offset) => this.QueueTimeDivKeyboardStep(offset);

        void IBaselineScope.SetVoltsDiv(double volts) => this.QueueVoltsDivKeyboardSet(volts);

        Task IBaselineScope.RunCommandAsync(ScopeCommandPalette command, CancellationToken cancellationToken) =>
            this.RunPaletteAsync(command, cancellationToken);

        // ###########################################################################################
        // T/DIV, V/DIV and the trigger level as the oscilloscope answers them, over the session - each asked only when its model
        // has the command (a palette with no command text ends the session, ExecutePaletteAsync), and null when it was not asked
        // or not answered. What was held of each before is let go of first, so an answer that cannot be read is no answer -
        // the keyboard steppers read the oscilloscope again then, as they do when they hold nothing.
        // ###########################################################################################
        internal async Task<ScopeSettingsReading> ReadScopeSettingsAsync(CancellationToken cancellationToken)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot = Dispatcher.UIThread.CheckAccess()
                ? this.CreateOscilloscopeSelectionSnapshot()
                : await Dispatcher.UIThread.InvokeAsync(this.CreateOscilloscopeSelectionSnapshot, DispatcherPriority.Background, cancellationToken);

            double? timeDiv = null;
            double? voltsDiv = null;
            double? triggerLevel = null;

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                selectionSnapshot,
                async (scopeClient, oscilloscopeEntry, token) =>
                {
                    if (TabOscilloscope.HasCommand(oscilloscopeEntry, ScopeCommand.QueryTimeDiv))
                    {
                        this.thisLastTimeDivSeconds = null;
                        await this.ExecutePaletteAsync(scopeClient, oscilloscopeEntry, ScopeCommandPalette.QueryTimeDiv, token).ConfigureAwait(false);
                        timeDiv = this.thisLastTimeDivSeconds;
                    }

                    if (TabOscilloscope.HasCommand(oscilloscopeEntry, ScopeCommand.QueryVoltsDiv))
                    {
                        this.thisLastVoltsDivVolts = null;
                        await this.ExecutePaletteAsync(scopeClient, oscilloscopeEntry, ScopeCommandPalette.QueryVoltsDiv, token).ConfigureAwait(false);
                        voltsDiv = this.thisLastVoltsDivVolts;
                    }

                    if (TabOscilloscope.HasCommand(oscilloscopeEntry, ScopeCommand.QueryTriggerLevel))
                    {
                        this.thisLastTriggerLevelVolts = null;
                        await this.ExecutePaletteAsync(scopeClient, oscilloscopeEntry, ScopeCommandPalette.QueryTriggerLevel, token).ConfigureAwait(false);
                        triggerLevel = this.thisLastTriggerLevelVolts;
                    }
                },
                cancellationToken,
                writeWarnings: false).ConfigureAwait(false);

            return new ScopeSettingsReading(timeDiv, voltsDiv, triggerLevel);
        }

        private static bool HasCommand(OscilloscopeEntry oscilloscopeEntry, ScopeCommand command) =>
            !string.IsNullOrWhiteSpace(ScopeCommandResolver.GetCommandText(oscilloscopeEntry, command));
    }
}
