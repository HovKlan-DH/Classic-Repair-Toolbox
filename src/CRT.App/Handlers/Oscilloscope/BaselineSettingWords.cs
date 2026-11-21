using Handlers.DataHandling;
using System;
using System.Globalization;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // WHAT THE OSCILLOSCOPE SAYS IT IS SET TO, WRITTEN AS AN OSCILLOSCOPE BASELINE STORES IT (owner decision, 2026-11-18: a
    // capture reads the settings back from the oscilloscope, so the baseline's settings match its image - "the picture shows the
    // scope's real settings, so the row should store those"). The published data's spellings, counted on the converted trees
    // that day: T/DIV "50us", "1ms", "500ns" (ComponentVocabulary.TimeDivisions' units), V/DIV "1V", "500mV", a trigger level
    // always in volts, "1.5V", "0.75V". A value the oscilloscope did not answer, or one no setting can be, is null - the
    // baseline then keeps what it had.
    // ###########################################################################################
    public static class BaselineSettingWords
    {
        public static string? TimeDiv(double? seconds) =>
            seconds is double value && double.IsFinite(value) && value > 0
                ? BaselineSettingWords.InUnits(value, [(1, "s"), (1e-3, "ms"), (1e-6, "us"), (1e-9, "ns")])
                : null;

        public static string? VoltsDiv(double? volts) =>
            volts is double value && double.IsFinite(value) && value > 0
                ? BaselineSettingWords.InUnits(value, [(1, "V"), (1e-3, "mV")])
                : null;

        // A trigger level may be at or below 0 V - always in volts, as every published one is.
        public static string? TriggerLevel(double? volts) =>
            volts is double value && double.IsFinite(value)
                ? BaselineSettingWords.Number(value) + "V"
                : null;

        // The largest unit the value is at least one of - a hair under it counting (5e-05 reads back as 4.9999999999e-05).
        private static string InUnits(double value, (double Size, string Name)[] units)
        {
            foreach ((double size, string name) in units)
            {
                if (value >= size * (1 - 1e-9))
                {
                    return BaselineSettingWords.Number(value / size) + name;
                }
            }

            (double smallest, string smallestName) = units[^1];
            return BaselineSettingWords.Number(value / smallest) + smallestName;
        }

        private static string Number(double value)
        {
            string text = Math.Round(value, 3, MidpointRounding.AwayFromZero).ToString("0.###", CultureInfo.InvariantCulture);
            return text == "-0" ? "0" : text;
        }

        // ###########################################################################################
        // THE BASELINE A CAPTURE MAKES - ONE RULE for what "Capture from oscilloscope pin-by-pin" shows before it is accepted and what
        // the pane puts into the draft once it is (code review, 2026-11-21: the window showed a new pin's baseline as "All regions",
        // and the pane put it in the region shown). `baseline` as the pane holds it, with the image `file` and what the oscilloscope
        // said it is set to - a setting it did not say keeping the baseline's - or, with none, a new baseline of `pin` named `signal`,
        // measured in `machineState`, in `region` (where a new baseline goes: the region shown on a board with regions, else all).
        // ###########################################################################################
        public static ComponentImageEntry Captured(
            ComponentImageEntry? baseline,
            ScopeSettingsReading settings,
            string file,
            string pin,
            string signal,
            string machineState,
            string region)
        {
            ArgumentNullException.ThrowIfNull(settings);

            ComponentImageEntry from = baseline ?? new ComponentImageEntry
            {
                Region = region ?? string.Empty,
                Pin = pin ?? string.Empty,
                Name = signal ?? string.Empty,
                MachineState = (machineState ?? string.Empty).Trim(),
            };

            return from.WithCapture(
                file ?? string.Empty,
                BaselineSettingWords.TimeDiv(settings.TimeDivSeconds) ?? from.TimeDiv,
                BaselineSettingWords.VoltsDiv(settings.VoltsDivVolts) ?? from.VoltsDiv,
                BaselineSettingWords.TriggerLevel(settings.TriggerLevelVolts) ?? from.TriggerLevelVolts);
        }
    }

    // What the oscilloscope said it is set to - each null when it was not asked (its model has no such command) or not answered.
    public sealed record ScopeSettingsReading(double? TimeDivSeconds, double? VoltsDivVolts, double? TriggerLevelVolts);
}
