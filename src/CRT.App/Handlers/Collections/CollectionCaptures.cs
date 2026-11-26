using System;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace Handlers.Collections
{
    // ###########################################################################################
    // A CAPTURE AS A COLLECTION KEEPS IT (2026-11-26) - made in the component info window from the image shown as the capture was
    // taken: its pin, signal name and expected reading; the oscilloscope's settings READ BACK from it (BaselineSettingWords spells
    // them as a baseline does, "50us", "500mV"), a value it did not answer left out; and, when the image shown was the board's
    // baseline for the pin, that baseline - its picture by checksum and file name, and how it was measured - for the page to
    // show beside it.
    // ###########################################################################################
    public static class CollectionCaptures
    {
        public static LocalCapture CaptureOf(
            ComponentImageEntry shown,
            ScopeSettingsReading? reading,
            string? baselineSha256,
            DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(shown);

            bool isBaseline = ComponentImageQueries.IsOscilloscopeImage(shown);

            return new LocalCapture
            {
                ComponentLabel = shown.BoardLabel.Trim(),
                Pin = CollectionCaptures.Blank(shown.Pin),
                SignalName = CollectionCaptures.Blank(shown.Name),
                ExpectedReading = CollectionCaptures.Blank(shown.ExpectedOscilloscopeReading),
                TimePerDivision = BaselineSettingWords.TimeDiv(reading?.TimeDivSeconds),
                VoltsPerDivision = BaselineSettingWords.VoltsDiv(reading?.VoltsDivVolts),
                TriggerLevel = BaselineSettingWords.TriggerLevel(reading?.TriggerLevelVolts),
                BaselineSha256 = isBaseline && DataTreeLayout.IsSha256(baselineSha256) ? baselineSha256 : null,
                BaselineFileName = isBaseline ? CollectionCaptures.Blank(shown.File) : null,
                BaselineTimePerDivision = isBaseline ? CollectionCaptures.Blank(shown.TimeDiv) : null,
                BaselineVoltsPerDivision = isBaseline ? CollectionCaptures.Blank(shown.VoltsDiv) : null,
                BaselineTriggerLevel = isBaseline ? CollectionCaptures.Blank(shown.TriggerLevelVolts) : null,
                CapturedUtc = now
            };
        }

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();
    }
}
