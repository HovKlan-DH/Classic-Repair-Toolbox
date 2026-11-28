using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;
using Handlers.Oscilloscope;

namespace Handlers.Collections
{
    // ###########################################################################################
    // A CAPTURE AS A COLLECTION KEEPS IT (2026-11-26) - made in the component info window from the image shown as the capture was
    // taken: its pin, signal name and expected reading; the oscilloscope's settings READ BACK from it (BaselineSettingWords spells
    // them as a baseline does, "50us", "500mV"), a value it did not answer left out; and what the window showed of its component
    // then (ComponentOf - owner request, 2026-11-27: the page shows "something like the component info window" beside a capture
    // seen large). The board's baseline for the pin is no longer kept (owner decision, 2026-11-27: "not relevant").
    // ###########################################################################################
    public static class CollectionCaptures
    {
        public static LocalCapture CaptureOf(
            ComponentImageEntry shown,
            ScopeSettingsReading? reading,
            CollectionComponent? component,
            string? oscilloscope,
            DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(shown);

            return new LocalCapture
            {
                ComponentLabel = shown.BoardLabel.Trim(),
                Pin = CollectionCaptures.Blank(shown.Pin),
                SignalName = CollectionCaptures.Blank(shown.Name),
                ExpectedReading = CollectionCaptures.Blank(shown.ExpectedOscilloscopeReading),
                TimePerDivision = BaselineSettingWords.TimeDiv(reading?.TimeDivSeconds),
                VoltsPerDivision = BaselineSettingWords.VoltsDiv(reading?.VoltsDivVolts),
                TriggerLevel = BaselineSettingWords.TriggerLevel(reading?.TriggerLevelVolts),
                Component = component,
                Oscilloscope = CollectionCaptures.Cut(oscilloscope, CollectionRules.MaximumNameLength),
                CapturedUtc = now
            };
        }

        // ###########################################################################################
        // THE OSCILLOSCOPE A CAPTURE IS TAKEN WITH (owner request, 2026-11-28: "the oscilloscope brand and model should be shown"):
        // its vendor and model as the "Oscilloscope" tab lists them - "Rigol" and "DS1054Z", "Rigol DS1054Z" - a model whose name
        // says its vendor already said once. Null with neither.
        // ###########################################################################################
        public static string? OscilloscopeNameOf(string? vendor, string? model)
        {
            string? brand = CollectionCaptures.Blank(vendor);
            string? series = CollectionCaptures.Blank(model);

            if (brand is null || series is null)
                return series ?? brand;

            return series.StartsWith(brand, StringComparison.OrdinalIgnoreCase) ? series : $"{brand} {series}";
        }

        // ###########################################################################################
        // THE COMPONENT AS THE WINDOW SHOWS IT: its title ("U7 | VIC-II | 6569"), category line, short description, what it is
        // there for as plain text (a note with formatting read as its words - the page shows no formatting), and its pins as the
        // window lists them, of the region shown - each held to what the server takes (CollectionRules), cut rather than
        // refused. Null with nothing to tell.
        // ###########################################################################################
        public static CollectionComponent? ComponentOf(
            string label,
            string? title,
            string? categoryLine,
            string? description,
            string? purposeMarkup,
            IEnumerable<ComponentPopupPin> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            List<CollectionPin> listed = pins
                .Where(pin => !string.IsNullOrWhiteSpace(pin.Pin))
                .Take(CollectionRules.MaximumPinsPerComponent)
                .Select(pin => new CollectionPin(
                    CollectionCaptures.Cut(pin.Pin, CollectionRules.MaximumSettingLength)!,
                    CollectionCaptures.Cut(pin.Name, CollectionRules.MaximumNameLength),
                    CollectionCaptures.Cut(pin.Direction, CollectionRules.MaximumSettingLength),
                    CollectionCaptures.Cut(pin.Net, CollectionRules.MaximumNameLength),
                    CollectionCaptures.Cut(pin.Reading, CollectionRules.MaximumReadingLength)))
                .ToList();

            string? purpose = string.IsNullOrWhiteSpace(purposeMarkup) ? null : RichNote.PlainOf(RichNote.TrimMarkup(purposeMarkup));

            var component = new CollectionComponent(
                label.Trim(),
                CollectionCaptures.Cut(title, CollectionRules.MaximumComponentTextLength),
                CollectionCaptures.Cut(categoryLine, CollectionRules.MaximumComponentTextLength),
                CollectionCaptures.Cut(description, CollectionRules.MaximumComponentTextLength),
                CollectionCaptures.Cut(purpose, CollectionRules.MaximumPurposeLength),
                listed.Count == 0 ? null : listed);

            return component is { Title: null, CategoryLine: null, Description: null, Purpose: null, Pins: null } ? null : component;
        }

        private static string? Blank(string? text) => string.IsNullOrWhiteSpace(text) ? null : text.Trim();

        // Trimmed and at most `length` characters, never half of one written as two; null for nothing.
        private static string? Cut(string? text, int length)
        {
            if (CollectionCaptures.Blank(text) is not { } trimmed)
                return null;

            if (trimmed.Length <= length)
                return trimmed;

            int cut = char.IsHighSurrogate(trimmed[length - 1]) ? length - 1 : length;

            return trimmed[..cut].TrimEnd();
        }
    }
}
