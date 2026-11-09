using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // One pin as the component popup lists it - with its net on this board, from the board's KiCad
    // data ("" without), whether the image shown is of it (owner request, 2026-11-02), what its
    // oscilloscope baselines say the scope should read on it ("" for none - owner request, 2026-11-09),
    // and whether the pointer is on it in the drawing beside the list (owner request, 2026-11-09).
    public sealed record ComponentPopupPin(
        string Pin,
        string Name,
        string Direction,
        string Net = "",
        bool IsViewed = false,
        string Reading = "",
        bool IsPointedAt = false)
    {
        public bool HasNet => this.Net.Length > 0;

        public bool HasReading => this.Reading.Length > 0;
    }

    // One replacement part as the component popup lists it - its part number and kind on the first line.
    public sealed record ComponentPopupReplacement(string Title, string Note)
    {
        public bool HasNote => this.Note.Length > 0;
    }

    // ###########################################################################################
    // WHAT THE COMPONENT POPUP SHOWS OF CRT 4.0's NEW COMPONENT DATA (step 2, slice 3c, 2026-10-26 -
    // the plan: "Every part has to change: ... the popup and the Overview tab"): the package beside the
    // category and part number, the pins and the replacement parts of the region the popup shows, and
    // what a scope baseline was measured under. Words and lists only, so they are tested without a
    // window; ComponentInfoWindow only shows them.
    // ###########################################################################################
    internal static class ComponentPopupContent
    {
        // ###########################################################################################
        // "U8 | CPU | 6510" - what is filled of the label, the friendly name and the technical name or
        // value. The popup's title, and the headline of the component in the Contribute tab (owner request,
        // 2026-11-04: "should be identical to what is shown in the component info popup"). "" with nothing
        // filled.
        // ###########################################################################################
        public static string TitleOf(string? label, ComponentEntry? entry) =>
            string.Join(
                " | ",
                new[] { label, entry?.FriendlyName, entry?.TechnicalNameOrValue }
                    .Select(part => part?.Trim() ?? string.Empty)
                    .Where(part => part.Length > 0));

        // "IC | 6510 | DIP-40" - what is filled of category, part number and package, in that order.
        public static string CategoryLine(ComponentEntry? entry) =>
            string.Join(
                " | ",
                new[] { entry?.Category, entry?.PartNumber, entry?.Package }
                    .Select(part => part?.Trim() ?? string.Empty)
                    .Where(part => part.Length > 0));

        // ###########################################################################################
        // The component's pins in `region` (a pin with no region is in every one), in number order -
        // each with its net on this board from `nets` (by pin number; the plan: "The popup shows the
        // signal name and 'net on this board' side by side"). `viewedPin`: the pin of the image shown -
        // its row says so (owner request, 2026-11-02). `images`: the board's component images, whose
        // oscilloscope baselines of this component and region give each pin its expected reading
        // (ReadingOf). `pointedPin`: the pin the pointer is on in the drawing - its row says so.
        // ###########################################################################################
        public static List<ComponentPopupPin> PinsOf(
            IEnumerable<ComponentPinEntry> pins,
            string? label,
            string? region,
            IReadOnlyDictionary<string, string>? nets = null,
            string? viewedPin = null,
            IEnumerable<ComponentImageEntry>? images = null,
            int? pointedPin = null)
        {
            ArgumentNullException.ThrowIfNull(pins);

            // The component's baselines in the region, found once for all its pins.
            List<ComponentImageEntry> baselines = (images ?? [])
                .Where(image => ComponentImageKinds.IsScopeBaseline(image) && ComponentPopupContent.InScope(image.BoardLabel, image.Region, label, region))
                .ToList();

            return pins
                .Where(pin => ComponentPopupContent.InScope(pin.BoardLabel, pin.Region, label, region))
                .OrderBy(pin => pin.Pin, ComponentPaneEdits.PinOrder)
                .Select(pin =>
                {
                    string number = pin.Pin?.Trim() ?? string.Empty;
                    string net = nets is not null && nets.TryGetValue(number, out string? found) ? found : string.Empty;

                    return new ComponentPopupPin(
                        number,
                        pin.Name?.Trim() ?? string.Empty,
                        pin.Direction?.Trim() ?? string.Empty,
                        net,
                        ComponentVocabulary.IsSamePin(number, viewedPin),
                        ComponentPopupContent.ReadingOf(baselines, number),
                        pointedPin is int pointed && ComponentVocabulary.PinNumberOf(number) == pointed);
                })
                .ToList();
        }

        // ###########################################################################################
        // WHAT THE SCOPE SHOULD READ ON `pin` (owner request, 2026-11-09: "add another column ... named
        // 'Expected oscilloscope reading' and have the value from the oscilloscope pin"): the expected
        // reading of its baselines among `baselines`, "2 [64K]" being pin 2. Nearly every pin has one
        // baseline (measured 2026-11-09: 4 of the 13,858 pins with one have two that read differently,
        // the C16's "2" HIGH and "2 [64K]" Pulsing), so one reading is the reading; readings that
        // differ are each a line, "Pulsing [64K]" - with what its pin says beside its number, unless a
        // baseline of the plain pin reads it too. Capitals aside, a reading is said once. "" for none.
        // ###########################################################################################
        internal static string ReadingOf(IEnumerable<ComponentImageEntry> baselines, string? pin)
        {
            ArgumentNullException.ThrowIfNull(baselines);

            List<IGrouping<string, ComponentImageEntry>> readings = baselines
                .Where(baseline => ComponentVocabulary.IsSamePin(baseline.Pin, pin) &&
                                   !string.IsNullOrWhiteSpace(baseline.ExpectedOscilloscopeReading))
                .GroupBy(baseline => baseline.ExpectedOscilloscopeReading.Trim(), StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (readings.Count <= 1)
            {
                return readings.FirstOrDefault()?.Key ?? string.Empty;
            }

            return string.Join(
                "\n",
                readings.Select(reading =>
                {
                    List<string> asides = reading.Select(baseline => ComponentPopupContent.BesideItsNumber(baseline.Pin)).ToList();

                    return asides.Any(aside => aside.Length == 0)
                        ? reading.Key
                        : $"{reading.Key} {string.Join(" ", asides.Distinct(StringComparer.OrdinalIgnoreCase))}";
                }));
        }

        // What a baseline's pin says beside its number - "[64K]" of "2 [64K]" - "" for a pin that is only its
        // number, or no number at all.
        private static string BesideItsNumber(string? pin)
        {
            string text = (pin ?? string.Empty).Trim();

            return ComponentVocabulary.PinNumberOf(text) is null
                ? string.Empty
                : text[text.TakeWhile(char.IsAsciiDigit).Count()..].Trim();
        }

        // The pins as the drawing takes them - the same scope.
        public static List<ComponentPinEntry> PinEntriesOf(IEnumerable<ComponentPinEntry> pins, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(pins);

            return pins.Where(pin => ComponentPopupContent.InScope(pin.BoardLabel, pin.Region, label, region)).ToList();
        }

        // ###########################################################################################
        // The component's replacement parts in `region`, as the board lists them: "8500 - Drop-in replacement" (the
        // part number alone without a kind, the kind alone without a part number) and its note.
        // ###########################################################################################
        public static List<ComponentPopupReplacement> ReplacementsOf(IEnumerable<ComponentReplacementEntry> replacements, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(replacements);

            return replacements
                .Where(row => ComponentPopupContent.InScope(row.BoardLabel, row.Region, label, region))
                .Select(row => new ComponentPopupReplacement(
                    string.Join(" - ", new[] { row.PartNumber, row.Kind }.Select(part => part?.Trim() ?? string.Empty).Where(part => part.Length > 0)),
                    row.Note?.Trim() ?? string.Empty))
                .Where(row => row.Title.Length > 0 || row.HasNote)
                .ToList();
        }

        // ###########################################################################################
        // What a scope baseline was measured under, for the popup's line under the picture - its machine
        // state - or null when it does not say (every baseline converted from the Excel data, and every
        // picture).
        // ###########################################################################################
        public static string? BaselineConditions(ComponentImageEntry? image) =>
            string.IsNullOrWhiteSpace(image?.MachineState) ? null : $"Machine state: {image.MachineState.Trim()}";

        // ###########################################################################################
        // Which of the popup's images goes with pin `pin` (owner request, 2026-11-01: "the pins clickable
        // ... so the user quickly can go to that pin and see what images are available for that pin"): the
        // first whose pin is it, "2 [64K]" being pin 2 (ComponentVocabulary.PinNumberOf) - images of a pin
        // stand together. -1 for none.
        // ###########################################################################################
        public static int FirstImageOfPin(IReadOnlyList<string> imagePins, string? pin)
        {
            ArgumentNullException.ThrowIfNull(imagePins);

            for (int index = 0; index < imagePins.Count; index++)
            {
                if (ComponentVocabulary.IsSamePin(imagePins[index], pin))
                {
                    return index;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // The component's files and links in `region` (owner request, 2026-11-02: each can be of one
        // region): by the same scope as the pins, so the popup and the pane agree.
        // ###########################################################################################
        public static List<ComponentLocalFileEntry> FilesOf(IEnumerable<ComponentLocalFileEntry> files, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(files);

            return files.Where(file => ComponentPopupContent.InScope(file.BoardLabel, file.Region, label, region)).ToList();
        }

        public static List<ComponentLinkEntry> LinksOf(IEnumerable<ComponentLinkEntry> links, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(links);

            return links.Where(link => ComponentPopupContent.InScope(link.BoardLabel, link.Region, label, region)).ToList();
        }

        // Said over the image when a pin clicked has none - "Not found" is what typing a pin says.
        public static string NoImageOfPin(string? pin) => $"No image of pin {pin?.Trim()}";

        // The pane's own rule, so the popup and the pane agree on which rows are the component's
        // (code review, 2026-10-26 - a copy of it was here).
        private static bool InScope(string? rowLabel, string? rowRegion, string? label, string? region) =>
            ComponentPaneEdits.InScope(rowLabel, rowRegion, label, region);
    }
}
