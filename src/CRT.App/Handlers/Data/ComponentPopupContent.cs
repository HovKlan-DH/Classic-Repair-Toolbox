using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // One pin as the component popup lists it.
    public sealed record ComponentPopupPin(string Pin, string Name, string Direction);

    // One replacement part as the component popup lists it - its part number and kind on the first line.
    public sealed record ComponentPopupReplacement(string Title, string Changes, string Source)
    {
        public bool HasChanges => this.Changes.Length > 0;

        public bool HasSource => this.Source.Length > 0;
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
        // "IC | 6510 | DIP-40" - what is filled of category, part number and package, in that order.
        public static string CategoryLine(ComponentEntry? entry) =>
            string.Join(
                " | ",
                new[] { entry?.Category, entry?.PartNumber, entry?.Package }
                    .Select(part => part?.Trim() ?? string.Empty)
                    .Where(part => part.Length > 0));

        // The component's pins in `region` (a pin with no region is in every one), in number order.
        public static List<ComponentPopupPin> PinsOf(IEnumerable<ComponentPinEntry> pins, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(pins);

            return pins
                .Where(pin => ComponentPopupContent.InScope(pin.BoardLabel, pin.Region, label, region))
                .OrderBy(pin => pin.Pin, ComponentPaneEdits.PinOrder)
                .Select(pin => new ComponentPopupPin(pin.Pin?.Trim() ?? string.Empty, pin.Name?.Trim() ?? string.Empty, pin.Direction?.Trim() ?? string.Empty))
                .ToList();
        }

        // The pins as the drawing takes them - the same scope.
        public static List<ComponentPinEntry> PinEntriesOf(IEnumerable<ComponentPinEntry> pins, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(pins);

            return pins.Where(pin => ComponentPopupContent.InScope(pin.BoardLabel, pin.Region, label, region)).ToList();
        }

        // ###########################################################################################
        // The component's replacement parts in `region`, as the board lists them: "8500 - Drop-in" (the
        // part number alone without a kind, the kind alone without a part number), what to change, and
        // where the advice comes from.
        // ###########################################################################################
        public static List<ComponentPopupReplacement> ReplacementsOf(IEnumerable<ComponentReplacementEntry> replacements, string? label, string? region)
        {
            ArgumentNullException.ThrowIfNull(replacements);

            return replacements
                .Where(row => ComponentPopupContent.InScope(row.BoardLabel, row.Region, label, region))
                .Select(row => new ComponentPopupReplacement(
                    string.Join(" - ", new[] { row.PartNumber, row.Kind }.Select(part => part?.Trim() ?? string.Empty).Where(part => part.Length > 0)),
                    row.Changes?.Trim() ?? string.Empty,
                    string.IsNullOrWhiteSpace(row.Source) ? string.Empty : $"Comes from: {row.Source.Trim()}"))
                .Where(row => row.Title.Length > 0 || row.HasChanges || row.HasSource)
                .ToList();
        }

        // ###########################################################################################
        // What a scope baseline was measured under, for the popup's line under the picture - its
        // machine state and where it comes from - or null when it says neither (every baseline
        // converted from the Excel data, and every picture).
        // ###########################################################################################
        public static string? BaselineConditions(ComponentImageEntry? image)
        {
            if (image is null)
            {
                return null;
            }

            var parts = new List<string>(2);

            if (!string.IsNullOrWhiteSpace(image.MachineState))
            {
                parts.Add($"Machine state: {image.MachineState.Trim()}");
            }

            if (!string.IsNullOrWhiteSpace(image.Source))
            {
                parts.Add($"Comes from: {image.Source.Trim()}");
            }

            return parts.Count == 0 ? null : string.Join("\n", parts);
        }

        // The pane's own rule, so the popup and the pane agree on which rows are the component's
        // (code review, 2026-10-26 - a copy of it was here).
        private static bool InScope(string? rowLabel, string? rowRegion, string? label, string? region) =>
            ComponentPaneEdits.InScope(rowLabel, rowRegion, label, region);
    }
}
