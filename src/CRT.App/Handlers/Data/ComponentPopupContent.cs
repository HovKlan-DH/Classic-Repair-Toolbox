using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // One pin as the component popup lists it - with its net on this board, from the board's KiCad
    // data ("" without), and whether the image shown is of it (owner request, 2026-11-02).
    public sealed record ComponentPopupPin(string Pin, string Name, string Direction, string Net = "", bool IsViewed = false)
    {
        public bool HasNet => this.Net.Length > 0;
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
        // its row says so (owner request, 2026-11-02).
        // ###########################################################################################
        public static List<ComponentPopupPin> PinsOf(
            IEnumerable<ComponentPinEntry> pins,
            string? label,
            string? region,
            IReadOnlyDictionary<string, string>? nets = null,
            string? viewedPin = null)
        {
            ArgumentNullException.ThrowIfNull(pins);

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
                        ComponentVocabulary.IsSamePin(number, viewedPin));
                })
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
