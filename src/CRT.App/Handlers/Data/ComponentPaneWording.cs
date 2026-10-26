using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT THE COMPONENT PANE SAYS (CRT 4.0 step 2, 2026-10-25) - its title, the region line, the
    // offer to add the case's pins, and what is still edited in the component editor. Words only, so
    // they are tested without a display.
    // ###########################################################################################
    internal static class ComponentPaneWording
    {
        // "U8 - VIC-II | 6569" - the label, then what the component list shows beside it.
        public static string Title(ComponentEntry row)
        {
            ArgumentNullException.ThrowIfNull(row);

            string label = row.BoardLabel?.Trim() ?? string.Empty;
            string[] names = new[] { row.FriendlyName, row.TechnicalNameOrValue }
                .Select(name => name?.Trim() ?? string.Empty)
                .Where(name => name.Length > 0)
                .ToArray();

            return names.Length == 0 ? label : $"{label} - {string.Join(" | ", names)}";
        }

        // Said only for a component with a row of its own per region - which one is open.
        public static string RegionLine(string? rowRegion)
        {
            string region = rowRegion?.Trim() ?? string.Empty;

            return region.Length == 0
                ? string.Empty
                : $"The {region} version - each region has a row of its own, and this is the one CRT shows now.";
        }

        public static string NotOnBoard(string label) =>
            $"[{label}] is not on the board in your draft any more. If that was not meant, Undo in the draft's header brings it back.";

        // ###########################################################################################
        // What the component editor window still edits for this component: its pictures and scope
        // baselines in the region shown, and its files.
        // ###########################################################################################
        public static string Elsewhere(BoardData board, string label, string region)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            List<ComponentImageEntry> images = board.ComponentImages
                .Where(row => string.Equals(row.BoardLabel?.Trim(), wanted, StringComparison.OrdinalIgnoreCase) &&
                              (string.IsNullOrWhiteSpace(row.Region) || string.Equals(row.Region.Trim(), region?.Trim(), StringComparison.OrdinalIgnoreCase)))
                .ToList();

            int baselines = images.Count(ComponentImageKinds.IsScopeBaseline);
            int pictures = images.Count - baselines;
            int files = board.ComponentLocalFiles.Count(row => string.Equals(row.BoardLabel?.Trim(), wanted, StringComparison.OrdinalIgnoreCase));

            if (pictures + baselines + files == 0)
            {
                return "This component has no pictures, scope baselines or files yet. They are added in the component editor for now.";
            }

            var parts = new List<string>(3);

            if (pictures > 0)
                parts.Add(ComponentPaneWording.Count(pictures, "picture", "pictures"));
            if (baselines > 0)
                parts.Add(ComponentPaneWording.Count(baselines, "scope baseline", "scope baselines"));
            if (files > 0)
                parts.Add(ComponentPaneWording.Count(files, "file", "files"));

            string listed = parts.Count == 1
                ? parts[0]
                : $"{string.Join(", ", parts.Take(parts.Count - 1))} and {parts[^1]}";

            return $"This component has {listed}. They are edited in the component editor for now.";
        }

        // How many of a case's pin numbers the pins do not list yet - 0 for a case whose count CRT
        // cannot tell.
        public static int MissingCasePins(string? package, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            if (ComponentVocabulary.PinCountOf(package) is not int count)
            {
                return 0;
            }

            var listed = new HashSet<string>(pins.Select(pin => pin.Pin?.Trim() ?? string.Empty), StringComparer.OrdinalIgnoreCase);

            return Enumerable.Range(1, count).Count(number => !listed.Contains(number.ToString(CultureInfo.InvariantCulture)));
        }

        // The button that adds them: every pin of the case while none is listed, else the ones missing.
        public static string AddCasePins(string? package, int missing, bool noPinsYet)
        {
            string name = package?.Trim() ?? string.Empty;

            return noPinsYet && ComponentVocabulary.PinCountOf(name) is int count
                ? $"Add the {count.ToString(CultureInfo.InvariantCulture)} pins of {name}"
                : missing == 1
                    ? "Add the missing pin"
                    : $"Add the {missing.ToString(CultureInfo.InvariantCulture)} missing pins";
        }

        private static string Count(int count, string one, string many) =>
            count == 1 ? $"1 {one}" : $"{count.ToString(CultureInfo.InvariantCulture)} {many}";
    }
}
