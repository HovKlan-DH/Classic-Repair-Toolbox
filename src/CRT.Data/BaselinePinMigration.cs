using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** TEMPORARY - REMOVED BEFORE CRT 4.0 GOES LIVE (owner request, 2026-11-01: "I need a possibility
    // to temporarily create/migrate 'scope baseline data' into 'pinout data' ... This convert/migrate
    // functionality will be removed just before the go-live"). ***
    //
    // CRT 4.0's pinouts are new, so no published board has one - while their scope baselines already
    // say which pin carries which signal. This makes a component's pins of its scope baselines, so the
    // owner need not type them: 677 components across the 14 boards measured on 2026-11-01, not one
    // saying two names for a pin. A pin number and a name; the direction told from the name where it
    // can be, else "Unknown".
    //
    //   - THE SAME PINS IN EVERY REGION ARE ONE SET, for every region (323 of the 329 components with
    //     baselines of both regions; the other six only look different - below). A set that differs
    //     stays per region.
    //   - A BASELINE WITH NO REGION BESIDE REGIONAL ONES SAYS NOTHING ABOUT THE PIN: those six are the
    //     VIC-II's "SYNC/LUM explanation", a second picture of pin 15, not its name.
    //   - WHAT WAS TYPED STAYS: only a pin missing is added, and only an empty name or direction filled
    //     in, as "Fill from KiCad" does.
    //   - THE PACKAGE ONLY WHEN IT IS NO GUESS: an IC with an even number of pins and a baseline on
    //     every one of them is a DIP of that many, set where the component has none. Anything else
    //     gets its pins and leaves the package to be picked - an RF modulator's eight are no DIP-8.
    //
    // PURE: takes a board, returns a board. BaselinePinMigrationTests fails a release version while this
    // class still exists, which stops the release workflow - see .claude/CLAUDE.md, "Not done yet".
    // ###########################################################################################
    public static class BaselinePinMigration
    {
        private const string IcCategory = "IC";

        // Whether the component has a scope baseline to make pins of, and that would change anything.
        public static bool CanRun(BoardData board, string? label) =>
            !ReferenceEquals(BaselinePinMigration.Apply(board, label), board);

        // ###########################################################################################
        // The board with the component's pins made of its scope baselines - the same board when it has
        // none, or nothing would change.
        // ###########################################################################################
        public static BoardData Apply(BoardData board, string? label)
        {
            ArgumentNullException.ThrowIfNull(board);

            string wanted = label?.Trim() ?? string.Empty;

            List<ComponentImageEntry> baselines = board.ComponentImages
                .Where(row => BaselinePinMigration.SameLabel(row.BoardLabel, wanted) && ComponentImageKinds.IsScopeBaseline(row))
                .ToList();

            if (wanted.Length == 0 || baselines.Count == 0 || !board.Components.Any(row => BaselinePinMigration.SameLabel(row.BoardLabel, wanted)))
            {
                return board;
            }

            Dictionary<string, Dictionary<string, string>> sets = BaselinePinMigration.SetsByRegion(baselines);

            List<ComponentPinEntry> own = board.ComponentPins.Where(row => BaselinePinMigration.SameLabel(row.BoardLabel, wanted)).ToList();
            List<ComponentPinEntry> pins = own.ToList();
            bool changed = false;

            foreach ((string region, Dictionary<string, string> set) in sets)
            {
                foreach ((string pin, string name) in set)
                {
                    changed |= BaselinePinMigration.AddOrFill(pins, wanted, region, pin, name);
                }
            }

            BoardData result = changed
                ? board.WithComponentSections(componentPins: BaselinePinMigration.WithOwnPins(board.ComponentPins, wanted, pins))
                : board;

            return BaselinePinMigration.WithPackage(result, wanted, sets);
        }

        // ###########################################################################################
        // Every component of the board with a baseline, as one change - `components` says how many
        // changed.
        // ###########################################################################################
        public static BoardData ApplyToBoard(BoardData board, out int components)
        {
            ArgumentNullException.ThrowIfNull(board);

            components = 0;
            BoardData result = board;

            foreach (string label in board.ComponentImages
                .Where(ComponentImageKinds.IsScopeBaseline)
                .Select(row => row.BoardLabel?.Trim() ?? string.Empty)
                .Where(label => label.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList())
            {
                BoardData next = BaselinePinMigration.Apply(result, label);

                if (!ReferenceEquals(next, result))
                {
                    components++;
                    result = next;
                }
            }

            return result;
        }

        // Each region's pins and names - one set under "" when every region says the same.
        private static Dictionary<string, Dictionary<string, string>> SetsByRegion(IReadOnlyList<ComponentImageEntry> baselines)
        {
            List<ComponentImageEntry> regional = baselines.Where(row => !string.IsNullOrWhiteSpace(row.Region)).ToList();
            IReadOnlyList<ComponentImageEntry> telling = regional.Count > 0 ? regional : baselines;

            Dictionary<string, Dictionary<string, string>> sets = telling
                .GroupBy(row => row.Region?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(
                    group => group.Key,
                    group => BaselinePinMigration.PinsOf(group),
                    StringComparer.OrdinalIgnoreCase);

            List<Dictionary<string, string>> all = sets.Values.ToList();

            return all.Skip(1).All(set => BaselinePinMigration.SameSet(all[0], set))
                ? new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase) { [string.Empty] = all[0] }
                : sets;
        }

        // Pin -> name: the pin's number on the case ("2 [64K]" is pin 2), an exact "2" before such a
        // row, otherwise the first row of a pin.
        private static Dictionary<string, string> PinsOf(IEnumerable<ComponentImageEntry> rows)
        {
            var pins = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (ComponentImageEntry row in rows.OrderBy(row => ComponentVocabulary.PinNumberOf(row.Pin)?.ToString(System.Globalization.CultureInfo.InvariantCulture) == row.Pin?.Trim() ? 0 : 1))
            {
                string pin = ComponentVocabulary.PinNumberOf(row.Pin)?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? row.Pin?.Trim() ?? string.Empty;

                if (pin.Length > 0)
                {
                    pins.TryAdd(pin, row.Name?.Trim() ?? string.Empty);
                }
            }

            return pins;
        }

        private static bool SameSet(Dictionary<string, string> first, Dictionary<string, string> second) =>
            first.Count == second.Count && first.All(pair =>
                second.TryGetValue(pair.Key, out string? name) && string.Equals(pair.Value, name, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // A pin added, or an empty name or direction filled in - true when anything changed. A pin the
        // component has for this region or for every region is the same pin; new pins for every region
        // match it in any region.
        // ###########################################################################################
        private static bool AddOrFill(List<ComponentPinEntry> pins, string label, string region, string pin, string name)
        {
            int index = pins.FindIndex(row =>
                string.Equals(row.Pin?.Trim(), pin, StringComparison.OrdinalIgnoreCase) &&
                (region.Length == 0 || string.IsNullOrWhiteSpace(row.Region) || string.Equals(row.Region?.Trim(), region, StringComparison.OrdinalIgnoreCase)));

            string direction = BaselinePinMigration.DirectionOf(name);

            if (index < 0)
            {
                pins.Add(new ComponentPinEntry { BoardLabel = label, Region = region, Pin = pin, Name = name, Direction = direction });
                return true;
            }

            ComponentPinEntry row = pins[index];
            bool fillName = string.IsNullOrWhiteSpace(row.Name) && name.Length > 0;
            bool fillDirection = string.IsNullOrWhiteSpace(row.Direction);

            if (!fillName && !fillDirection)
            {
                return false;
            }

            pins[index] = new ComponentPinEntry
            {
                BoardLabel = row.BoardLabel,
                Region = row.Region,
                Pin = row.Pin,
                Name = fillName ? name : row.Name,
                Direction = fillDirection ? direction : row.Direction,
            };

            return true;
        }

        // What the pin's name says - the rule "Fill from KiCad" tells a supply pin by, and NC.
        private static string DirectionOf(string name)
        {
            string fromName = KiCadComponentPins.Direction(null, name);

            if (fromName.Length > 0)
            {
                return fromName;
            }

            return string.Equals(name, "NC", StringComparison.OrdinalIgnoreCase) || string.Equals(name, "N/C", StringComparison.OrdinalIgnoreCase)
                ? "Not connected"
                : "Unknown";
        }

        // The other components' pins left where they stand, this one's in pin order where its first was.
        private static List<ComponentPinEntry> WithOwnPins(List<ComponentPinEntry> all, string label, List<ComponentPinEntry> own)
        {
            List<ComponentPinEntry> ordered = own
                .OrderBy(row => row.Region?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(row => row.Pin, ComponentPaneEdits.PinOrder)
                .ToList();

            int first = all.FindIndex(row => BaselinePinMigration.SameLabel(row.BoardLabel, label));
            var rows = all.Where(row => !BaselinePinMigration.SameLabel(row.BoardLabel, label)).ToList();

            rows.InsertRange(first < 0 ? rows.Count : first, ordered);

            return rows;
        }

        // A DIP of the pins' count on an IC with a baseline on every pin and no package yet.
        private static BoardData WithPackage(BoardData board, string label, Dictionary<string, Dictionary<string, string>> sets)
        {
            BoardData result = board;

            foreach (ComponentEntry row in board.Components.Where(row => BaselinePinMigration.SameLabel(row.BoardLabel, label)).ToList())
            {
                string rowRegion = row.Region?.Trim() ?? string.Empty;

                Dictionary<string, string>? set = sets.TryGetValue(string.Empty, out Dictionary<string, string>? common)
                    ? common
                    : sets.TryGetValue(rowRegion, out Dictionary<string, string>? own) ? own : null;

                if (!string.IsNullOrWhiteSpace(row.Package) ||
                    !string.Equals(row.Category?.Trim(), BaselinePinMigration.IcCategory, StringComparison.OrdinalIgnoreCase) ||
                    set is null ||
                    BaselinePinMigration.DipOf(set) is not { } package)
                {
                    continue;
                }

                result = ComponentPaneEdits.WithField(result, label, rowRegion, ComponentField.Package, package);
            }

            return result;
        }

        // "DIP-40" when the pins are exactly 1 to an even count of at least four.
        private static string? DipOf(Dictionary<string, string> set)
        {
            List<int> numbers = set.Keys
                .Select(pin => int.TryParse(pin, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int number) ? number : 0)
                .OrderBy(number => number)
                .ToList();

            bool complete = numbers.Count >= 4 && numbers.Count % 2 == 0 && numbers.Select((number, index) => number == index + 1).All(same => same);

            return complete ? $"DIP-{numbers.Count.ToString(System.Globalization.CultureInfo.InvariantCulture)}" : null;
        }

        private static bool SameLabel(string? rowLabel, string wanted) =>
            string.Equals(rowLabel?.Trim(), wanted, StringComparison.OrdinalIgnoreCase);
    }
}
