using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A PIN'S DIRECTION COMES FROM (owner decision, 2026-11-17: "OK to what you suggest for 1-4", after being shown that
    // the boards' KiCad projects type the same chip three different ways - a symbol's pin type is typed by whoever drew the
    // symbol, for KiCad's wiring check, and says nothing of the board). A direction belongs to the chip, so, in this order:
    //
    //   1. A DIRECTION A CONTRIBUTOR TYPED IS NEVER CHANGED: every writer fills only a pin that says none - and the
    //      conversion, whose pins nobody typed, takes its own "Unknown" for none.
    //   2. THE DATASHEET, for a chip CRT holds it of - a logic chip (LogicChips) or another IC (IcPinouts) - judged by the
    //      chip's number and nothing the board says against it (their For). Two things the board says stay: a pin "Not
    //      connected" on that board (an unused output of a 74LS138), and what it says of a pin the datasheet leaves
    //      unconnected or does not settle (a ROM's pin 1, wired to +5V).
    //   3. KiCad'S TYPE (KiCadComponentPins), unless its symbol types every signal pin "Input" - such a symbol was never typed.
    //   4. "Unknown".
    // ###########################################################################################
    // A datasheet's pin-out as a list of pins - the chip it is of ("6526", "7408") and each pin's number, name and direction -
    // the logic chip whose insides CRT draws when it is one (LogicChips.For's answer, so a caller drawing it looks it up
    // once), and the case the pin-out is of ("DIP-14", "QFP-44" - IcPinout.Package).
    public sealed record DatasheetPinout(string Chip, IReadOnlyList<KiCadPin> Pins, LogicChip? Logic = null, string? Package = null);

    public static class DatasheetPins
    {
        private const string NotConnected = "Not connected";

        // ###########################################################################################
        // The datasheet's every pin of `entry` with its `pins` (of one region) - its number, name and direction - for "Fill pins
        // from datasheet" (owner request, 2026-11-17), which fills them in as "Fill pins from KiCad" fills KiCad's. Null where
        // Of has none.
        // ###########################################################################################
        public static DatasheetPinout? PinoutOf(ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(entry);
            ArgumentNullException.ThrowIfNull(pins);

            static List<KiCadPin> Listed(int count, Func<int, string> name, Func<int, string> direction) =>
                Enumerable.Range(1, count)
                    .Select(pin => new KiCadPin(pin.ToString(System.Globalization.CultureInfo.InvariantCulture), name(pin), direction(pin), string.Empty))
                    .ToList();

            if (LogicChips.For(entry, pins) is { } logic)
            {
                return new DatasheetPinout(logic.Name, Listed(logic.PinCount, logic.NameOf, logic.DirectionOf), logic, logic.Package);
            }

            return IcPinouts.For(entry, pins) is { } chip
                ? new DatasheetPinout(chip.Name, Listed(chip.PinCount, chip.NameOf, chip.DirectionOf), Package: chip.Package)
                : null;
        }

        // ###########################################################################################
        // The datasheet's direction of each pin number of `entry` with its `pins` (of one region) - "" for a pin it does not
        // settle - or null when CRT holds no datasheet for the component, or the board says something against it.
        // ###########################################################################################
        public static Func<int, string>? Of(ComponentEntry entry, IReadOnlyList<ComponentPinEntry> pins)
        {
            ArgumentNullException.ThrowIfNull(entry);
            ArgumentNullException.ThrowIfNull(pins);

            // PinoutOf's lookup, the one place the catalogues are asked in their order (code review, 2026-11-17).
            return DatasheetPins.PinoutOf(entry, pins) is { } pinout
                ? pin => pin >= 1 && pin <= pinout.Pins.Count ? pinout.Pins[pin - 1].Direction : string.Empty
                : null;
        }

        // ###########################################################################################
        // What a pin of such a chip ends up as, `board` being what the board says of it now: the datasheet's - but a pin the
        // board says is "Not connected" stays so, and a pin the datasheet says nothing of, or calls unconnected, keeps what
        // the board says (when it says something).
        // ###########################################################################################
        public static string Settled(string? board, string datasheet)
        {
            string said = board?.Trim() ?? string.Empty;
            bool saysNothing = said.Length == 0 || string.Equals(said, "Unknown", StringComparison.OrdinalIgnoreCase);

            if (datasheet.Length == 0 ||
                string.Equals(said, DatasheetPins.NotConnected, StringComparison.OrdinalIgnoreCase) ||
                (datasheet == DatasheetPins.NotConnected && !saysNothing))
            {
                return board ?? string.Empty;
            }

            return datasheet;
        }

        // ###########################################################################################
        // WHAT "Fill pins from KiCad" AND THE KiCad IMPORT FILL IN - one rule for both (code review, 2026-11-17): KiCad's pins,
        // each direction settled by the datasheet of `entry`, the datasheet judged on `current` AS KiCad WOULD FILL IT - judged
        // on the pins held alone, a symbol naming the 6510's pin 9 "VCC" took the datasheet's "Output" on it. KiCad's pins as
        // they are with no component (several shown) or no datasheet.
        // ###########################################################################################
        public static IReadOnlyList<KiCadPin> OverKiCad(ComponentEntry? entry, IReadOnlyList<ComponentPinEntry> current, IReadOnlyList<KiCadPin> kiCad)
        {
            ArgumentNullException.ThrowIfNull(current);
            ArgumentNullException.ThrowIfNull(kiCad);

            return entry is null
                ? kiCad
                : DatasheetPins.Over(kiCad, DatasheetPins.Of(entry, KiCadComponentPins.Fill(current, kiCad, string.Empty, out _, out _)));
        }

        // KiCad's pins with each direction settled by the datasheet. The same list without one.
        public static IReadOnlyList<KiCadPin> Over(IReadOnlyList<KiCadPin> kiCad, Func<int, string>? datasheet)
        {
            ArgumentNullException.ThrowIfNull(kiCad);

            return datasheet is null
                ? kiCad
                : kiCad.Select(pin => ComponentVocabulary.PinNumberOf(pin.Pin) is int number
                    ? pin with { Direction = DatasheetPins.Settled(pin.Direction, datasheet(number)) }
                    : pin).ToList();
        }
    }
}
