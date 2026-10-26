using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // One pin of a component as the board's KiCad data has it: its number, the name and direction its
    // symbol gives it, and the net it sits on on THIS board ("" for none). Every value already as CRT
    // shows it.
    public sealed record KiCadPin(string Pin, string Name, string Direction, string Net);

    // A pad of a footprint in the board's PCB file - its pin's name and type as KiCad copied them from
    // the symbol (blank before KiCad 6), and the net's name ("" for none).
    public sealed record KiCadPadInfo(string Reference, string Number, string Function, string Type, string Net);

    // A pin of a symbol on the board's schematic - its number, name and electrical type.
    public sealed record KiCadSymbolPinInfo(string Reference, string Number, string Name, string Type);

    // ###########################################################################################
    // A COMPONENT'S PINS FROM THE BOARD'S KICAD DATA (CRT 4.0 step 2 - the plan: "On a KiCad board,
    // Fill from KiCad adds each pin's net and its name from the symbol. The popup shows the signal
    // name and 'net on this board' side by side").
    //
    // Plain strings in, like KiCadReferenceMatcher: the KiCad types live in CRT.App, which hands over
    // the footprints' pads and the symbols' pins.
    //
    //   - THE PADS ARE THE PIN NUMBERS AND THE NETS: a footprint whose reference IS the label (trimmed,
    //     capitals ignored - the render-time rule KiCadReferenceMatcher keeps). Since KiCad 6 each pad
    //     also carries its pin's name and type ("pinfunction", "pintype"), which reverse-engineered
    //     boards with no schematic have nothing else for. Some projects write every pad's name with its
    //     own number appended ("A2_6" on pad 6); when EVERY named pad of the footprint does, it is
    //     taken off - a real name may hold an underscore ("AUDIO_IN").
    //   - THE SYMBOL'S PINS, where the board has a schematic, name a pin before its pad does.
    //   - NAMES are CRT's spelling: KiCad's overbar "~{RESET}" is "~RESET" (the pane's own hint), and
    //     "~" - KiCad's "no name" - is none.
    //   - DIRECTIONS are ComponentVocabulary's words; a passive or unspecified pin gets none, and a
    //     supply pin is told by its name (Direction).
    //   - THE NET is never stored: it is this board's, read from its KiCad data when shown. KiCad's
    //     "no net" and its "unconnected-(...)" names are none.
    // ###########################################################################################
    public static partial class KiCadComponentPins
    {
        // ###########################################################################################
        // The pins of `label`: every pad number of its footprints and every pin of its symbols, once
        // each, in pin order (ComponentPaneEdits.PinOrder).
        // ###########################################################################################
        public static List<KiCadPin> Of(IEnumerable<KiCadPadInfo> pads, IEnumerable<KiCadSymbolPinInfo> symbolPins, string? label)
        {
            ArgumentNullException.ThrowIfNull(pads);
            ArgumentNullException.ThrowIfNull(symbolPins);

            string wanted = label?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
                return [];

            List<KiCadPadInfo> own = pads
                .Where(pad => KiCadComponentPins.IsOf(pad.Reference, wanted) && (pad.Number?.Trim() ?? string.Empty).Length > 0)
                .ToList();

            bool numbersAppended = KiCadComponentPins.NumbersAppended(own);

            var byPad = new Dictionary<string, KiCadPadInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (KiCadPadInfo pad in own)
                byPad.TryAdd(pad.Number.Trim(), pad);

            var bySymbol = new Dictionary<string, KiCadSymbolPinInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (KiCadSymbolPinInfo pin in symbolPins.Where(pin => KiCadComponentPins.IsOf(pin.Reference, wanted)))
            {
                string number = pin.Number?.Trim() ?? string.Empty;

                if (number.Length > 0)
                    bySymbol.TryAdd(number, pin);
            }

            return byPad.Keys
                .Concat(bySymbol.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(number => number, ComponentPaneEdits.PinOrder)
                .Select(number =>
                {
                    bySymbol.TryGetValue(number, out KiCadSymbolPinInfo? symbol);
                    byPad.TryGetValue(number, out KiCadPadInfo? pad);

                    string padName = pad is null ? string.Empty : KiCadComponentPins.WithoutNumber(pad.Function, number, numbersAppended);
                    string name = KiCadComponentPins.PinName(symbol?.Name) is { Length: > 0 } symbolName ? symbolName : KiCadComponentPins.PinName(padName);

                    // The symbol's type first - unless it says nothing (passive) where the pad does.
                    string direction = KiCadComponentPins.Direction(symbol?.Type, name) is { Length: > 0 } symbolDirection
                        ? symbolDirection
                        : KiCadComponentPins.Direction(pad?.Type, name);

                    return new KiCadPin(number, name, direction, KiCadComponentPins.NetName(pad?.Net));
                })
                .ToList();
        }

        // A pin's name as CRT spells it: "~{RESET}" is "~RESET", KiCad's "~" (no name) is none - the one
        // rule, KiCadNetNames.Readable.
        public static string PinName(string? kiCadName) => KiCadNetNames.Readable(kiCadName);

        // ###########################################################################################
        // A KiCad pin type as ComponentVocabulary's direction - "" for a passive, free or unspecified
        // pin. A pad's "+no_connect" (the pin is left unconnected on this board) is the pin's own type
        // still: the direction is the chip's.
        //
        // *** A SUPPLY PIN IS TOLD BY ITS NAME TOO. *** A library drawn for a reverse-engineered board
        // may type every pin "input" (the live 250469's does): a pin named GND or VSS is "Ground", and
        // one named VCC, VDD or +5V "Power", whatever its type - but a pin KiCad says is not connected.
        // ###########################################################################################
        public static string Direction(string? kiCadType, string? name)
        {
            string type = (kiCadType?.Trim() ?? string.Empty).ToLowerInvariant();
            int plus = type.IndexOf('+', StringComparison.Ordinal);

            if (plus > 0)
                type = type[..plus];

            if (type != "no_connect" && KiCadComponentPins.IsGroundName(name))
                return "Ground";

            if (type != "no_connect" && KiCadComponentPins.IsSupplyName(name))
                return "Power";

            return type switch
            {
                "input" => "Input",
                "output" or "tri_state" or "open_collector" or "open_emitter" => "Output",
                "bidirectional" => "Input/output",
                "power_in" or "power_out" => "Power",
                "no_connect" => "Not connected",
                _ => string.Empty,
            };
        }

        // ###########################################################################################
        // The net as CRT shows it - its last part ("/CPU/PHI2" is "PHI2"), spelt as a pin's name is
        // (KiCad writes a "/" inside a name as "{slash}"); none for KiCad's "no net" and for a pin
        // KiCad names "unconnected-(...)".
        // ###########################################################################################
        public static string NetName(string? net)
        {
            string name = net?.Trim() ?? string.Empty;
            int slash = name.LastIndexOf('/');

            if (slash >= 0 && slash < name.Length - 1)
                name = name[(slash + 1)..];

            return name.StartsWith("unconnected-", StringComparison.OrdinalIgnoreCase)
                ? string.Empty
                : KiCadNetNames.Readable(name);
        }

        // ###########################################################################################
        // "Fill from KiCad": `current` - the component's pins in its region - with every pin of
        // `kiCad` it lacks added (in `region`), and the name and direction of each it has filled in
        // where they are EMPTY - what was typed is never changed. In pin order. `added` and `filled`
        // count what changed.
        // ###########################################################################################
        public static List<ComponentPinEntry> Fill(
            IReadOnlyList<ComponentPinEntry> current,
            IReadOnlyList<KiCadPin> kiCad,
            string? region,
            out int added,
            out int filled)
        {
            ArgumentNullException.ThrowIfNull(current);
            ArgumentNullException.ThrowIfNull(kiCad);

            added = 0;
            filled = 0;

            var rows = current.ToList();

            foreach (KiCadPin pin in kiCad)
            {
                int index = rows.FindIndex(row => string.Equals(row.Pin?.Trim(), pin.Pin, StringComparison.OrdinalIgnoreCase));

                if (index < 0)
                {
                    rows.Add(new ComponentPinEntry { Region = region?.Trim() ?? string.Empty, Pin = pin.Pin, Name = pin.Name, Direction = pin.Direction });
                    added++;
                    continue;
                }

                ComponentPinEntry row = rows[index];
                bool nameEmpty = string.IsNullOrWhiteSpace(row.Name) && pin.Name.Length > 0;
                bool directionEmpty = string.IsNullOrWhiteSpace(row.Direction) && pin.Direction.Length > 0;

                if (!nameEmpty && !directionEmpty)
                    continue;

                rows[index] = new ComponentPinEntry
                {
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    Pin = row.Pin,
                    Name = nameEmpty ? pin.Name : row.Name,
                    Direction = directionEmpty ? pin.Direction : row.Direction,
                };

                filled++;
            }

            return rows.OrderBy(row => row.Pin, ComponentPaneEdits.PinOrder).ToList();
        }

        // Each pin's net by its number - for the popup's "net on this board".
        public static Dictionary<string, string> NetsByPin(IEnumerable<KiCadPin> pins)
        {
            ArgumentNullException.ThrowIfNull(pins);

            var nets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (KiCadPin pin in pins.Where(pin => pin.Net.Length > 0))
                nets.TryAdd(pin.Pin, pin.Net);

            return nets;
        }

        // ---- the rules ----------------------------------------------------------------------------

        private static bool IsOf(string? reference, string label) =>
            string.Equals(reference?.Trim(), label, StringComparison.OrdinalIgnoreCase);

        // Whether every named pad's name ends in "_" and its own number - then none of them means it.
        private static bool NumbersAppended(IReadOnlyList<KiCadPadInfo> pads)
        {
            List<KiCadPadInfo> named = pads.Where(pad => (pad.Function?.Trim() ?? string.Empty).Length > 0).ToList();

            return named.Count > 0 && named.All(pad =>
            {
                string function = pad.Function.Trim();
                string suffix = "_" + pad.Number.Trim();

                return function.Length > suffix.Length && function.EndsWith(suffix, StringComparison.OrdinalIgnoreCase);
            });
        }

        private static string WithoutNumber(string? function, string number, bool numbersAppended)
        {
            string name = function?.Trim() ?? string.Empty;
            string suffix = "_" + number;

            return numbersAppended && name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                ? name[..^suffix.Length]
                : name;
        }

        // GND, AGND, GND1, VSS, 0V - a supply pin that is ground.
        private static bool IsGroundName(string? name)
        {
            string upper = (name?.Trim() ?? string.Empty).TrimStart('~').ToUpperInvariant();

            return upper.Contains("GND", StringComparison.Ordinal) || upper.StartsWith("VSS", StringComparison.Ordinal) || upper == "0V";
        }

        // VCC, VDD, VBB, VEE, +5V, 12V - a supply pin that is no ground.
        private static bool IsSupplyName(string? name)
        {
            string upper = (name?.Trim() ?? string.Empty).ToUpperInvariant();

            return upper is "VCC" or "VDD" or "VBB" or "VEE" || KiCadComponentPins.Voltage().IsMatch(upper);
        }

        [GeneratedRegex(@"^\+?[1-9][0-9]*(\.[0-9]+)?V$")]
        private static partial Regex Voltage();
    }
}
