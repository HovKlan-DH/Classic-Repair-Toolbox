using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // IMPORTANT SIGNALS SUGGESTED FROM THE BOARD'S KICAD NET NAMES (CRT 4.0 step 2 - the plan:
    // "Board > Important signals: pick nets from a searchable KiCad list, with suggestions and
    // click-to-light-up"; step 3's "Important signals: suggested from net names").
    //
    // A net whose name says what it is - a supply, ground, a clock, reset, an interrupt, a bus line -
    // is offered as a row, named as the live boards name theirs (C64 250407, 250469, KU-14194HB):
    // "+5VDC", "9VAC", "Ground", "Clock; PHI2", "RESET", "IRQ", "A0" ... "D7". Several nets may share
    // one name, as there ("9VAC" is 9VAC1 and 9VAC2; "AEC" is AEC and ~{AEC}).
    //
    //   - A NET IS TOLD BY ITS NAME ALONE, spelt down to a key: KiCad's overbar "~{RESET}" is RESET,
    //     a variant's quote ("RESET'") and a trailing "~" go, a Greek phi is PHI, capitals ignored - and
    //     the key must be the WHOLE name: "IN_GND" and "ATN_OUT" are not suggested.
    //   - A NET A ROW ALREADY HAS is not offered again, by the panel's own rule (KiCadNetChoices).
    //   - In the catalogue's order - supplies, ground, clocks, the control lines, the buses - each name's
    //     nets in natural order.
    //
    // PURE: names in, rows out.
    // ###########################################################################################
    public static partial class ImportantSignalSuggestions
    {
        // ###########################################################################################
        // The rows to suggest for a board whose KiCad nets are `choices` and whose important signals
        // are `existing` - each a display name and the net, as the picker spells it.
        // ###########################################################################################
        public static List<KiCadImportantSignalEntry> For(KiCadNetChoices choices, IEnumerable<KiCadImportantSignalEntry> existing)
        {
            ArgumentNullException.ThrowIfNull(choices);
            ArgumentNullException.ThrowIfNull(existing);

            var used = new HashSet<string>(
                existing.Select(row => choices.NormalizedOf(row.KiCadNetName)).OfType<string>(),
                StringComparer.OrdinalIgnoreCase);

            return choices.Names
                .Where(net => !used.Contains(net))
                .Select(net => (Net: net, Suggestion: ImportantSignalSuggestions.NameOf(net)))
                .Where(found => found.Suggestion is not null)
                .OrderBy(found => found.Suggestion!.Value.Rank)
                .ThenBy(found => found.Suggestion!.Value.DisplayName, NaturalLabelComparer.Instance)
                .ThenBy(found => found.Net, NaturalLabelComparer.Instance)
                .Select(found => new KiCadImportantSignalEntry { DisplayName = found.Suggestion!.Value.DisplayName, KiCadNetName = found.Net })
                .ToList();
        }

        // ###########################################################################################
        // The name a net is suggested under, with its place in the catalogue - null for a net whose
        // name says nothing CRT knows.
        // ###########################################################################################
        public static (int Rank, string DisplayName)? NameOf(string? net)
        {
            string key = ImportantSignalSuggestions.KeyOf(net);

            if (key.Length == 0)
                return null;

            if (key is "+9V_UNREG" or "9V_UNREG" or "+9VUNREG")
                return (1, "+9VDC (unregulated)");

            if (key == "CAN+5V")
                return (1, "+5VDC CAN");

            if (ImportantSignalSuggestions.Supply().Match(key) is { Success: true } supply)
                return (1, $"+{supply.Groups[1].Value}VDC");

            if (ImportantSignalSuggestions.Mains().IsMatch(key))
                return (2, "9VAC");

            if (key is "GND" or "AGND" or "DGND" or "GNDA" or "GNDD" or "EARTH" or "0V")
                return (3, "Ground");

            if (ImportantSignalSuggestions.Clock(key) is { } clock)
                return (4, clock);

            if (ImportantSignalSuggestions.ControlLines.TryGetValue(key, out string? control))
                return (5, control);

            if (ImportantSignalSuggestions.AddressLine().IsMatch(key))
                return (6, key);

            if (ImportantSignalSuggestions.DataLine().IsMatch(key))
                return (7, key);

            return null;
        }

        // ###########################################################################################
        // A net's name - the last part of its hierarchical name, as the picker offers it, where KiCad
        // writes a "/" as "{slash}" - spelt down to what it says: "~{RESET'}" is RESET,
        // "R{slash}~{W}" is R/W, a phi and "DOT" is PHIDOT, "'9VAC~" is 9VAC. Read through
        // KiCadNetNames.Readable, the one rule (code review, 2026-10-26), then every overbar mark dropped.
        // ###########################################################################################
        public static string KeyOf(string? net)
        {
            string name = KiCadNetNames.Readable(net).Replace("~", string.Empty, StringComparison.Ordinal);

            // Capital phi, the phi symbol and small phi - each written as PHI too.
            name = name.Replace("\u03A6", "PHI", StringComparison.Ordinal)
                .Replace("\u03D5", "PHI", StringComparison.Ordinal)
                .Replace("\u03C6", "PHI", StringComparison.Ordinal);

            return name.Trim('\'', ' ').ToUpperInvariant();
        }

        // The clocks - "Clock; ..." as the live boards name them.
        private static string? Clock(string key) => key switch
        {
            "PHI0" or "PHI_0" => "Clock; PHI0",
            "PHI1" or "PHI_1" => "Clock; PHI1",
            "PHI2" or "PHI_2" => "Clock; PHI2",
            "DOT_CLOCK" or "DOTCLK" or "DOT_CLK" or "DOTCLOCK" or "PHIDOT" or "PHI_DOT" => "Clock; DOT Clock",
            "COLOR_CLOCK" or "COLORCLK" or "COLOR_CLK" or "PHI_COLOR" or "PHICOLOR" => "Clock; PHI Color",
            "CLK" or "CLOCK" => "Clock",
            _ when ImportantSignalSuggestions.Megahertz().Match(key) is { Success: true } mhz =>
                string.Create(CultureInfo.InvariantCulture, $"Clock; {mhz.Groups[1].Value} MHz"),
            _ => null,
        };

        // The control lines every 8-bit board has some of - by the name they are known by.
        private static readonly Dictionary<string, string> ControlLines = new(StringComparer.Ordinal)
        {
            ["RESET"] = "RESET",
            ["RES"] = "RESET",
            ["RST"] = "RESET",
            ["IRQ"] = "IRQ",
            ["NMI"] = "NMI",
            ["INT"] = "INT",
            ["R/W"] = "R/W",
            ["RW"] = "R/W",
            ["R_W"] = "R/W",
            ["BA"] = "BA",
            ["AEC"] = "AEC",
            ["RDY"] = "RDY",
            ["DMA"] = "DMA",
            ["ATN"] = "ATN",
            ["CAS"] = "CAS",
            ["RAS"] = "RAS",
            ["EXROM"] = "EXROM",
            ["GAME"] = "GAME",
            ["WAIT"] = "WAIT",
            ["BUSRQ"] = "BUSRQ",
            ["BUSAK"] = "BUSAK",
            ["MREQ"] = "MREQ",
            ["IORQ"] = "IORQ",
            ["RD"] = "RD",
            ["WR"] = "WR",
            ["M1"] = "M1",
        };

        // +5V, 5V, +12V, +5VDC - a supply of so many volts (0V is ground).
        [GeneratedRegex(@"^\+?([1-9][0-9]*)V(DC)?$")]
        private static partial Regex Supply();

        // 9VAC, 9VAC1, 9V_AC2 - the C64 family's AC supply.
        [GeneratedRegex(@"^\+?9V_?AC[0-9]*$")]
        private static partial Regex Mains();

        [GeneratedRegex(@"^([0-9]+(\.[0-9]+)?)MHZ$")]
        private static partial Regex Megahertz();

        [GeneratedRegex(@"^A([0-9]|1[0-5])$")]
        private static partial Regex AddressLine();

        [GeneratedRegex(@"^D[0-7]$")]
        private static partial Regex DataLine();
    }
}
