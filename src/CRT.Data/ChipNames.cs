using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // HOW A CHIP'S NAME AND ITS SUPPLY PINS ARE READ BY THE DATASHEETS' CHECKS - one rule for LogicChips and IcPinouts (code
    // review, 2026-11-17: each held its own ground and supply names and its own 74-series name, and they had drifted apart).
    //
    // *** THE SUPPLY NAMES ARE EXACT, ON PURPOSE. *** A board names a pin strapped to a rail by its signal AND the rail - the
    // C16's and Plus 4's PLA has "CE - GND" on its /CE input - and a check reading "contains GND" took that for a ground, so the
    // whole chip lost its datasheet (measured on the shipped boards, 2026-11-17). KiCadComponentPins types a pin by a looser
    // rule of its own (GND inside a name, any voltage) - a direction to start from, no judgement against a datasheet.
    // ###########################################################################################
    public static partial class ChipNames
    {
        // ###########################################################################################
        // The first word of `text` in capitals, a micro sign or a Greek mu written as the "U" a maker's "uPD" and "uA" are
        // written with in plain letters (code review, 2026-11-17): ToUpperInvariant turns the micro sign into a CAPITAL Greek
        // mu, which no rule knew, so "µPD4164C" was no 4164. "" for none.
        // ###########################################################################################
        public static string FirstWord(string? text)
        {
            string word = (text ?? string.Empty).Trim().Split((char[]?)null, 2, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? string.Empty;

            return word.Replace('µ', 'U').Replace('μ', 'U').ToUpperInvariant();
        }

        // ###########################################################################################
        // The number of a 74-series name as written - "08" of "SN74LS08N" - or null for none: a maker's prefix (SN, DM, MC,
        // HD, M, N, MB, TC, PC, GD, uPD, F, CD), "74" or "54", the logic family (LS, HCT, ...), the number, and a suffix that
        // starts with a letter (the case, the grade).
        // ###########################################################################################
        public static string? SeventyFourNumber(string word) =>
            ChipNames.SeventyFour().Match(word ?? string.Empty) is { Success: true } match ? match.Groups["number"].Value : null;

        // The logic family of a 74-series name as written, in capitals - "LS" of "SN74LS51N", "" of "SN7451N" - or null for no
        // 74-series name: for the few numbers whose families differ in their pins (IcPinout.Families).
        public static string? SeventyFourFamily(string word) =>
            ChipNames.SeventyFour().Match(word ?? string.Empty) is { Success: true } match ? match.Groups["family"].Value.ToUpperInvariant() : null;

        // ###########################################################################################
        // *** A RESISTOR, CAPACITOR, DIODE OR CONNECTOR IS NO CHIP (code review, 2026-11-17): a name matches whatever follows
        // its number, so a Bourns "4116R-1-103" resistor network read as a 4116 DRAM, and a "5400pF" capacitor as a 7400 - and
        // the conversion gave their pins the chip's directions. A component of such a category - the words KiCadBoardComponents
        // files those references under, and any longer one starting with them ("Resistor pack") - has no datasheet here.
        // ###########################################################################################
        private static readonly string[] NoChipCategories = ["Resistor", "Capacitor", "Diode", "Connector"];

        public static bool IsNoChip(string? category)
        {
            string said = category?.Trim() ?? string.Empty;

            return ChipNames.NoChipCategories.Any(word => said.StartsWith(word, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Whether the component's package says nothing against the case the chip's pin-out is of, `chipPackage` - an empty
        // package is none: the chip says its own. It fits when it names the same case, its layout and its count
        // (ComponentVocabulary.ChipCaseOf): a SOIC-14 a 74LS08's DIP-14, numbered alike (owner request, 2026-11-18: the
        // Amiga "has a lot of SMD components"), but no PLCC-44 a chip of a QFP-44 or a DIP-40. LogicChips.For's and
        // IcPinouts.For's one rule (code review, 2026-11-17).
        // ###########################################################################################
        public static bool PackageFits(ComponentEntry entry, string chipPackage)
        {
            ArgumentNullException.ThrowIfNull(entry);

            string package = (entry.Package ?? string.Empty).Trim();

            return package.Length == 0 ||
                   (ComponentVocabulary.ChipCaseOf(package) is { } said && said == ComponentVocabulary.ChipCaseOf(chipPackage));
        }

        // GND, GROUND, VSS, 0V, EARTH - a pin named as ground, the whole name.
        public static bool IsGround(string? name) => ChipNames.GroundName().IsMatch(name?.Trim() ?? string.Empty);

        // VCC, VDD, VBB, VPP, +5V, +5VDC, +VCC, V+, +12V, +9V - a pin named as a supply, the whole name.
        public static bool IsSupply(string? name) => ChipNames.SupplyName().IsMatch(name?.Trim() ?? string.Empty);

        [GeneratedRegex(@"^(?:SN|DM|MC|HD|MB|TC|PC|GD|UPD|M|N|F|CD)?(?:74|54)(?<family>LS|ALS|AS|HCT|HCU|HC|ACT|AC|AHCT|AHC|ABT|LVC|LV|FCT|S|F|L|H|C)?(?<number>\d{2,4})(?:[A-Z][A-Z0-9-]*)?$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
        private static partial Regex SeventyFour();

        [GeneratedRegex(@"^(gnd|ground|vss|0v|earth)$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
        private static partial Regex GroundName();

        [GeneratedRegex(@"^(vcc|vdd|vbb|vpp|\+?5v|\+5vdc|\+vcc|v\+|\+?12v|\+?9v)$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase)]
        private static partial Regex SupplyName();
    }
}
