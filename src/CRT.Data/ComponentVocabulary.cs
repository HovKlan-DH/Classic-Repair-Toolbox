using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // How a chip's case lays out its pins (ComponentVocabulary.ChipCaseOf): two rows, as a DIP; four sides with pin 1 in
    // the middle of the top one, as a PLCC; four sides with pin 1 at the top of the left one, as a QFP.
    public enum ChipCaseKind
    {
        TwoRows,
        Plcc,
        Qfp,
    }

    // A chip's case as its name says it: how its pins are laid out, and how many. Two names of one case are equal ("DIP-14"
    // and "SOIC-14", "QFP-44" and "LQFP-44").
    public sealed record ChipCase(ChipCaseKind Kind, int Pins)
    {
        // How many pins stand on each side: half of them in two rows, a quarter on four sides.
        public int PerSide => this.Kind == ChipCaseKind.TwoRows ? this.Pins / 2 : this.Pins / 4;
    }

    // A crystal oscillator's metal can on a DIP's pads (ComponentVocabulary.OscillatorOf): how many places the DIP it stands
    // on has, and the four of them it has a pin on, in order.
    public sealed record OscillatorCan(int Positions, IReadOnlyList<int> Pins);

    // ###########################################################################################
    // THE WORDS CRT 4.0's NEW COMPONENT DATA IS WRITTEN IN (step 2, 2026-10-25) - a pin's direction,
    // what kind of stand-in a replacement is, and the usual chip cases.
    //
    // Stored as these words, not as codes: a board file is read in a table too, where "Datasheet"
    // says what it is and "2" does not. So the words here ARE the data - rewording one changes
    // what every board already published holds. Each list is what the component pane offers and what
    // the table's checks know (BoardDataChecks: anything else is a WARNING, never a refusal - an
    // empty value is fine everywhere, as unfinished work is). Compared ignoring capitals, as every
    // other value of a board is.
    // ###########################################################################################
    public static class ComponentVocabulary
    {
        // How long a short description may be (owner decision, 2026-11-03; 100 since 2026-10-24). Measured
        // that day: of the 539 published ones the longest is 118, 525 are at most 50 and none lies between
        // 51 and 60 - the 14 over it are repair notes on three boards. Longer is a warning, never a refusal
        // (BoardDataChecks).
        public const int ShortDescriptionLimit = 50;

        // ###########################################################################################
        // A pin's job, as a hobbyist reads a pinout - KiCad's dozen electrical types fold into these.
        //
        // "Power" is the pin that feeds the chip (VCC, +5V); a pin carrying a VOLTAGE LEVEL as its
        // signal - a sound chip's audio out, a video chip's luminance - is simply "Input" or "Output"
        // (owner request, 2026-11-02: "Analog input" and "Analog output" went again after a day, as "it
        // can be covered by the existing two choices"). "Unknown" is last: nobody has looked yet, which
        // "Other" does not say. The order is the owner's (2026-11-09): "Not connected" first, then the signals,
        // ground before power.
        // ###########################################################################################
        public static IReadOnlyList<string> PinDirections { get; } =
        [
            "Not connected",
            "Input",
            "Output",
            "Input/output",
            "Ground",
            "Power",
            "Other",
            "Unknown",
        ];

        // A replacement part fits as it is, only after something is changed (said in its Note), or is
        // better than the original (owner requests, 2026-11-02 and 2026-11-03: "Drop-in" and "Needs
        // changes" until then). No published board held a kind yet, nor any draft that day.
        public static IReadOnlyList<string> ReplacementKinds { get; } =
        [
            "Drop-in replacement",
            "Needs modification",
            "Improvement",
        ];

        // ###########################################################################################
        // The cases a component comes in (the plan: "DIP-8 to DIP-40, wide DIP, SIP, TO-92, TO-220, or
        // 'other'"; the rest read off the KiCad footprints of the six published boards with a PCB file,
        // owner request, 2026-11-03: "what package is a resistor? Or a ferrite bead or a capacitor?").
        // Offered, not enforced - a board may name any case, and one CRT cannot draw is shown as its pin
        // list alone. A DIP's pin count is in its name; the leaded parts' say nothing of any count.
        // ###########################################################################################
        public static IReadOnlyList<string> Packages { get; } =
        [
            "DIP-8",
            "DIP-14",
            "DIP-16",
            "DIP-18",
            "DIP-20",
            "DIP-22",
            "DIP-24",
            "DIP-24 wide",
            "DIP-28",
            "DIP-28 wide",
            "DIP-40",
            "DIP-48",
            "DIP-64",
            "SOIC-8",
            "SOIC-14",
            "SOIC-16",
            "SOIC-20",
            "SOIC-24",
            "SOIC-28",
            "PLCC-20",
            "PLCC-28",
            "PLCC-44",
            "PLCC-52",
            "PLCC-68",
            "PLCC-84",
            "QFP-32",
            "QFP-44",
            "QFP-48",
            "QFP-64",
            "QFP-80",
            "QFP-100",
            "QFP-144",
            "QFP-208",
            "SIP",
            "TO-92",
            "TO-220",
            "Axial",
            "Radial",
            "Disc",
            "DO-35",
            "DO-41",
            "EMI filter",
            "HC-49",
            "HC-18",
            "Oscillator DIP-14",
            "Oscillator DIP-8",
            "Pin header",
            "Other",
        ];

        // ###########################################################################################
        // A scope baseline's T/DIV and V/DIV as the pane offers them (the plan: "T/DIV and V/DIV as
        // drop-downs") - the 1-2-5 steps every scope uses, spelt as the converted data spells them
        // ("50us", "1ms", "500ns", "1V", "500mV": all 13,862 published baselines, 2026-10-26). Offered,
        // not enforced: a box takes any value, and one already published reads as it is.
        // ###########################################################################################
        public static IReadOnlyList<string> TimeDivisions { get; } = ComponentVocabulary.Steps(["ns", "us", "ms", "s"], lastUnitUpTo: 10);

        public static IReadOnlyList<string> VoltDivisions { get; } = ComponentVocabulary.Steps(["mV", "V"], lastUnitUpTo: 10);

        // 1, 2, 5, 10, 20, 50, 100, 200, 500 of each unit, then the last unit up to `lastUnitUpTo`.
        private static List<string> Steps(IReadOnlyList<string> units, int lastUnitUpTo)
        {
            int[] values = [1, 2, 5, 10, 20, 50, 100, 200, 500];
            var steps = new List<string>();

            for (int unit = 0; unit < units.Count; unit++)
            {
                bool last = unit == units.Count - 1;

                steps.AddRange(values
                    .Where(value => !last || value <= lastUnitUpTo)
                    .Select(value => $"{value.ToString(System.Globalization.CultureInfo.InvariantCulture)}{units[unit]}"));
            }

            return steps;
        }

        // ###########################################################################################
        // How many pins a case has, by its name - what the pane needs to number the pins and draw the
        // case: a chip's case says it (ChipCaseOf - "DIP-40", "DIP-24 wide", "dip 28", "SOIC-14",
        // "PLCC-84", "QFP-44"), TO-92 and TO-220 have three. Null for a case whose count is not in its
        // name - SIP, "Other", a SOT-23, anything typed.
        // ###########################################################################################
        public static int? PinCountOf(string? package)
        {
            string text = (package ?? string.Empty).Trim();

            if (string.Equals(text, "TO-92", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(text, "TO-220", StringComparison.OrdinalIgnoreCase))
            {
                return 3;
            }

            return ComponentVocabulary.ChipCaseOf(text)?.Pins;
        }

        // ###########################################################################################
        // THE CASES A CHIP IS DRAWN IN, BY THEIR NAMES (owner request, 2026-11-18: "Drawing square packages ... yes, I want
        // ... because there might come new hardware soon - e.g. Amiga 500 or Amiga 1000", and "It has a lot of SMD
        // components"):
        //
        //   - TWO ROWS: a DIP and every surface-mounted case NUMBERED AS A DIP - pin 1 at the top of the left row, down it and
        //     up the other (SOIC, SO, SOP, SSOP, TSSOP, TSOP, SOJ) - an even count up to 128, anything after it ("DIP-24
        //     wide", "SOIC-16 wide"). A 74HC08 in a SOIC-14 has its DIP's pin-out.
        //   - FOUR SIDES, SQUARE: a PLCC - pin 1 in the middle of the top side, an odd number of pins a side (20, 28, 44, 52,
        //     68, 84 - the rectangular PLCC-32 has eight and nine, and is none), anything after its count ("PLCC-84 socket")
        //     - and a QFP (QFP, LQFP, TQFP, PQFP) - pin 1 at the top of the left side - a quarter of its pins a side. A QFP only
        //     as the plain name: "QFP-100 14x20mm" says a rectangular case, whose sides CRT cannot tell, and is none.
        //
        // Pins counted counter-clockwise from pin 1, seen from above, as every maker numbers them. A SOT-23 and the like
        // are none: three pins of no row.
        // ###########################################################################################
        public static ChipCase? ChipCaseOf(string? package)
        {
            string text = (package ?? string.Empty).Trim();
            string family = new string(text.TakeWhile(char.IsAsciiLetter).ToArray()).ToUpperInvariant();
            string rest = text[family.Length..].TrimStart('-', ' ');
            string digits = new(rest.TakeWhile(char.IsAsciiDigit).ToArray());
            string after = rest[digits.Length..].Trim();

            if (!int.TryParse(digits, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count) || count <= 0)
            {
                return null;
            }

            if (ComponentVocabulary.TwoRowFamilies.Contains(family))
            {
                return count <= 128 && count % 2 == 0 ? new ChipCase(ChipCaseKind.TwoRows, count) : null;
            }

            if (count % 4 != 0)
            {
                return null;
            }

            // A PLCC's shape is settled by its count, so words after it say nothing against it - "PLCC-84 socket", the A500's
            // Fat Agnus (code review, 2026-11-18: refused, nothing of its datasheet was drawn or checked).
            if (family == "PLCC")
            {
                return count is >= 20 and <= 124 && (count / 4) % 2 == 1 ? new ChipCase(ChipCaseKind.Plcc, count) : null;
            }

            return after.Length == 0 && ComponentVocabulary.QfpFamilies.Contains(family) && count is >= 8 and <= 256 ? new ChipCase(ChipCaseKind.Qfp, count) : null;
        }

        private static readonly string[] TwoRowFamilies = ["DIP", "SOIC", "SO", "SOP", "SSOP", "TSSOP", "TSOP", "SOJ"];

        private static readonly string[] QfpFamilies = ["QFP", "LQFP", "TQFP", "PQFP"];

        // ###########################################################################################
        // A CRYSTAL OSCILLATOR IN A METAL CAN, by its case's name (owner request, 2026-11-18: "investigate what kind of
        // components an Amiga 500 and Amiga 1000 have, and then support those" - the A500's X1 and the A1000's OSC1): a
        // full-size can stands on a DIP-14's pads with a pin at each corner - 1, 7, 8 and 14 - and a half-size one on a
        // DIP-8's - 1, 4, 5 and 8 - as every maker places them. "Oscillator DIP-14", "Oscillator_DIP-8" as KiCad names its
        // footprint. No chip's case (ChipCaseOf): its other places have no pin, so none is counted (PinCountOf) and the pane
        // offers no "Add the N pins". Null for anything else.
        // ###########################################################################################
        public static OscillatorCan? OscillatorOf(string? package)
        {
            System.Text.RegularExpressions.Match match = ComponentVocabulary.OscillatorName.Match((package ?? string.Empty).Trim());

            return !match.Success ? null
                : match.Groups["count"].Value == "14" ? new OscillatorCan(14, [1, 7, 8, 14])
                : new OscillatorCan(8, [1, 4, 5, 8]);
        }

        private static readonly System.Text.RegularExpressions.Regex OscillatorName = new(
            @"^oscillator[-_ ]?dip[-_ ]?(?<count>14|8)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

        // ###########################################################################################
        // The number on the case a pin's text starts with - "2 [64K]" is pin 2, as four published
        // baselines write it - or null for one that starts with none ("A1", blank). Digits running straight
        // on into a letter are a NAME (code review, 2026-11-01): a regulator's "12V" is no pin 12, and Submit
        // refused it for it; "1A" and "1B" on a 7400 are two pins, and no pin 1.
        // ###########################################################################################
        public static int? PinNumberOf(string? pin)
        {
            string text = (pin ?? string.Empty).Trim();
            string digits = new(text.TakeWhile(char.IsAsciiDigit).ToArray());

            return digits.Length is > 0 and < 6 && (digits.Length == text.Length || !char.IsLetterOrDigit(text[digits.Length]))
                ? int.Parse(digits, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture)
                : null;
        }

        // ###########################################################################################
        // Whether two pins' texts name one pin: the same words, capitals aside, or the same number -
        // "2 [64K]" being pin 2 (PinNumberOf). A blank is no pin.
        // ###########################################################################################
        public static bool IsSamePin(string? first, string? second)
        {
            string one = (first ?? string.Empty).Trim();
            string other = (second ?? string.Empty).Trim();

            if (one.Length == 0 || other.Length == 0)
            {
                return false;
            }

            return string.Equals(one, other, StringComparison.OrdinalIgnoreCase) ||
                   (ComponentVocabulary.PinNumberOf(one) is int number && ComponentVocabulary.PinNumberOf(other) == number);
        }

        // ###########################################################################################
        // Whether a pin can be on the case - false only when both are known and the number is past the
        // case's last pin (owner request, 2026-11-01: "if I e.g. put pin 48 on a 24 pin package, this is
        // an error"), or none of an oscillator's four (OscillatorOf). A case whose count CRT cannot tell, and a
        // pin named rather than numbered, say nothing.
        // ###########################################################################################
        public static bool IsOnPackage(string? package, string? pin)
        {
            if (ComponentVocabulary.PinNumberOf(pin) is not int number)
            {
                return true;
            }

            if (ComponentVocabulary.OscillatorOf(package) is { } can)
            {
                return can.Pins.Contains(number);
            }

            return ComponentVocabulary.PinCountOf(package) is not int count || (number >= 1 && number <= count);
        }

        // ###########################################################################################
        // WHAT THE PANE'S "Package" AND "Direction" OFFER ON A BOARD (owner request, 2026-11-11: "if I manually create/change a
        // package in the table view - e.g. from "Disc" to "Diskette", then in UI view I should be able to use this new
        // "Diskette" value also in the drop-down. Use the same behaviour for the "Direction" drop-down"): the usual words, then
        // every other word the board's components already use, alphabetically - each once, capitals aside, in the board's first
        // spelling. "Unknown" stays the last direction (owner's order, 2026-11-09), so the board's own come before it.
        // ###########################################################################################
        public static IReadOnlyList<string> PackagesOf(BoardData? board) =>
            ComponentVocabulary.WithWordsOf(ComponentVocabulary.Packages, board?.Components.Select(row => row.Package), lastStaysLast: false);

        public static IReadOnlyList<string> PinDirectionsOf(BoardData? board) =>
            ComponentVocabulary.WithWordsOf(ComponentVocabulary.PinDirections, board?.ComponentPins.Select(pin => pin.Direction), lastStaysLast: true);

        private static IReadOnlyList<string> WithWordsOf(IReadOnlyList<string> words, IEnumerable<string?>? used, bool lastStaysLast)
        {
            List<string> extra = (used ?? [])
                .Select(word => word?.Trim() ?? string.Empty)
                .Where(word => word.Length > 0 && !words.Contains(word, StringComparer.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Order(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (extra.Count == 0)
            {
                return words;
            }

            return lastStaysLast
                ? [.. words.Take(words.Count - 1), .. extra, words[^1]]
                : [.. words, .. extra];
        }

        // Whether `value` is one of `words` - or empty, which every one of these allows.
        public static bool IsKnownOrEmpty(IReadOnlyList<string> words, string? value)
        {
            ArgumentNullException.ThrowIfNull(words);

            string trimmed = (value ?? string.Empty).Trim();

            return trimmed.Length == 0 || words.Contains(trimmed, StringComparer.OrdinalIgnoreCase);
        }
    }
}
