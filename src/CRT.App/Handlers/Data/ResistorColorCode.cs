using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A RESISTOR'S COLOR BANDS, READ OFF THE VALUE TYPED FOR IT (owner request, 2026-11-07: "I create per resistor
    // value ... It does take me some time to make this image, so would it be possible for you to automatically create
    // such an image and show it under the pinout info in component info popup? ... This should of course be based on
    // the value typed in, so do check for all combinations in existing data").
    //
    // THE VALUES AS THE PUBLISHED DATA WRITES THEM (measured 2026-11-07 on every shipped board, 116 spellings of 1,600-odd
    // resistors): a number, a multiplier letter and the ohm sign - "560Ω", "1.5KΩ", "1MΩ", "470 Ω" - or the letter in
    // the number's place, as printed on boards ("4K7Ω", "8K2Ω", "1M5Ω"), then a power rating and a tolerance in any
    // order ("1/4W 5%", "1W 5%", "10%"). Read here: the resistance, a tolerance with a color of its own (0.05% to 20%),
    // and the power rating. Anything else - "wire", "N/A", "(unknown)", a lower-case "m" (milli or mega?), a tolerance no
    // band stands for, a comma before three digits (a thousands separator or a decimal?) - reads as no bands: a wrong
    // picture is worse than none.
    //
    // A NUMBER ALONE IS OHMS (owner request, 2026-11-07: "if it has any value, e.g. "100" or "100k" or "100 K" then it
    // should be able to read that as a resistor value also" - a board made from KiCad data writes "68", "100", "3.3k"):
    // only the bands of a component whose category is "Resistor" are drawn, so its value is a resistance. The number must
    // stand alone, nothing after it but a power rating and a tolerance - "2 resistors" or "1/4W" says no resistance.
    //
    // A SINGLE RESISTOR ONLY (BandsOf): a resistor pack - "RP" labels, "Resistor pack" (owner: "I assume you cannot
    // do the same for the resistor packs, as there are, to my knowledge, multiple types of those"; they are printed
    // with a number code, not bands) - a trimmer, a resistor array or network, has none.
    //
    // FOUR BANDS for a value two digits say - two digits, the multiplier, the tolerance (none for 20%) - and FIVE for
    // one that needs three, as a 1% resistor is printed: "249Ω" is red, yellow, white, black, then its tolerance. A value
    // three digits cannot say has no bands.
    // ###########################################################################################
    public static class ResistorColorCode
    {
        // The colors, as a resistor is printed - IEC 60062. Fixed, as the code is: the drawing is of the part.
        public enum Band
        {
            Black,
            Brown,
            Red,
            Orange,
            Yellow,
            Green,
            Blue,
            Violet,
            Grey,
            White,
            Gold,
            Silver,
        }

        // One band: its color, what it stands for as the picture's label says it, and what kind it is.
        public sealed record ColorBand(Band Color, string Label, BandKind Kind);

        public enum BandKind
        {
            Digit,
            Multiplier,
            Tolerance,
        }

        // What was read: the resistance in ohms, the tolerance in percent (null: none typed), the power rating as typed.
        public sealed record Reading(decimal Ohms, decimal? TolerancePercent, string Power);

        // The ohm sign either way it is typed: the published data's OHM SIGN (U+2126), or the Greek capital omega (U+03A9).
        private static readonly Regex ResistancePattern = new(
            @"^\s*(?<whole>\d+)(?:[.,](?<fraction>\d+))?\s*(?:(?<letter>[KMR])(?![a-z])(?<after>\d+)?)?\s*(?<ohm>[ΩΩ]|ohms?\b)?",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        // ###########################################################################################
        // Past what any bands say: three digits and white's x1G stay under a teraohm (code review, 2026-11-08). Checked
        // before any arithmetic - a long run of digits a decimal still holds overflowed it, times a million for "M" or
        // times a hundred in the bands' own search, and the exception escaped while the component popup was filled.
        // ###########################################################################################
        private const decimal MostOhms = 1_000_000_000_000m;

        // What says a component is no single resistor, in its names or package.
        private static readonly Regex NoSingleResistor = new(
            @"pack|array|network|trimmer|trim\s*pot|potentiometer|variable|\d\s*-?\s*pins?\b|\bsip\b",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        private static readonly Regex TolerancePattern = new(@"(?<!\d)(?<percent>\d+(?:[.,]\d+)?)\s*%", RegexOptions.CultureInvariant);

        private static readonly Regex PowerPattern = new(@"(?<!\d)(?<power>\d+(?:[./]\d+)?)\s*W\b", RegexOptions.CultureInvariant);

        // The tolerances a band stands for, and its color - 20% has no band.
        private static readonly (decimal Percent, Band? Color)[] Tolerances =
        [
            (0.05m, Band.Grey),
            (0.1m, Band.Violet),
            (0.25m, Band.Blue),
            (0.5m, Band.Green),
            (1m, Band.Brown),
            (2m, Band.Red),
            (5m, Band.Gold),
            (10m, Band.Silver),
            (20m, null),
        ];

        // ###########################################################################################
        // Whether `bands` are a metal-film resistor's - a tolerance of 2% or closer - which the picture draws with the light-blue
        // body such parts usually have; a carbon-film one, 5% or none, keeps the beige (owner, 2026-11-08: "ok, let us try this
        // out", to "a light-blue body for 1% and 2% resistors"). The closer tolerances are metal film too.
        // ###########################################################################################
        public static bool IsMetalFilm(IReadOnlyList<ColorBand> bands)
        {
            ArgumentNullException.ThrowIfNull(bands);

            return bands.Count > 0 &&
                   bands[^1].Kind == BandKind.Tolerance &&
                   Tolerances.Any(each => each.Color == bands[^1].Color && each.Percent <= 2m);
        }

        // ###########################################################################################
        // The bands a component gets drawn - null for none: a single resistor - its category "Resistor", no pack, array,
        // network, trimmer or potentiometer by its label, names or package - whose value reads (TryRead) as one the bands
        // can say.
        // ###########################################################################################
        public static IReadOnlyList<ColorBand>? BandsOf(ComponentEntry entry)
        {
            ArgumentNullException.ThrowIfNull(entry);

            if (!string.Equals(entry.Category?.Trim(), "Resistor", StringComparison.OrdinalIgnoreCase) || IsNoSingleResistor(entry))
            {
                return null;
            }

            return TryRead(entry.TechnicalNameOrValue, out Reading? reading) ? Bands(reading!) : null;
        }

        private static bool IsNoSingleResistor(ComponentEntry entry)
        {
            string label = entry.BoardLabel?.Trim() ?? string.Empty;

            if (label.StartsWith("RP", StringComparison.OrdinalIgnoreCase) || label.StartsWith("RN", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return NoSingleResistor.IsMatch(string.Join(' ', entry.FriendlyName, entry.TechnicalNameOrValue, entry.Package));
        }

        // ###########################################################################################
        // The resistance, tolerance and power rating out of a value as typed - false when it says no resistance the
        // bands can show (see the header).
        // ###########################################################################################
        public static bool TryRead(string? value, out Reading? reading)
        {
            reading = null;

            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            Match match = ResistancePattern.Match(value);

            if (!match.Success)
            {
                return false;
            }

            string letter = match.Groups["letter"].Value;
            bool hasOhm = match.Groups["ohm"].Success;

            // A lower-case "m" is milli in the units' own spelling and a slip for mega in a schematic's: no bands.
            if (letter == "m")
            {
                return false;
            }

            // A letter in the number's place ("4K7") takes no fraction before it.
            if (match.Groups["after"].Success && match.Groups["fraction"].Success)
            {
                return false;
            }

            // A comma before exactly three digits may as well be a thousands separator as a decimal comma (code review,
            // 2026-11-08: "1,000" and "2,200 Ohm" were drawn as 1 and 2.2 ohms) - a guess, so no bands. One or two digits
            // after it ("4,7K", "0,47") are a decimal.
            if (match.Groups["fraction"].Success && match.Groups["fraction"].Value.Length == 3 &&
                value[match.Groups["fraction"].Index - 1] == ',')
            {
                return false;
            }

            string rest = value[(match.Index + match.Length)..];

            // A number alone is ohms, as a resistor's value - when nothing but a power rating and a tolerance follows it
            // (see the header): "2 resistors" and "1/4W" are no resistance.
            if (letter.Length == 0 && !hasOhm && !IsOnlyRatingAndTolerance(rest))
            {
                return false;
            }

            string digits = match.Groups["fraction"].Success
                ? match.Groups["whole"].Value + "." + match.Groups["fraction"].Value
                : match.Groups["after"].Success
                    ? match.Groups["whole"].Value + "." + match.Groups["after"].Value
                    : match.Groups["whole"].Value;

            if (!decimal.TryParse(digits, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal number) ||
                number >= ResistorColorCode.MostOhms)
            {
                return false;
            }

            decimal ohms = letter.ToUpperInvariant() switch
            {
                "K" => number * 1_000m,
                "M" => number * 1_000_000m,
                _ => number,
            };

            if (ohms <= 0 || ohms >= ResistorColorCode.MostOhms)
            {
                return false;
            }

            decimal? tolerance = null;
            Match percent = TolerancePattern.Match(rest);

            if (percent.Success)
            {
                if (!decimal.TryParse(percent.Groups["percent"].Value.Replace(',', '.'), NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out decimal typed) ||
                    !Tolerances.Any(each => each.Percent == typed))
                {
                    return false;
                }

                tolerance = typed;
            }

            Match power = PowerPattern.Match(rest);

            reading = new Reading(ohms, tolerance, power.Success ? power.Groups["power"].Value + "W" : string.Empty);
            return true;
        }

        // Whether `rest` holds nothing but a power rating and a tolerance - "", " 5%", " 1/4W 5%", ", 1W".
        private static bool IsOnlyRatingAndTolerance(string rest)
        {
            string left = PowerPattern.Replace(TolerancePattern.Replace(rest, " "), " ");

            return left.All(character => char.IsWhiteSpace(character) || character is ',' or ';');
        }

        // ###########################################################################################
        // The bands for a reading, first to last: two digits and the multiplier - three digits for a value two cannot
        // say, and for a 2% or closer resistor, which is printed with five bands - then the tolerance's band when it has
        // one. Null for a value three digits cannot say, or a multiplier outside silver (x0.01) to white (x1G).
        // ###########################################################################################
        public static IReadOnlyList<ColorBand>? Bands(Reading reading)
        {
            ArgumentNullException.ThrowIfNull(reading);

            // A reading made by hand past what bands say has none, as TryRead reads none (MostOhms).
            if (reading.Ohms <= 0 || reading.Ohms >= ResistorColorCode.MostOhms)
            {
                return null;
            }

            int[] counts = reading.TolerancePercent is decimal close && close <= 2m ? [3] : [2, 3];

            foreach (int significant in counts)
            {
                if (SignificantDigits(reading.Ohms, significant) is not (string digits, int exponent))
                {
                    continue;
                }

                var bands = digits.Select(digit => new ColorBand((Band)(digit - '0'), digit.ToString(), BandKind.Digit)).ToList();
                bands.Add(new ColorBand(MultiplierColor(exponent), MultiplierLabel(exponent), BandKind.Multiplier));

                if (reading.TolerancePercent is decimal percent && Tolerances.First(each => each.Percent == percent).Color is Band color)
                {
                    bands.Add(new ColorBand(color, "± " + Number(percent) + "%", BandKind.Tolerance));
                }

                return bands;
            }

            return null;
        }

        // ###########################################################################################
        // `ohms` as `count` significant digits and a power of ten, exactly - "560" is "56" and 1; null when it needs more
        // digits, or the power is past the multiplier bands (-2 to 9).
        // ###########################################################################################
        internal static (string Digits, int Exponent)? SignificantDigits(decimal ohms, int count)
        {
            decimal low = Pow10(count - 1);
            decimal high = Pow10(count);

            for (int exponent = 9; exponent >= -2; exponent--)
            {
                decimal mantissa = ohms / Pow10(exponent);

                if (mantissa >= low && mantissa < high && mantissa == decimal.Truncate(mantissa))
                {
                    return (((long)mantissa).ToString(CultureInfo.InvariantCulture), exponent);
                }
            }

            return null;
        }

        private static decimal Pow10(int exponent)
        {
            decimal result = 1m;

            for (int i = 0; i < Math.Abs(exponent); i++)
            {
                result = exponent > 0 ? result * 10m : result / 10m;
            }

            return result;
        }

        private static Band MultiplierColor(int exponent) => exponent switch
        {
            -1 => Band.Gold,
            -2 => Band.Silver,
            _ => (Band)exponent,
        };

        // "x10 Ω", "x1 kΩ", "x0.1 Ω" - the multiplier as the picture's label says it.
        internal static string MultiplierLabel(int exponent)
        {
            decimal factor = Pow10(exponent);

            return "×" + (factor >= 1_000_000_000m ? Number(factor / 1_000_000_000m) + " GΩ"
                : factor >= 1_000_000m ? Number(factor / 1_000_000m) + " MΩ"
                : factor >= 1_000m ? Number(factor / 1_000m) + " kΩ"
                : Number(factor) + " Ω");
        }

        // ###########################################################################################
        // The value in words under the picture, as the owner wants it (2026-11-07: '"8.2K Ω 5%"', "8.2 kOhms 5%" until
        // then) - "560 Ω 5%", "8.2K Ω 5%", "1.5M Ω" - the tolerance only when typed.
        // ###########################################################################################
        public static string ValueText(Reading reading)
        {
            ArgumentNullException.ThrowIfNull(reading);

            string value = reading.Ohms >= 1_000_000m ? Number(reading.Ohms / 1_000_000m) + "M Ω"
                : reading.Ohms >= 1_000m ? Number(reading.Ohms / 1_000m) + "K Ω"
                : Number(reading.Ohms) + " Ω";

            return reading.TolerancePercent is decimal percent ? $"{value} {Number(percent)}%" : value;
        }

        // ###########################################################################################
        // Each band's color in words, first to last, the tolerance's included - the list under the value (owner request,
        // 2026-11-07: "Colors: Grey, Red, Red, Gold").
        // ###########################################################################################
        public static IReadOnlyList<string> ColorNames(IReadOnlyList<ColorBand> bands)
        {
            ArgumentNullException.ThrowIfNull(bands);

            return bands.Select(band => band.Color.ToString()).ToList();
        }

        private static string Number(decimal value) => value.ToString("0.###", CultureInfo.InvariantCulture);

        // ###########################################################################################
        // A band's color as the picture draws it, and the color its label is written in on it - white on the dark ones,
        // nearly black on the light. Hex, as the code fixes them.
        // ###########################################################################################
        public static (string Fill, string Text) ColorsOf(Band band) => band switch
        {
            Band.Black => ("#1E1E1E", "#FFFFFF"),
            Band.Brown => ("#5A2D22", "#FFFFFF"),
            Band.Red => ("#C62828", "#FFFFFF"),
            Band.Orange => ("#EF7C00", "#FFFFFF"),
            Band.Yellow => ("#F6C90E", "#1E1E1E"),
            Band.Green => ("#4C8C5A", "#FFFFFF"),
            Band.Blue => ("#11519A", "#FFFFFF"),
            Band.Violet => ("#7B3FA0", "#FFFFFF"),
            Band.Grey => ("#8A8A8A", "#FFFFFF"),
            Band.White => ("#F5F5F5", "#1E1E1E"),
            Band.Gold => ("#C1872B", "#FFFFFF"),
            Band.Silver => ("#B8B8B8", "#1E1E1E"),
            _ => ("#8A8A8A", "#FFFFFF"),
        };
    }
}
