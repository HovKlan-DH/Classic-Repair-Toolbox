using System;
using System.Collections.Generic;
using System.Linq;
using static Handlers.DataHandling.ResistorColorCode;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // AN AXIAL CERAMIC CAPACITOR'S COLOR BANDS, READ OFF ITS VALUE, AS A RESISTOR'S ARE (owner request, 2026-11-09: "since
    // that has color rings, would it then be possible to show and translate those also, just alike a resistor?"; asked,
    // "Every axial ceramic", and a band the data does not say "Show it as not known").
    //
    // THE CODE IS READ OFF THE PARTS ON THE OWNER'S OWN 250407 (2026-11-09): C23, "360pF 50V 10%", is orange | blue | brown |
    // silver | silver, and C36, "20pF 50V", red | black | black | gold | white - FIVE bands: two digits, the multiplier in
    // picofarads and the tolerance, all in the RESISTOR'S colors (gold 5%, silver 10% - not the ceramic capacitor code's
    // green and white), then a fifth band that is not the same from part to part. Both are 50 V and their fifth bands
    // differ, so it is no voltage - most likely how the capacitance changes with temperature, which the data never says:
    // drawn as NOT KNOWN (Band.Unknown, BandKind.Extra), the part having it. A tolerance the value does not type is not known
    // either, as the part has that band too (C36). A first reading of the board's photo as a ceramic code with the
    // temperature band FIRST was wrong; read so, the photo is green | brown | black ... = 51 pF, the board's C38.
    //
    // ONLY WHAT THE CODE SAYS FOR CERTAIN, else no bands (and PartPicture's plain tube with its printed code): a value of
    // 10 pF or more that two digits say, its multiplier black (x1) to yellow (x10 nF) - nothing under 10 pF has been seen -
    // and a tolerance typed only when it has a color: brown 1%, red 2%, gold 5%, silver 10%. The value is read by
    // PartPicture.Picofarads and the tolerance by ResistorColorCode.TryReadTolerance - the readers the other pictures use.
    // ###########################################################################################
    public static class CapacitorColorCode
    {
        // What was read: the capacitance in picofarads, the tolerance in percent (null: none typed), the voltage as typed.
        public sealed record Reading(decimal Picofarads, decimal? TolerancePercent, string Voltage);

        // The tolerances a band stands for, and its color - the resistor's.
        private static readonly (decimal Percent, Band Color)[] Tolerances =
        [
            (1m, Band.Brown),
            (2m, Band.Red),
            (5m, Band.Gold),
            (10m, Band.Silver),
        ];

        // The multipliers the bands say for certain: black x1 pF to yellow x10 nF.
        private const int LowestExponent = 0;
        private const int HighestExponent = 4;

        // ###########################################################################################
        // The bands a component gets drawn - null for none: an axial ceramic capacitor (PartPicture.IsAxialCeramicCapacitor)
        // whose value reads (TryRead) as one the bands say.
        // ###########################################################################################
        public static IReadOnlyList<ColorBand>? BandsOf(ComponentEntry entry) => TryBandsOf(entry, out IReadOnlyList<ColorBand>? bands, out _) ? bands : null;

        // BandsOf, with the reading the bands were made from - the value read once.
        public static bool TryBandsOf(ComponentEntry entry, out IReadOnlyList<ColorBand>? bands, out Reading? reading)
        {
            ArgumentNullException.ThrowIfNull(entry);

            bands = null;
            reading = null;

            if (!PartPicture.IsAxialCeramicCapacitor(entry) || !TryRead(entry.TechnicalNameOrValue, out reading))
            {
                return false;
            }

            bands = Bands(reading!);
            return bands is not null;
        }

        // ###########################################################################################
        // The capacitance, tolerance and voltage out of a value as typed - "82pF 5%", "360pF 50V 10%", "100nF/50V" - false
        // when it says no capacitance, or a tolerance typed that no band stands for.
        // ###########################################################################################
        public static bool TryRead(string? value, out Reading? reading)
        {
            reading = null;

            if (PartPicture.Picofarads(value) is not decimal picofarads ||
                !TryReadTolerance(value!, HasBand, out decimal? tolerance))
            {
                return false;
            }

            reading = new Reading(picofarads, tolerance, PartPicture.Voltage(value));
            return true;
        }

        private static bool HasBand(decimal percent) => Tolerances.Any(each => each.Percent == percent);

        // ###########################################################################################
        // The five bands for a reading, first to last: two digits, the multiplier, the tolerance - not known when none was
        // typed - and the fifth band, not known. Null for a value under 10 pF, one two digits cannot say, a multiplier past
        // black to yellow - or, for a reading made by hand, a tolerance no band stands for, as TryRead reads none (code
        // review, 2026-11-09: it threw).
        // ###########################################################################################
        public static IReadOnlyList<ColorBand>? Bands(Reading reading)
        {
            ArgumentNullException.ThrowIfNull(reading);

            if (reading.Picofarads < 10m ||
                (reading.TolerancePercent is decimal unbanded && !HasBand(unbanded)) ||
                SignificantDigits(reading.Picofarads, 2) is not (string digits, int exponent) ||
                exponent is < LowestExponent or > HighestExponent)
            {
                return null;
            }

            var bands = digits.Select(digit => new ColorBand((Band)(digit - '0'), digit.ToString(), BandKind.Digit)).ToList();
            bands.Add(new ColorBand((Band)exponent, MultiplierLabel(exponent), BandKind.Multiplier));

            bands.Add(reading.TolerancePercent is decimal percent
                ? new ColorBand(Tolerances.First(each => each.Percent == percent).Color, "± " + Number(percent) + "%", BandKind.Tolerance)
                : new ColorBand(Band.Unknown, "± ?", BandKind.Tolerance));

            bands.Add(new ColorBand(Band.Unknown, "?", BandKind.Extra));
            return bands;
        }

        // ###########################################################################################
        // What the picture says under the colors of the bands it cannot name - the fifth always, the tolerance when none was
        // typed.
        // ###########################################################################################
        public static string NoteOf(IReadOnlyList<ColorBand> bands)
        {
            ArgumentNullException.ThrowIfNull(bands);

            bool toleranceUnknown = bands.Any(band => band.Kind == BandKind.Tolerance && band.Color == Band.Unknown);

            return "The 5th band is not the same from maker to maker - most likely how the capacitance changes with temperature - " +
                   "and the data does not say it." +
                   (toleranceUnknown ? " Nor does it say the tolerance, the 4th band." : string.Empty);
        }

        // "×1 pF", "×10 pF", "×100 pF", "×1 nF", "×10 nF" - the multiplier as the picture's label says it.
        internal static string MultiplierLabel(int exponent) => exponent switch
        {
            0 => "×1 pF",
            1 => "×10 pF",
            2 => "×100 pF",
            3 => "×1 nF",
            _ => "×10 nF",
        };

        // ###########################################################################################
        // The value in words under the picture, as a resistor's is said - "82 pF 5%", "1.8 nF 10%", "100 nF" - the tolerance
        // only when typed.
        // ###########################################################################################
        public static string ValueText(Reading reading)
        {
            ArgumentNullException.ThrowIfNull(reading);

            string value = reading.Picofarads >= 1_000_000m ? Number(reading.Picofarads / 1_000_000m) + " µF"
                : reading.Picofarads >= 1_000m ? Number(reading.Picofarads / 1_000m) + " nF"
                : Number(reading.Picofarads) + " pF";

            return reading.TolerancePercent is decimal percent ? $"{value} {Number(percent)}%" : value;
        }
    }
}
