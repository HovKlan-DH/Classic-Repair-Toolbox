using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;
using static Handlers.DataHandling.ResistorColorCode;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A RESISTOR'S COLOR BANDS, FROM THE VALUE TYPED FOR IT (owner request, 2026-11-07: "automatically create such an
// image ... This should of course be based on the value typed in, so do check for all combinations in existing
// data"). The spellings below are the published data's own (measured 2026-11-07 on every shipped board: the ohm sign
// there is U+2126), and the colors IEC 60062's.
// ###########################################################################################
public sealed class ResistorColorCodeTests
{
    private const string Ohm = "Ω";

    private static string Colors(string value) =>
        TryRead(value, out Reading? reading) && Bands(reading!) is { } bands
            ? string.Join(" ", bands.Select(band => band.Color))
            : "none";

    // Whether the component popup draws the component's bands - BandsOf, as it asks.
    private static bool HasBands(ComponentEntry? entry) => entry is not null && BandsOf(entry) is not null;

    private static ComponentEntry Resistor(string value, string label = "R12", string friendly = "", string package = "") =>
        new() { BoardLabel = label, Category = "Resistor", TechnicalNameOrValue = value, FriendlyName = friendly, Package = package };

    // ###########################################################################################
    // Every way the published data writes a resistance reads - the number with K or M, the letter in the number's place
    // as printed on boards, a space before the ohm sign, the power rating and the tolerance after it in any order.
    // ###########################################################################################
    [Theory]
    [InlineData("560" + Ohm + " 5%", "Green Blue Brown Gold")]
    [InlineData("560" + Ohm + " 1/4W 5%", "Green Blue Brown Gold")]
    [InlineData("560" + Ohm, "Green Blue Brown")]
    [InlineData("470 " + Ohm, "Yellow Violet Brown")]
    [InlineData("1.5K" + Ohm + " 1/4W 5%", "Brown Green Red Gold")]
    [InlineData("4K7" + Ohm + " 1/4W 5%", "Yellow Violet Red Gold")]
    [InlineData("4.7K" + Ohm + " 5%", "Yellow Violet Red Gold")]
    [InlineData("8K2" + Ohm + " 1/4W 5%", "Grey Red Red Gold")]
    [InlineData("1M" + Ohm + " 1/4W 5%", "Brown Black Green Gold")]
    [InlineData("1M5" + Ohm, "Brown Green Green")]
    [InlineData("10" + Ohm + " 1/2W 5%", "Brown Black Black Gold")]
    [InlineData("100K" + Ohm + " 5%", "Brown Black Yellow Gold")]
    [InlineData("3.3K" + Ohm + " 10%", "Orange Orange Red Silver")]
    [InlineData("110" + Ohm + " 1W 5%", "Brown Brown Brown Gold")]
    [InlineData("1K" + Ohm + " 5W 5%", "Brown Black Red Gold")]
    [InlineData("1" + Ohm + " 5%", "Brown Black Gold Gold")]
    [InlineData("0.47" + Ohm + " 5%", "Yellow Violet Silver Gold")]
    [InlineData("4R7 5%", "Yellow Violet Gold Gold")]
    [InlineData("560R", "Green Blue Brown")]
    [InlineData("2.2k ohm", "Red Red Red")]
    [InlineData("560 Ohms 5%", "Green Blue Brown Gold")]
    [InlineData("560Ω", "Green Blue Brown")]
    public void Every_spelling_of_the_published_data_reads_as_its_bands(string value, string expected)
    {
        Assert.Equal(expected, Colors(value));
    }

    // ###########################################################################################
    // A NUMBER ALONE IS OHMS (owner request, 2026-11-07: '"100" or "100k" or "100 K" ... should be able to read that as a
    // resistor value') - the spellings of the owner's own board made from KiCad data ("68", "100", "3.3k", "1.2k"), a
    // letter in either case, a space before it, and a rating and tolerance after the number.
    // ###########################################################################################
    [Theory]
    [InlineData("68", "Blue Grey Black")]
    [InlineData("100", "Brown Black Brown")]
    [InlineData("47", "Yellow Violet Black")]
    [InlineData("100k", "Brown Black Yellow")]
    [InlineData("100 K", "Brown Black Yellow")]
    [InlineData("3.3k", "Orange Orange Red")]
    [InlineData("1.2k", "Brown Red Red")]
    [InlineData("470k", "Yellow Violet Yellow")]
    [InlineData("100 5%", "Brown Black Brown Gold")]
    [InlineData("100 1/4W 5%", "Brown Black Brown Gold")]
    public void A_number_alone_is_ohms_and_a_letter_in_either_case_its_multiplier(string value, string expected)
    {
        Assert.Equal(expected, Colors(value));
    }

    // ###########################################################################################
    // FIVE BANDS for a 2% or closer resistor - printed so - and for a value two digits cannot say; a 20% resistor has no
    // tolerance band at all.
    // ###########################################################################################
    [Theory]
    [InlineData("1K" + Ohm + " 1%", "Brown Black Black Brown Brown")]
    [InlineData("33" + Ohm + " 2%", "Orange Orange Black Gold Red")]
    [InlineData("249" + Ohm + " 5%", "Red Yellow White Black Gold")]
    [InlineData("10K" + Ohm + " 20%", "Brown Black Orange")]
    [InlineData("1K" + Ohm + " 0.1%", "Brown Black Black Brown Violet")]
    public void A_close_tolerance_or_a_third_digit_gives_five_bands_and_20_percent_none_for_tolerance(string value, string expected)
    {
        Assert.Equal(expected, Colors(value));
    }

    // ###########################################################################################
    // NO BANDS where they would be a guess: the published data's "wire", "N/A" and "(unknown)", a number with words
    // after it, a power rating alone, a lower-case "m" (milli, or a slip for mega?), a tolerance no band stands for, a
    // value four digits need, a word starting with R after the number.
    // ###########################################################################################
    [Theory]
    [InlineData("wire")]
    [InlineData("N/A")]
    [InlineData("(unknown)")]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("1/4W")]
    [InlineData("47 see note")]
    [InlineData("1m" + Ohm)]
    [InlineData("1K" + Ohm + " 3%")]
    [InlineData("1234" + Ohm)]
    [InlineData("2 resistors")]
    [InlineData("1.5K7" + Ohm)]
    public void A_value_the_bands_cannot_say_for_sure_has_none(string? value)
    {
        Assert.Equal("none", Colors(value!));
    }

    // ###########################################################################################
    // *** A COMMA BEFORE THREE DIGITS MAY BE A THOUSANDS SEPARATOR, SO IT GIVES NO BANDS (code review, 2026-11-08). ***
    // "1,000" and "2,200 Ohm" read as 1 and 2.2 ohms and were drawn brown-black-gold and red-red-gold - confidently
    // wrong, where a wrong picture is worse than none. A decimal comma before one or two digits stays the decimal it is.
    // ###########################################################################################
    [Theory]
    [InlineData("1,000")]
    [InlineData("2,200 Ohm")]
    [InlineData("4,700" + Ohm + " 5%")]
    [InlineData("1,500K")]
    public void A_comma_before_three_digits_gives_no_bands(string value)
    {
        Assert.Equal("none", Colors(value));
    }

    // ###########################################################################################
    // *** A VALUE PAST WHAT ANY BANDS SAY HAS NONE - AND NEVER THROWS (code review, 2026-11-08). *** A long run of digits
    // that a decimal still holds overflowed the arithmetic - "99999999999999999999999M" times a million, or the bands'
    // own search times a hundred - and the exception escaped while the component popup was being filled. Three digits
    // and white's x1G say under a teraohm; anything from there up is no picture.
    // ###########################################################################################
    [Theory]
    [InlineData("99999999999999999999999M")]
    [InlineData("79228162514264337593543950335")]
    [InlineData("7922816251426433759354395033" + Ohm)]
    [InlineData("1000000000000")]
    [InlineData("1000000M")]
    public void A_value_past_what_any_bands_say_has_none_and_never_throws(string value)
    {
        Assert.Equal("none", Colors(value));
        Assert.False(HasBands(Resistor(value)));
    }

    // A reading handed in whole, past what bands say, has none either.
    [Fact]
    public void A_reading_past_what_any_bands_say_has_no_bands()
    {
        Assert.Null(Bands(new Reading(decimal.MaxValue, null, string.Empty)));
        Assert.Null(Bands(new Reading(1_000_000_000_000m, 5m, string.Empty)));
        Assert.NotNull(Bands(new Reading(470_000_000_000m, 5m, string.Empty)));
    }

    [Theory]
    [InlineData("4,7K" + Ohm, "Yellow Violet Red")]
    [InlineData("0,47" + Ohm + " 5%", "Yellow Violet Silver Gold")]
    [InlineData("2,2k", "Red Red Red")]
    public void A_decimal_comma_before_one_or_two_digits_reads_as_a_decimal(string value, string expected)
    {
        Assert.Equal(expected, Colors(value));
    }

    // ###########################################################################################
    // A SINGLE RESISTOR ONLY (the owner: "I assume you cannot do the same for the resistor packs, as there are, to my
    // knowledge, multiple types of those") - the published packs are "RP" labels or "Resistor pack" with a pin count;
    // a trimmer, an array or a network has none either; nor does anything but a resistor.
    // ###########################################################################################
    [Fact]
    public void Only_a_single_resistor_has_bands_never_a_pack_a_trimmer_or_another_part()
    {
        Assert.True(HasBands(Resistor("560" + Ohm + " 5%")));

        Assert.False(HasBands(Resistor("3.3K" + Ohm, label: "RP3")));
        Assert.False(HasBands(Resistor("1K" + Ohm + ", 6pin", label: "R1", friendly: "Resistor pack")));
        Assert.False(HasBands(Resistor("68" + Ohm + ", 8pin 4 pair isolated", label: "R2")));
        Assert.False(HasBands(Resistor("10K" + Ohm, friendly: "Trimmer resistor")));
        Assert.False(HasBands(Resistor("10K" + Ohm, friendly: "Resistor network")));
        Assert.False(HasBands(Resistor("10K" + Ohm, package: "SIP-9")));
        Assert.False(HasBands(Resistor("10K" + Ohm, label: "RN1")));
        Assert.False(HasBands(new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", TechnicalNameOrValue = "560" + Ohm }));
        Assert.False(HasBands(null));

        // Its friendly name saying "Resistor" is no pack - the published data names 29 single ones so.
        Assert.True(HasBands(Resistor("4K7" + Ohm + " 1/4W 5%", friendly: "Resistor", package: "Axial")));
    }

    // The labels the picture's pills carry - the owner's own: "5", "6", "x10 Ω", "± 5%".
    [Fact]
    public void The_labels_are_a_digit_the_multiplier_and_the_tolerance()
    {
        Assert.True(TryRead("560" + Ohm + " 5%", out Reading? reading));

        IReadOnlyList<ColorBand> bands = Bands(reading!)!;

        Assert.Equal(["5", "6", "×10 Ω", "± 5%"], bands.Select(band => band.Label));
        Assert.Equal([BandKind.Digit, BandKind.Digit, BandKind.Multiplier, BandKind.Tolerance], bands.Select(band => band.Kind));
    }

    [Theory]
    [InlineData(-2, "×0.01 Ω")]
    [InlineData(-1, "×0.1 Ω")]
    [InlineData(0, "×1 Ω")]
    [InlineData(2, "×100 Ω")]
    [InlineData(3, "×1 kΩ")]
    [InlineData(5, "×100 kΩ")]
    [InlineData(6, "×1 MΩ")]
    [InlineData(9, "×1 GΩ")]
    public void The_multiplier_reads_in_ohms_kilohms_megohms_and_gigohms(int exponent, string expected)
    {
        Assert.Equal(expected, MultiplierLabel(exponent));
    }

    // The value in words under the picture, as the owner wants it (2026-11-07) - "8.2K Ω 5%".
    [Theory]
    [InlineData("560" + Ohm + " 1/4W 5%", "560 Ω 5%", "1/4W")]
    [InlineData("8K2" + Ohm + " 1/4W 5%", "8.2K Ω 5%", "1/4W")]
    [InlineData("4K7" + Ohm, "4.7K Ω", "")]
    [InlineData("1M5" + Ohm + " 10%", "1.5M Ω 10%", "")]
    [InlineData("0.47" + Ohm + " 5%", "0.47 Ω 5%", "")]
    [InlineData("330" + Ohm + " 1W 5%", "330 Ω 5%", "1W")]
    public void The_value_is_said_in_words_with_its_power_rating_apart(string value, string words, string power)
    {
        Assert.True(TryRead(value, out Reading? reading));

        Assert.Equal(words, ValueText(reading!));
        Assert.Equal(power, reading!.Power);
    }

    // The colors listed under the value, first to last, the tolerance's included - the owner's example, "8.2K Ω 5%".
    [Fact]
    public void The_colors_are_named_band_by_band_the_tolerance_last()
    {
        Assert.True(TryRead("8K2" + Ohm + " 5%", out Reading? reading));

        Assert.Equal(["Grey", "Red", "Red", "Gold"], ColorNames(Bands(reading!)!));
    }

    // Two digits exactly, at the power of ten that makes them - or none when it takes more.
    [Theory]
    [InlineData(560, 2, "56", 1)]
    [InlineData(10, 2, "10", 0)]
    [InlineData(1, 2, "10", -1)]
    [InlineData(1000000, 2, "10", 5)]
    [InlineData(249, 3, "249", 0)]
    public void A_value_is_its_digits_and_a_power_of_ten(int ohms, int count, string digits, int exponent)
    {
        Assert.Equal((digits, exponent), SignificantDigits(ohms, count));
    }

    [Fact]
    public void A_value_its_digits_cannot_say_exactly_has_none()
    {
        Assert.Null(SignificantDigits(249, 2));
        Assert.Null(SignificantDigits(0.001m, 2));
    }

    // Every band has a color to draw and a color for its label.
    [Fact]
    public void Every_band_has_its_colors()
    {
        foreach (Band band in System.Enum.GetValues<Band>())
        {
            (string fill, string text) = ColorsOf(band);

            Assert.Matches("^#[0-9A-F]{6}$", fill);
            Assert.Matches("^#[0-9A-F]{6}$", text);
            Assert.NotEqual(fill, text);
        }
    }
}
