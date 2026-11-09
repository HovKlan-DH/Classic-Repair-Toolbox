using Handlers.DataHandling;
using System.Linq;
using static Handlers.DataHandling.ResistorColorCode;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// AN AXIAL CERAMIC CAPACITOR'S COLOR BANDS, FROM ITS VALUE (owner request, 2026-11-09: "since that has color rings, would
// it then be possible to show and translate those also, just alike a resistor?"; asked: "Every axial ceramic", a band the
// data does not say "Show it as not known"). The code is the one the owner read off the parts on his own 250407: C23,
// "360pF 50V 10%", orange | blue | brown | silver | silver, and C36, "20pF 50V", red | black | black | gold | white - two
// digits, the multiplier and the tolerance in the resistor's colors, then a fifth band the data does not say. The values
// are the published boards' own axial ceramics.
// ###########################################################################################
public sealed class CapacitorColorCodeTests
{
    // ###########################################################################################
    // The two parts read on the owner's board, as the data writes them: every band the data says is the part's color in
    // its place - and the ones it does not say (C23's fifth, C36's tolerance and fifth) are not known, never guessed.
    // ###########################################################################################
    [Theory]
    [InlineData("360pF 50V 10%", "Orange Blue Brown Silver Unknown")]
    [InlineData("20pF 50V", "Red Black Black Unknown Unknown")]
    public void The_parts_on_the_owners_board_read_as_he_read_them(string value, string colors)
    {
        Assert.Equal(colors, Colors(value));
    }

    private static string Colors(string value) =>
        CapacitorColorCode.TryRead(value, out CapacitorColorCode.Reading? reading) && CapacitorColorCode.Bands(reading!) is { } bands
            ? string.Join(" ", bands.Select(band => band.Color))
            : "none";

    private static ComponentEntry Capacitor(string value, string friendly = "Ceramic", string package = "Axial") =>
        new() { BoardLabel = "C83", Category = "Capacitor", FriendlyName = friendly, TechnicalNameOrValue = value, Package = package };

    // ###########################################################################################
    // Two digits, the multiplier in picofarads, the tolerance in the resistor's colors - not known when none is typed - and
    // the fifth band, not known, as the published values read.
    // ###########################################################################################
    [Theory]
    [InlineData("82pF 5%", "Grey Red Black Gold Unknown")]
    [InlineData("82pF", "Grey Red Black Unknown Unknown")]
    [InlineData("51pF 50V 5%", "Green Brown Black Gold Unknown")]
    [InlineData("15pF", "Brown Green Black Unknown Unknown")]
    [InlineData("470pF 50V 10%", "Yellow Violet Brown Silver Unknown")]
    [InlineData("1.8nF 50V 10%", "Brown Grey Red Silver Unknown")]
    [InlineData("10nF 25V", "Brown Black Orange Unknown Unknown")]
    [InlineData("100nF/50V", "Brown Black Yellow Unknown Unknown")]
    [InlineData("100pF 1%", "Brown Black Brown Brown Unknown")]
    [InlineData("100pF 2%", "Brown Black Brown Red Unknown")]
    public void The_bands_are_the_digits_the_multiplier_the_tolerance_and_a_fifth_not_known(string value, string colors)
    {
        Assert.Equal(colors, Colors(value));
    }

    // ###########################################################################################
    // What the code does not say for certain draws no bands: a value under 10 pF (none seen, and the multipliers under x1
    // differ between makers), one two digits cannot say, a multiplier past yellow, a tolerance with no color of its own (20%
    // has none on a resistor), and no capacitance at all.
    // ###########################################################################################
    [Theory]
    [InlineData("4.7pF")]
    [InlineData("8pF 5%")]
    [InlineData("125pF")]
    [InlineData("1uF 50V")]
    [InlineData("10uF 16V")]
    [InlineData("82pF 3%")]
    [InlineData("82pF 0.5%")]
    [InlineData("100pF 20%")]
    [InlineData("N/A")]
    [InlineData("")]
    public void What_the_code_does_not_say_for_certain_has_no_bands(string value)
    {
        Assert.Equal("none", Colors(value));
    }

    [Fact]
    public void Each_band_says_what_it_stands_for_and_the_colors_name_the_unknown_ones()
    {
        var bands = CapacitorColorCode.BandsOf(Capacitor("1.8nF 50V 10%"))!;

        Assert.Equal(["1", "8", "×100 pF", "± 10%", "?"], bands.Select(band => band.Label));
        Assert.Equal([BandKind.Digit, BandKind.Digit, BandKind.Multiplier, BandKind.Tolerance, BandKind.Extra], bands.Select(band => band.Kind));
        Assert.Equal([Band.Brown, Band.Grey, Band.Red, Band.Silver, Band.Unknown], bands.Select(band => band.Color));

        // No tolerance typed: its band is there, not known.
        var untyped = CapacitorColorCode.BandsOf(Capacitor("20pF 50V"))!;

        Assert.Equal(("± ?", BandKind.Tolerance, Band.Unknown), (untyped[3].Label, untyped[3].Kind, untyped[3].Color));
    }

    // ###########################################################################################
    // A COMMA BEFORE THREE DIGITS IS NO VALUE (code review, 2026-11-09): "10,000pF" read as 10 pF drew a 10 nF part's bands
    // a thousand times out. As a resistor's "1,000" (ResistorColorCode.CommaMayBeThousands): a thousands separator or a
    // decimal comma, a guess either way - no bands. A comma before one or two digits is a decimal.
    // ###########################################################################################
    [Theory]
    [InlineData("10,000pF", "none")]
    [InlineData("22,000 pF 5%", "none")]
    [InlineData("1,8nF 10%", "Brown Grey Red Silver Unknown")]
    public void A_comma_before_three_digits_is_no_value(string value, string colors)
    {
        Assert.Equal(colors, Colors(value));
    }

    // ###########################################################################################
    // A reading made by hand past what the bands say has none, as TryRead reads none (code review, 2026-11-09: a tolerance
    // with no band of its own threw from Bands, where the resistor's returns null).
    // ###########################################################################################
    [Theory]
    [InlineData(20)]
    [InlineData(3)]
    public void A_reading_made_by_hand_with_a_tolerance_no_band_says_has_no_bands(int percent)
    {
        Assert.Null(CapacitorColorCode.Bands(new CapacitorColorCode.Reading(100m, percent, string.Empty)));
    }

    // The bands and the reading they were made from, in one read - what the picture shows.
    [Fact]
    public void TryBandsOf_hands_back_the_reading_with_the_bands()
    {
        Assert.True(CapacitorColorCode.TryBandsOf(Capacitor("360pF 50V 10%"), out var bands, out var reading));
        Assert.Equal(5, bands!.Count);
        Assert.Equal((360m, (decimal?)10m, "50V"), (reading!.Picofarads, reading.TolerancePercent, reading.Voltage));

        Assert.False(CapacitorColorCode.TryBandsOf(Capacitor("4.7pF"), out bands, out _));
        Assert.Null(bands);
    }

    // The line under the colors says what is not known: the fifth band always, the tolerance only when none was typed.
    [Fact]
    public void The_note_says_the_fifth_band_and_a_tolerance_not_typed_are_not_known()
    {
        string typed = CapacitorColorCode.NoteOf(CapacitorColorCode.BandsOf(Capacitor("82pF 5%"))!);
        string untyped = CapacitorColorCode.NoteOf(CapacitorColorCode.BandsOf(Capacitor("82pF"))!);

        Assert.Contains("5th band", typed, System.StringComparison.Ordinal);
        Assert.DoesNotContain("tolerance", typed, System.StringComparison.Ordinal);
        Assert.Contains("tolerance, the 4th band", untyped, System.StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("82pF 5%", "82 pF 5%")]
    [InlineData("360pF 50V 10%", "360 pF 10%")]
    [InlineData("1.8nF 50V 10%", "1.8 nF 10%")]
    [InlineData("100nF/50V", "100 nF")]
    public void The_value_is_said_in_pF_or_nF_with_the_tolerance_typed(string value, string text)
    {
        Assert.True(CapacitorColorCode.TryRead(value, out CapacitorColorCode.Reading? reading));
        Assert.Equal(text, CapacitorColorCode.ValueText(reading!));
    }

    [Theory]
    [InlineData("360pF 50V 10%", "50V")]
    [InlineData("100nF/50V", "50V")]
    [InlineData("82pF 5%", "")]
    public void The_voltage_is_read_as_typed(string value, string voltage)
    {
        Assert.True(CapacitorColorCode.TryRead(value, out CapacitorColorCode.Reading? reading));
        Assert.Equal(voltage, reading!.Voltage);
    }

    // ###########################################################################################
    // Only an axial ceramic: its package "Axial", or with none its friendly name saying so (VIC-20 324003's "Ceramic
    // Axial"). A disc, a radial, a film or an electrolytic one, or no capacitor at all, gets none.
    // ###########################################################################################
    [Theory]
    [InlineData("Capacitor", "Ceramic", "Axial", true)]
    [InlineData("Capacitor", "Ceramic Axial", "", true)]
    [InlineData("Capacitor", "Ceramic", "", false)]
    [InlineData("Capacitor", "Ceramic", "Disc", false)]
    [InlineData("Capacitor", "Ceramic", "Radial", false)]
    [InlineData("Capacitor", "Film", "Axial", false)]
    [InlineData("Capacitor", "Ceramic film", "Axial", false)]
    [InlineData("Capacitor", "Electrolytic Axial", "Axial", false)]
    [InlineData("Capacitor", "Ceramic trimmer", "Axial", false)]
    [InlineData("Resistor", "Ceramic", "Axial", false)]
    public void Only_an_axial_ceramic_capacitor_has_bands(string category, string friendly, string package, bool banded)
    {
        var entry = new ComponentEntry { BoardLabel = "C1", Category = category, FriendlyName = friendly, Package = package, TechnicalNameOrValue = "82pF" };

        Assert.Equal(banded, CapacitorColorCode.BandsOf(entry) is not null);
    }
}
