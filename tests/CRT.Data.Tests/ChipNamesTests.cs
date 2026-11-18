using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ChipNames - how the datasheets' checks read a chip's name and its supply pins: one rule for LogicChips and IcPinouts (code
// review, 2026-11-17: each held its own, and they had drifted apart).
// ###########################################################################################
public sealed class ChipNamesTests
{
    // The first word, in capitals - a micro sign or a Greek mu the "U" of a maker's "uPD" (upper-cased as it stands, the micro
    // sign became a capital Greek mu, which no rule knew).
    [Theory]
    [InlineData("  µPD4164C-15 DRAM", "UPD4164C-15")]
    [InlineData("μA555", "UA555")]
    [InlineData("mos6526", "MOS6526")]
    [InlineData("", "")]
    [InlineData("   ", "")]
    [InlineData(null, "")]
    public void The_first_word_is_read_in_capitals_with_a_micro_sign_as_U(string? text, string word) =>
        Assert.Equal(word, ChipNames.FirstWord(text));

    // A 74-series name's number as written - the leading zero kept, so "74LS008" is no "08".
    [Theory]
    [InlineData("SN74LS08N", "08")]
    [InlineData("7408", "08")]
    [InlineData("HD74LS629P", "629")]
    [InlineData("UPD74HC00C", "00")]
    [InlineData("74LS008", "008")]
    [InlineData("6510", null)]
    [InlineData("74", null)]
    [InlineData("", null)]
    public void A_74_series_name_gives_its_number_as_written(string word, string? number) =>
        Assert.Equal(number, ChipNames.SeventyFourNumber(word));

    // Its family as written, in capitals - "" for the plain one - for the numbers whose families differ in their pins (the 51).
    [Theory]
    [InlineData("SN74LS51N", "LS")]
    [InlineData("SN7451N", "")]
    [InlineData("74S51", "S")]
    [InlineData("DM74ALS51N", "ALS")]
    [InlineData("74hc51", "HC")]
    [InlineData("6510", null)]
    public void A_74_series_name_gives_its_family_as_written(string word, string? family) =>
        Assert.Equal(family, ChipNames.SeventyFourFamily(word));

    // ###########################################################################################
    // WHOLE NAMES ONLY: a pin named for the rail it is strapped to as well as its signal - the C16's PLA has "CE - GND" on its
    // /CE input - is no ground pin, or the chip lost its datasheet (measured on the shipped boards, 2026-11-17).
    // ###########################################################################################
    [Theory]
    [InlineData("GND", true)]
    [InlineData(" gnd ", true)]
    [InlineData("VSS", true)]
    [InlineData("0V", true)]
    [InlineData("Ground", true)]
    [InlineData("earth", true)]
    [InlineData("CE - GND", false)]
    [InlineData("AGND", false)]
    [InlineData("VCC", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void A_ground_is_told_by_its_whole_name(string? name, bool ground) =>
        Assert.Equal(ground, ChipNames.IsGround(name));

    [Theory]
    [InlineData("VCC", true)]
    [InlineData("vdd", true)]
    [InlineData("VBB", true)]
    [InlineData("VPP", true)]
    [InlineData("+VCC", true)]
    [InlineData("V+", true)]
    [InlineData("+5V", true)]
    [InlineData("5V", true)]
    [InlineData("+12V", true)]
    [InlineData("9V", true)]
    [InlineData("+5VDC", true)]
    [InlineData("-5V", false)]
    [InlineData("0V", false)]
    [InlineData("GND", false)]
    [InlineData("VCC1", false)]
    [InlineData("CS3 - VCC", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void A_supply_is_told_by_its_whole_name(string? name, bool supply) =>
        Assert.Equal(supply, ChipNames.IsSupply(name));

    // A resistor, capacitor, diode or connector is no chip, whatever its value reads as - a category starting with those words
    // too; any other category, or none, says nothing against one.
    [Theory]
    [InlineData("Resistor", true)]
    [InlineData(" resistor pack", true)]
    [InlineData("Capacitor", true)]
    [InlineData("Diode", true)]
    [InlineData("Connector", true)]
    [InlineData("IC", false)]
    [InlineData("Transistor", false)]
    [InlineData("Misc", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void A_category_no_chip_has_is_told(string? category, bool noChip) =>
        Assert.Equal(noChip, ChipNames.IsNoChip(category));

    // ###########################################################################################
    // A package says nothing against the chip's own case - the same layout and count, or nothing at all, left empty. A
    // surface-mounted case numbered as a DIP fits a DIP's pin-out (owner request, 2026-11-18: the Amiga "has a lot of SMD
    // components"); a PLCC is no QFP of as many pins, nor a DIP.
    // ###########################################################################################
    [Theory]
    [InlineData("DIP-14", "DIP-14", true)]
    [InlineData(" DIP14 ", "DIP-14", true)]
    [InlineData("SOIC-14", "DIP-14", true)]
    [InlineData("DIP-28 wide", "DIP-28", true)]
    [InlineData("", "DIP-14", true)]
    [InlineData(null, "DIP-40", true)]
    [InlineData("QFP-44", "QFP-44", true)]
    [InlineData("LQFP-44", "QFP-44", true)]
    [InlineData("DIP-16", "DIP-14", false)]
    [InlineData("SIP-14", "DIP-14", false)]
    [InlineData("PLCC-28", "DIP-28", false)]
    [InlineData("PLCC-44", "QFP-44", false)]
    [InlineData("DIP-44", "QFP-44", false)]
    public void A_package_fits_the_chips_own_case(string? package, string chipPackage, bool fits) =>
        Assert.Equal(fits, ChipNames.PackageFits(new ComponentEntry { Package = package! }, chipPackage));
}
