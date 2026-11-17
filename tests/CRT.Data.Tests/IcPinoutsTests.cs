using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// IcPinouts - the other ICs' pin directions from their datasheets (owner decision, 2026-11-17): the catalogue reads, a
// component is matched by its number whatever is written around it, and a board saying something against the datasheet
// leaves the component without one.
// ###########################################################################################
public sealed class IcPinoutsTests
{
    // Every pin-out is a whole DIP with a ground, each pin named and in the pin list's words - "" only where no source settles.
    [Fact]
    public void Every_chip_is_a_whole_case_with_a_ground_and_every_direction_in_the_pin_lists_words()
    {
        Assert.NotEmpty(IcPinouts.All);

        foreach (IcPinout chip in IcPinouts.All)
        {
            Assert.True(chip.PinCount >= 14 && chip.PinCount % 2 == 0, $"{chip.Name}: {chip.PinCount} pins");
            Assert.Equal(chip.PinCount, chip.Directions.Count);
            Assert.All(chip.PinNames, name => Assert.False(string.IsNullOrWhiteSpace(name), $"{chip.Name}: a pin with no name"));
            Assert.All(chip.Directions, direction => Assert.True(direction.Length == 0 || ComponentVocabulary.PinDirections.Contains(direction), $"{chip.Name}: [{direction}]"));
            Assert.Contains("Ground", chip.Directions);
        }
    }

    // Two chips never share a name, and none is a logic chip LogicChips draws (it would never be reached).
    [Fact]
    public void Each_chip_is_found_by_its_own_name_and_none_is_a_logic_chip_CRT_draws()
    {
        Assert.Equal(IcPinouts.All.Count, IcPinouts.All.Select(chip => chip.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count());

        foreach (IcPinout chip in IcPinouts.All)
        {
            Assert.Same(chip, IcPinouts.Named(chip.Name));
            Assert.Null(LogicChips.Named(chip.Name));
        }
    }

    [Theory]
    [InlineData("6510", "6510")]
    [InlineData("8500", "6510")]
    [InlineData("MOS6526", "6526")]
    [InlineData("8521R4", "6526")]
    [InlineData("HD6845AP", "6845")]
    [InlineData("4164-2 DRAM", "4164")]
    [InlineData("4532", "4164")]
    [InlineData("4416-2 DRAM", "41464")]
    [InlineData("Z8400APS", "Z80")]
    [InlineData("Z80A", "Z80")]
    [InlineData("901227-03", "2364")]
    [InlineData("901225-01", "2332")]
    [InlineData("251913-01", "23128")]
    [InlineData("906114-01", "82S100")]
    [InlineData("SN74LS123N", "74123")]
    [InlineData("UPA2003C", "ULN2003")]
    [InlineData("M58725", "2016")]
    [InlineData("6551A", "6551")]
    [InlineData("7360/8360", "7360")]
    [InlineData("8566", "8564")]
    [InlineData("SN74LS629N", "74629")]
    [InlineData("4044", "MC4044")]
    [InlineData("MC4044P", "MC4044")]
    [InlineData("AY-3-8912", "AY-3-8912")]
    [InlineData("6529B", "6529")]
    [InlineData("251715-01", "251715")]
    // A micro sign or a Greek mu is the "U" of a maker's "uPD" (code review, 2026-11-17: upper-cased, the micro sign became a
    // capital Greek mu, and "µPD4164C" was no 4164).
    [InlineData("µPD4164C", "4164")]
    [InlineData("μPD4164C", "4164")]
    [InlineData("uPD4164C", "4164")]
    public void A_chip_is_found_whatever_is_written_around_its_number(string written, string chip)
    {
        Assert.Equal(chip, IcPinouts.Named(written)?.Name);
    }

    // A number that only starts like one CRT holds is another chip: a 65C02, a longer number, a word that is no chip - and a
    // CMOS CD4044 (a latch) is no MC4044, the phase detector a C64 board calls "4044", with its maker's letters or without.
    [Theory]
    [InlineData("CD4044BE")]
    [InlineData("4044BE")]
    // Another chip of a family, written as one word or two, is no CPU (code review, 2026-11-17).
    [InlineData("Z80 PIO")]
    [InlineData("Z80PIO")]
    [InlineData("Z80A CTC")]
    [InlineData("Z80SIO-2")]
    [InlineData("6510T")]
    [InlineData("MOS6510T")]
    [InlineData("65C02")]
    [InlineData("65100")]
    [InlineData("23640")]
    [InlineData("74LS08")]
    [InlineData("Phase detector")]
    [InlineData("")]
    [InlineData(null)]
    public void A_name_that_is_not_one_of_them_is_no_chip(string? written)
    {
        Assert.Null(IcPinouts.Named(written));
    }

    // Another chip of a family is told by a word of its own - never by letters running across two words (code review,
    // 2026-11-17: "Z80 BUS IO" joined read "SIO").
    [Theory]
    [InlineData("Z80 BUS IO", "Z80")]
    [InlineData("6510 TOD", "6510")]
    public void Letters_across_two_words_name_no_other_chip(string written, string chip)
    {
        Assert.Equal(chip, IcPinouts.Named(written)?.Name);
    }

    [Fact]
    public void The_directions_are_the_datasheets_in_the_pin_lists_words()
    {
        IcPinout cpu = IcPinouts.Named("6510")!;

        Assert.Equal(["Input", "Power", "Output", "Ground", "Input/output", "Input/output", "Output", "Output", "Input"],
            new[] { 1, 6, 7, 21, 24, 30, 38, 39, 40 }.Select(cpu.DirectionOf));
        Assert.Equal("A0", cpu.NameOf(7));
        Assert.Equal(string.Empty, cpu.DirectionOf(41));

        // A filter capacitor's pin is neither; the VIC's oscillator pins no source settles.
        Assert.Equal("Other", IcPinouts.Named("6581")!.DirectionOf(1));
        Assert.Equal(string.Empty, IcPinouts.Named("6561")!.DirectionOf(38));
    }

    // ###########################################################################################
    // THE BOARD MAY SAY NOTHING AGAINST THE DATASHEET: a package of another pin count, a pin past the case, or a ground or
    // supply named at a pin the datasheet gives a signal - and then the component has none. A supply wired to a pin the
    // datasheet leaves unconnected (the 250469's ROM, pin 1) and an empty package are no contradiction.
    // ###########################################################################################
    [Fact]
    public void A_board_saying_something_against_the_datasheet_leaves_the_component_without_one()
    {
        ComponentEntry Rom(string package) => new() { BoardLabel = "U4", TechnicalNameOrValue = "251913-01", Package = package };
        ComponentPinEntry Pin(string pin, string name) => new() { BoardLabel = "U4", Pin = pin, Name = name };

        Assert.NotNull(IcPinouts.For(Rom("DIP-28"), [Pin("1", "VCC"), Pin("14", "VSS"), Pin("28", "VCC")]));
        Assert.NotNull(IcPinouts.For(Rom(string.Empty), [Pin("11", "D0")]));

        Assert.Null(IcPinouts.For(Rom("DIP-24"), []));
        Assert.Null(IcPinouts.For(Rom("PLCC-28"), []));
        Assert.Null(IcPinouts.For(Rom(string.Empty), [Pin("29", "A15")]));
        Assert.Null(IcPinouts.For(Rom(string.Empty), [Pin("12", "GND")]));
        Assert.Null(IcPinouts.For(new ComponentEntry { BoardLabel = "U9", TechnicalNameOrValue = "XC9572" }, []));

        // The part number names it when the technical name does not.
        Assert.Equal("6526", IcPinouts.For(new ComponentEntry { BoardLabel = "U1", TechnicalNameOrValue = "CIA", PartNumber = "MOS6526" }, [])?.Name);
    }

    // ###########################################################################################
    // AS LogicChips JUDGES A LOGIC CHIP (code review, 2026-11-17): a ground named where the datasheet has a supply, and a
    // supply where it has a ground, are said against it too. A DIRECTION typed is not (the shipped boards' KiCad projects type
    // a ROM's chip selects strapped to the rails "Ground" and "Power"), nor a name saying the rail a pin is strapped to ("CE -
    // GND", the C16's PLA), nor anything at a pin the datasheet leaves unconnected.
    // ###########################################################################################
    [Fact]
    public void A_supply_named_where_the_datasheet_has_another_is_against_it()
    {
        ComponentEntry cia = new() { BoardLabel = "U1", TechnicalNameOrValue = "6526", Package = "DIP-40" };
        ComponentPinEntry Pin(string pin, string name, string direction = "") => new() { BoardLabel = "U1", Pin = pin, Name = name, Direction = direction };

        // Pin 1 is VSS (ground), pin 20 VCC, pin 5 PA3.
        Assert.Null(IcPinouts.For(cia, [Pin("20", "GND")]));
        Assert.Null(IcPinouts.For(cia, [Pin("1", "VCC")]));
        Assert.Null(IcPinouts.For(cia, [Pin("5", "+5V")]));

        Assert.NotNull(IcPinouts.For(cia, [Pin("1", "VSS", "Power"), Pin("20", "+5V", "Power"), Pin("5", "PA3", "Power"), Pin("6", "PA4 - GND", "Ground")]));
        Assert.NotNull(IcPinouts.For(new ComponentEntry { BoardLabel = "U4", TechnicalNameOrValue = "251913-01" }, [Pin("1", "GND", "Ground")]));
    }

    // ###########################################################################################
    // A RESISTOR, CAPACITOR, DIODE OR CONNECTOR IS NO IC (code review, 2026-11-17): a Bourns "4116R-1-103" resistor network in
    // a DIP-16 was read as a 4116 DRAM. Its category says what it is - any category starting with those words.
    // ###########################################################################################
    [Theory]
    [InlineData("Resistor")]
    [InlineData("resistor pack")]
    [InlineData("Capacitor")]
    [InlineData("Diode")]
    [InlineData("Connector")]
    public void A_component_of_a_category_no_IC_has_is_none_of_them(string category)
    {
        Assert.Null(IcPinouts.For(new ComponentEntry { BoardLabel = "RN1", TechnicalNameOrValue = "4116R-1-103", Category = category, Package = "DIP-16" }, []));
    }

    [Fact]
    public void A_component_whose_category_says_nothing_against_it_is_its_chip()
    {
        Assert.Equal("4116", IcPinouts.For(new ComponentEntry { BoardLabel = "U12", TechnicalNameOrValue = "4116-2", Category = "IC", Package = "DIP-16" }, [])?.Name);
        Assert.Equal("4116", IcPinouts.For(new ComponentEntry { BoardLabel = "U12", TechnicalNameOrValue = "4116-2" }, [])?.Name);
    }
}
