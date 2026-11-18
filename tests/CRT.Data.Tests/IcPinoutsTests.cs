using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// IcPinouts - the other ICs' pin directions from their datasheets (owner decision, 2026-11-17): the catalogue reads, a
// component is matched by its number whatever is written around it, and a board saying something against the datasheet
// leaves the component without one.
// ###########################################################################################
public sealed class IcPinoutsTests
{
    // ###########################################################################################
    // Every pin-out is a whole case with a ground, each pin named and in the pin list's words - "" only where no source
    // settles - and a DIP of its pins (the package the conversion writes, owner request 2026-11-17), but the CX20185's
    // QFP-44 and the Fat Agnus's PLCC-84 (2026-11-18).
    // ###########################################################################################
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
            // A quad op-amp has its negative supply where a chip has its ground (the LF347's pin 11).
            Assert.True(chip.Directions.Contains("Ground") || chip.PinNames.Contains("VCC-"), $"{chip.Name}: no ground");
            Assert.Equal(
                chip.Name switch
                {
                    "CX20185" => "QFP-44",
                    "8370" or "8372" => "PLCC-84",
                    _ => $"DIP-{chip.PinCount}",
                },
                chip.Package);
            Assert.Equal(chip.PinCount, ComponentVocabulary.PinCountOf(chip.Package));
        }
    }

    // ###########################################################################################
    // THE C128 250477's AND THE CPC 664's CHIPS (owner request, 2026-11-17), as their sources give them: the 4464 as the
    // owner's picture names it, the 8568 as the C128 service manual's page (/CAS on pin 1, /INTR on 9, RES on 22 and /INIT on
    // 23 - the 8563's pins differ), the 5710 as William Levak's sheet, the Gate Array as the board's picture, the CX20185 as
    // the manual's read/write amplifier page.
    // ###########################################################################################
    [Fact]
    public void The_C128DCRs_and_the_CPC_664s_chips_read_as_their_sources()
    {
        IcPinout dram = IcPinouts.Named("4464")!;

        Assert.Equal(["/G", "D0", "D1", "/W", "/RAS", "A6", "A5", "A4", "VCC", "A7", "A3", "A2", "A1", "A0", "D2", "/CAS", "D3", "VSS"], dram.PinNames);
        Assert.Equal("DIP-18", dram.Package);

        IcPinout vdc = IcPinouts.Named("8568")!;

        Assert.Equal(48, vdc.PinCount);
        Assert.Equal(["/CAS", "CCLK", "DCLK", "HSYNC", "/INTR", "VSS", "CSYNC", "RES", "/INIT", "VCC", "CV", "/RAS"], new[] { 1, 2, 3, 4, 9, 12, 19, 22, 23, 37, 47, 48 }.Select(vdc.NameOf));
        Assert.Equal(["Output", "Input", "Input"], new[] { 9, 22, 23 }.Select(vdc.DirectionOf));
        Assert.NotEqual(IcPinouts.Named("8563")!.NameOf(1), vdc.NameOf(1));

        IcPinout fdc = IcPinouts.Named("5710")!;

        Assert.Equal(48, fdc.PinCount);
        Assert.Equal(["SRQ OUT", "CK", "A10", "A4", "A0", "16MHz", "/VIA1", "WD", "/INDEX", "NC", "/DATA OUT"], new[] { 1, 4, 22, 23, 27, 30, 34, 37, 41, 42, 48 }.Select(fdc.NameOf));
        Assert.Equal(["Output", "", "Input", "Output", "Not connected"], new[] { 1, 2, 4, 30, 43 }.Select(fdc.DirectionOf));

        IcPinout gateArray = IcPinouts.Named("40008/40010")!;

        Assert.Equal(40, gateArray.PinCount);
        Assert.Equal(["D5", "CCLK", "/N244EN", "CK16", "D0", "VSS3", "D4"], new[] { 1, 4, 23, 24, 35, 36, 40 }.Select(gateArray.NameOf));
        Assert.Equal(["Input", "Output", "Output", "Input", "Ground"], new[] { 1, 4, 23, 24, 36 }.Select(gateArray.DirectionOf));

        IcPinout amplifier = IcPinouts.Named("CX20185")!;

        Assert.Equal(44, amplifier.PinCount);
        Assert.Equal(["DIFF CONSTANT B", "GAIN SEL A", "HEAD0 A", "A.GND", "/WRITE DATA", "READ DATA", "DIFF OUT A"], new[] { 1, 9, 12, 16, 28, 33, 44 }.Select(amplifier.NameOf));
        Assert.Equal("QFP-44", amplifier.Package);

        // A component of it given its case is the chip still; one given a DIP's is not.
        Assert.NotNull(IcPinouts.For(new ComponentEntry { BoardLabel = "U108", TechnicalNameOrValue = "CX20185", Package = "QFP-44" }, []));
        Assert.Null(IcPinouts.For(new ComponentEntry { BoardLabel = "U108", TechnicalNameOrValue = "CX20185", Package = "DIP-44" }, []));
    }

    // ###########################################################################################
    // THE AMIGA 500's AND 1000's CHIPS (owner request, 2026-11-18), as their sources give them: the 68000 as Motorola's
    // manual, the A1000's original Agnus as the Hardware Reference Manual, the A500's Fat Agnus in its PLCC-84 and Gary by
    // the rev 6A schematic's names, Denise and Paula as the A500 service manual, the 8520 the 6526 pin for pin, the 27C400 as
    // Atmel's datasheet. Each sampled at the pins a mistake would show at first: the supplies, the ends of each bus, and the
    // pins its sources disagreed on.
    // ###########################################################################################
    [Fact]
    public void The_Amiga_500s_and_1000s_chips_read_as_their_sources()
    {
        IcPinout cpu = IcPinouts.Named("MC68000P8")!;

        Assert.Equal("68000", cpu.Name);
        Assert.Equal("DIP-64", cpu.Package);
        Assert.Equal(["D4", "D0", "/AS", "R//W", "VCC", "CLK", "GND", "/HALT", "/IPL2", "/IPL0", "FC0", "A1", "A20", "VCC", "A23", "GND", "D15", "D5"],
            new[] { 1, 5, 6, 9, 14, 15, 16, 17, 23, 25, 28, 29, 48, 49, 52, 53, 54, 64 }.Select(cpu.NameOf));
        Assert.Equal(["Input/output", "Output", "Input/output", "Input", "Output"], new[] { 1, 9, 18, 21, 52 }.Select(cpu.DirectionOf));

        IcPinout agnus = IcPinouts.Named("8367")!;

        Assert.Equal("8361", agnus.Name);
        Assert.Equal(48, agnus.PinCount);
        Assert.Equal(["D8", "VCC", "RGA8", "RGA1", "VSS", "DRA0", "DRA8", "/HSY", "VSS", "D9"], new[] { 1, 10, 17, 24, 27, 28, 36, 40, 41, 48 }.Select(agnus.NameOf));

        IcPinout fatAgnus = IcPinouts.Named("8372A")!;

        Assert.Equal("PLCC-84", fatAgnus.Package);
        Assert.Equal(84, fatAgnus.PinCount);
        Assert.Equal(["DRD13", "DRD0", "VCC", "/RESET", "RGA8", "RGA1", "28MHz", "CCK", "/NTSC/PAL", "VSS", "DRA0", "/RAS0", "A19", "A1", "A18", "/LPEN", "DRD14"],
            new[] { 1, 14, 15, 16, 26, 33, 34, 40, 41, 42, 43, 57, 59, 60, 77, 78, 84 }.Select(fatAgnus.NameOf));
        Assert.Equal(["Output", "Input", "Ground", "Output", "Input/output"], new[] { 26, 34, 58, 80, 84 }.Select(fatAgnus.DirectionOf));
        Assert.Equal("TEST", IcPinouts.Named("8371")!.NameOf(41));
        Assert.Equal(fatAgnus.PinNames.Where((_, index) => index != 40), IcPinouts.Named("8370")!.PinNames.Where((_, index) => index != 40));

        IcPinout denise = IcPinouts.Named("8362R8")!;

        Assert.Equal(["D6", "D0", "M1H", "RGA8", "/BURST", "VCC", "R0", "G3", "/CSYNC", "/ZD", "NC", "7M", "VSS", "M0V", "D15", "D7"],
            new[] { 1, 7, 8, 10, 18, 19, 20, 31, 32, 33, 34, 35, 37, 38, 40, 48 }.Select(denise.NameOf));
        Assert.Equal(["", "Not connected"], new[] { 32, 34 }.Select(denise.DirectionOf));
        Assert.Equal(["/CDAC", "Input"], [IcPinouts.Named("8373")!.NameOf(34), IcPinouts.Named("8373")!.DirectionOf(34)]);

        IcPinout paula = IcPinouts.Named("8364R7")!;

        Assert.Equal(["D8", "VSS", "D0", "/IPL0", "/IPL2", "/INT6", "VCC", "AUDB", "AUDA", "VSSANA", "DKWE", "RXD", "D9"],
            new[] { 1, 8, 10, 13, 15, 18, 27, 30, 31, 34, 39, 41, 48 }.Select(paula.NameOf));

        IcPinout gary = IcPinouts.Named("5719")!;

        Assert.Equal(["GND", "VCC", "DKWE", "/DKWD", "VCC", "GND", "A17", "A23", "GND", "/RST", "VCC"],
            new[] { 1, 6, 8, 9, 17, 24, 33, 39, 40, 41, 48 }.Select(gary.NameOf));
        Assert.Equal(["Input", "Input", "", "Output", "Output"], new[] { 8, 9, 40, 43, 47 }.Select(gary.DirectionOf));

        Assert.Same(IcPinouts.Named("6526"), IcPinouts.Named("8520A-1"));

        IcPinout eprom = IcPinouts.Named("M27C400")!;

        Assert.Equal(["A17", "A0", "/CE", "GND", "O0", "VCC", "O15/A-1", "GND", "/BYTE/VPP", "A16", "A8"], new[] { 1, 9, 10, 11, 13, 21, 29, 30, 31, 32, 40 }.Select(eprom.NameOf));
    }

    // ###########################################################################################
    // The Amiga's other parts (2026-11-18), as their makers' datasheets: the RAMs, the clock, the video encoder, the serial
    // port's driver and receiver, the audio filter's op-amp, a comparator, and the A1000's logic CRT draws no insides of.
    // ###########################################################################################
    [Fact]
    public void The_Amigas_other_parts_read_as_their_datasheets()
    {
        Assert.Equal(["A8", "DIN", "VCC", "DOUT", "/CAS", "GND"], new[] { 1, 2, 8, 14, 15, 16 }.Select(IcPinouts.Named("uPD41256C")!.NameOf));
        Assert.Equal(["DQ1", "NC", "A0", "VCC", "A8", "/OE", "DQ4", "VSS"], new[] { 1, 5, 6, 10, 15, 16, 19, 20 }.Select(IcPinouts.Named("MB81C4256")!.NameOf));
        Assert.Equal("Not connected", IcPinouts.Named("TMS44C256")!.DirectionOf(5));

        IcPinout clock = IcPinouts.Named("MSM6242B")!;

        Assert.Equal(["STD.P", "/CS0", "/RD", "GND", "D0", "CS1", "XT", "/XT", "VDD"], new[] { 1, 2, 8, 9, 14, 15, 16, 17, 18 }.Select(clock.NameOf));
        Assert.Equal(["Output", "Input/output", "Other", "Other"], new[] { 1, 11, 16, 17 }.Select(clock.DirectionOf));

        IcPinout encoder = IcPinouts.Named("MC1377P")!;

        Assert.Equal(["TR", "SYNC", "-YOUT", "CVOUT", "VCC", "GND", "VB", "NTSC/PAL"], new[] { 1, 2, 6, 9, 14, 15, 16, 20 }.Select(encoder.NameOf));
        Assert.Equal(["Other", "Other", "Ground", "Output"], new[] { 1, 7, 15, 16 }.Select(encoder.DirectionOf));

        Assert.Equal(["VCC-", "1Y", "GND", "VCC+"], new[] { 1, 3, 7, 14 }.Select(IcPinouts.Named("MC1488P")!.NameOf));
        Assert.Equal(["Other", "Output", "Power"], new[] { 2, 3, 14 }.Select(IcPinouts.Named("MC1489L")!.DirectionOf));
        Assert.Same(IcPinouts.Named("LF347N"), IcPinouts.Named("TL084CN"));
        Assert.Equal(["VCC+", "VCC-"], new[] { 4, 11 }.Select(IcPinouts.Named("TL084CN")!.NameOf));
        Assert.Same(IcPinouts.Named("LM2901N"), IcPinouts.Named("LM339AN"));
        Assert.Equal(["1OUT", "VCC", "GND", "3OUT"], new[] { 1, 3, 12, 14 }.Select(IcPinouts.Named("LM2901N")!.NameOf));

        Assert.Equal(["/OE", "1Q", "GND", "CLK", "8Q", "VCC"], new[] { 1, 2, 10, 11, 19, 20 }.Select(IcPinouts.Named("SN74F374N")!.NameOf));
        Assert.Equal(["S", "Qa", "GND", "CP", "Qd"], new[] { 1, 2, 8, 9, 15 }.Select(IcPinouts.Named("MC74F399N")!.NameOf));
        Assert.Equal(["1A", "1CLR", "GND", "2A"], new[] { 1, 2, 7, 13 }.Select(IcPinouts.Named("SN74LS393N")!.NameOf));
    }

    // ###########################################################################################
    // *** THE 51'S PINS DEPEND ON ITS FAMILY *** (TI's SDLS113): the 7451 and 74S51 have two AND-OR-INVERT gates of 2 x 2
    // inputs, pins 11 and 12 unused; the 74LS51 one of 3 x 3 on pins 9 to 13 - so the family as written decides, and a family no
    // source here gives (a 74HC51) is none.
    // ###########################################################################################
    [Theory]
    [InlineData("SN7451N", "7451")]
    [InlineData("74S51", "7451")]
    [InlineData("DM74S51N", "7451")]
    [InlineData("SN74LS51N", "74LS51")]
    [InlineData("74LS51", "74LS51")]
    [InlineData("74HC51", null)]
    public void The_51_is_the_pin_out_of_its_family(string written, string? chip)
    {
        Assert.Equal(chip, IcPinouts.Named(written)?.Name);
    }

    [Fact]
    public void The_7451_leaves_pins_11_and_12_unused_and_the_74LS51_takes_inputs_there()
    {
        Assert.Equal(["Not connected", "Not connected"], new[] { 11, 12 }.Select(IcPinouts.Named("74S51")!.DirectionOf));
        Assert.Equal(["1F", "1B"], new[] { 11, 12 }.Select(IcPinouts.Named("74LS51")!.NameOf));
        Assert.Equal(["Input", "Input"], new[] { 11, 12 }.Select(IcPinouts.Named("74LS51")!.DirectionOf));
    }

    // The A500+'s 8375 moves Agnus's pins 35 to 57, and the A3000's 8372B is checked against no source: neither is a Fat Agnus
    // here.
    [Theory]
    [InlineData("8375")]
    [InlineData("8372B")]
    [InlineData("8360")]
    public void An_Agnus_of_another_pin_out_is_none(string written)
    {
        Assert.NotEqual("8372", IcPinouts.Named(written)?.Name);
        Assert.NotEqual("8370", IcPinouts.Named(written)?.Name);
    }

    // ###########################################################################################
    // *** THE 64H156 IS KNOWN BY ITS PART NUMBER ALONE *** (2026-11-18): the 251828-01 and -03 have 40 pins, the -02 42 (the
    // 1571 service manual), so a 40-pin one is found by its revision and "64H156" alone, or the -02, is none - a drawing of
    // the wrong one numbers every pin past 20 wrongly. The 250477's is a -01, of 40 pins (the owner counted them).
    // ###########################################################################################
    [Theory]
    [InlineData("251828-01")]
    [InlineData("251828-03")]
    [InlineData("64H156-01")]
    [InlineData("64h156-03")]
    public void A_40_pin_64H156_is_found_by_its_revision(string written)
    {
        IcPinout gateArray = IcPinouts.Named(written)!;

        Assert.Equal(40, gateArray.PinCount);
        Assert.Equal(["TEST", "YB0", "VSS", "SYNC", "OSC", "ATN", "Y0", "XWR", "VDD", "BYTE"], new[] { 1, 2, 10, 17, 21, 24, 25, 29, 30, 40 }.Select(gateArray.NameOf));
        Assert.Equal(["Input", "Input/output", "Ground", "Output", "Input", "Output", "Output", "Output", "Power", "Output"], new[] { 1, 2, 10, 17, 21, 24, 25, 29, 30, 40 }.Select(gateArray.DirectionOf));
    }

    [Theory]
    [InlineData("64H156")]
    [InlineData("251828-02")]
    [InlineData("64H156-02")]
    [InlineData("251828")]
    public void A_64H156_of_no_40_pin_revision_is_no_chip_CRT_holds(string written)
    {
        Assert.Null(IcPinouts.Named(written));
        Assert.Null(LogicChips.Named(written));
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
    [InlineData("8568R9", "8568")]
    [InlineData("5710", "5710")]
    [InlineData("40008/40010", "40010")]
    [InlineData("40010", "40010")]
    [InlineData("CX20185", "CX20185")]
    [InlineData("41464", "41464")]
    [InlineData("4464", "41464")]
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
    [InlineData("Gate Array")]
    // The CPC's ROM is named by its part number, a 23256 (A_component_whose_category_says_nothing_against_it_is_its_chip).
    [InlineData("40022")]
    [InlineData("")]
    [InlineData(null)]
    public void A_name_that_is_not_one_of_them_is_no_chip(string? written)
    {
        Assert.Null(IcPinouts.Named(written));
    }

    // The CPC's Amstrad numbers start like CMOS 4000-series ones and are none: the 40022 ROM is no 4002, the Gate Array no 4000
    // or 4001.
    [Theory]
    [InlineData("40022")]
    [InlineData("40008/40010")]
    [InlineData("40010")]
    public void An_Amstrad_number_is_no_CMOS_logic_chip(string written)
    {
        Assert.Null(LogicChips.Named(written));
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

        // The CPC 664's ROM: "40022" names nothing, its part number the 23256.
        Assert.Equal("23256", IcPinouts.For(new ComponentEntry { BoardLabel = "IC103", TechnicalNameOrValue = "40022", PartNumber = "23256", Category = "IC" }, [])?.Name);
    }
}
