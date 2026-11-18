using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DatasheetPins - where a pin's direction comes from (owner decision, 2026-11-17): the datasheet of a chip CRT holds it of,
// before KiCad's type - but what the board must keep of it.
// ###########################################################################################
public sealed class DatasheetPinsTests
{
    [Theory]
    // The datasheet over whatever was said - KiCad's input on an output, nothing, the conversion's "Unknown".
    [InlineData("Input", "Output", "Output")]
    [InlineData("", "Input", "Input")]
    [InlineData("Unknown", "Input/output", "Input/output")]
    [InlineData("Power", "Ground", "Ground")]
    // A pin the board says is not connected stays so: an unused output of that board.
    [InlineData("Not connected", "Output", "Not connected")]
    // A pin the datasheet leaves unconnected keeps what the board says of it - a ROM's pin 1 wired to +5V - or is unconnected.
    [InlineData("Power", "Not connected", "Power")]
    [InlineData("Unknown", "Not connected", "Not connected")]
    // A pin the datasheet does not settle keeps the board's word.
    [InlineData("Input", "", "Input")]
    [InlineData("Unknown", "", "Unknown")]
    public void The_datasheet_settles_a_pin_but_what_the_board_must_keep(string board, string datasheet, string settled)
    {
        Assert.Equal(settled, DatasheetPins.Settled(board, datasheet));
    }

    // A logic chip's directions are LogicChips', another IC's IcPinouts', and a chip CRT holds neither of has none.
    [Fact]
    public void A_logic_chip_and_another_IC_have_their_datasheets_and_an_unknown_chip_none()
    {
        Func<int, string>? gates = DatasheetPins.Of(new ComponentEntry { TechnicalNameOrValue = "74LS08" }, []);
        Func<int, string>? cpu = DatasheetPins.Of(new ComponentEntry { TechnicalNameOrValue = "6510" }, []);

        Assert.Equal(["Input", "Output", "Ground", "Power", ""], new[] { 1, 3, 7, 14, 15 }.Select(gates!));
        Assert.Equal(["Output", "Input/output"], new[] { 7, 30 }.Select(cpu!));
        Assert.Null(DatasheetPins.Of(new ComponentEntry { TechnicalNameOrValue = "XC9572" }, []));
    }

    // "Fill pins from datasheet" (owner request, 2026-11-17) fills in the datasheet's every pin - its number, name and
    // direction - of a logic chip and of another IC alike; a chip CRT holds no datasheet of has none.
    [Fact]
    public void A_datasheets_pinout_is_every_pin_with_its_name_and_direction()
    {
        DatasheetPinout? cpu = DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "6510" }, []);
        DatasheetPinout? gates = DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "74LS08" }, []);

        Assert.Equal("6510", cpu!.Chip);
        Assert.Equal(40, cpu.Pins.Count);
        Assert.Equal(new KiCadPin("7", "A0", "Output", string.Empty), cpu.Pins[6]);
        Assert.Equal(new KiCadPin("21", "GND", "Ground", string.Empty), cpu.Pins[20]);

        Assert.Equal("7408", gates!.Chip);
        Assert.Equal(["1|1A|Input", "3|1Y|Output", "7|GND|Ground", "14|VCC|Power"],
            new[] { 0, 2, 6, 13 }.Select(index => gates.Pins[index]).Select(pin => $"{pin.Pin}|{pin.Name}|{pin.Direction}"));

        // A name the pin-out writes with "_" is shown with its space.
        Assert.Equal("PHI COLOR", DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "6569" }, [])!.Pins[20].Name);

        Assert.Null(DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "XC9572" }, []));
    }

    // The case each pin-out is of, which the conversion writes where a row has none (owner request, 2026-11-17): a DIP of its
    // pins for a logic chip and another IC alike, the CX20185's QFP-44.
    [Fact]
    public void A_datasheets_pinout_says_its_case()
    {
        DatasheetPinout? Of(string chip) => DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = chip }, []);

        Assert.Equal("DIP-40", Of("6510")!.Package);
        Assert.Equal("DIP-14", Of("74LS08")!.Package);
        Assert.Equal("DIP-18", Of("4464")!.Package);
        Assert.Equal("QFP-44", Of("CX20185")!.Package);

        // A logic chip in a SOIC is that chip, its pin-out its DIP's (owner request, 2026-11-18).
        Assert.Equal("7408", DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "74HC08", Package = "SOIC-14" }, [])!.Chip);
    }

    // "Fill pins from KiCad" and the import fill in KiCad's pins with each direction settled by the datasheet; a pin named,
    // not numbered, is left as KiCad has it; no datasheet leaves the list as it is.
    [Fact]
    public void KiCads_pins_are_settled_by_the_datasheet()
    {
        IReadOnlyList<KiCadPin> kiCad = [new KiCadPin("7", "A0", "Input", "A0"), new KiCadPin("24", "P5", "Not connected", ""), new KiCadPin("TP", "X", "Input", "")];

        Assert.Equal(
            ["7|Output", "24|Not connected", "TP|Input"],
            DatasheetPins.Over(kiCad, DatasheetPins.Of(new ComponentEntry { TechnicalNameOrValue = "6510" }, [])).Select(pin => $"{pin.Pin}|{pin.Direction}"));

        Assert.Same(kiCad, DatasheetPins.Over(kiCad, null));

        // Also with a DatasheetPinout's logic chip there, so a caller drawing it looks it up once.
        Assert.Same(LogicChips.Named("74LS08"), DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "74LS08" }, [])!.Logic);
        Assert.Null(DatasheetPins.PinoutOf(new ComponentEntry { TechnicalNameOrValue = "6510" }, [])!.Logic);
    }

    // ###########################################################################################
    // OverKiCad, "Fill pins from KiCad"'s and the import's one rule (code review, 2026-11-17): the datasheet is judged on the
    // pins AS KiCad WOULD FILL THEM - a symbol naming the 6510's pin 9 (A2, an output) "VCC" says something against it, and
    // KiCad's pins go in as they are; one that does not takes the datasheet's directions. No component, KiCad's pins.
    // ###########################################################################################
    [Fact]
    public void KiCads_pins_are_settled_by_a_datasheet_judged_on_the_pins_KiCad_fills_in()
    {
        var cpu = new ComponentEntry { BoardLabel = "U8", TechnicalNameOrValue = "6510", Package = "DIP-40" };
        List<ComponentPinEntry> held = [new() { BoardLabel = "U8", Pin = "1", Name = "PHI0", Direction = "Input" }];

        IReadOnlyList<KiCadPin> against = [new KiCadPin("7", "A0", "Input", ""), new KiCadPin("9", "VCC", "Power", "")];
        IReadOnlyList<KiCadPin> agreeing = [new KiCadPin("7", "A0", "Input", ""), new KiCadPin("9", "A2", "Input", "")];

        Assert.Equal(["7|Input", "9|Power"], DatasheetPins.OverKiCad(cpu, held, against).Select(pin => $"{pin.Pin}|{pin.Direction}"));
        Assert.Equal(["7|Output", "9|Output"], DatasheetPins.OverKiCad(cpu, held, agreeing).Select(pin => $"{pin.Pin}|{pin.Direction}"));
        Assert.Same(against, DatasheetPins.OverKiCad(null, held, against));
    }
}
