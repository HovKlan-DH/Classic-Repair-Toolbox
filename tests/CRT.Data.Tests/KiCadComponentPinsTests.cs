using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadComponentPins - a component's pins from the board's KiCad data (CRT 4.0 step 2: "Fill from
// KiCad adds each pin's net and its name from the symbol"): the pads give the numbers and the nets,
// the symbol names a pin before its pad does, names and directions in CRT's words - and "Fill"
// never changes what was typed.
// ###########################################################################################
public sealed class KiCadComponentPinsTests
{
    // U8's footprint: three pads, as KiCad 6 and later write them (with pin names and types).
    private static readonly KiCadPadInfo[] Pads =
    [
        new("U8", "1", "~{RESET}", "input", "/CPU/~{RESET}"),
        new("U8", "2", "GND", "power_in", "GND"),
        new("U8", "10", "D0", "bidirectional", "/Data bus/D0"),
        new("u8 ", "3", "R/~{W}", "output+no_connect", "unconnected-(U8-R{slash}~{W}-Pad3)"),
        new("U9", "1", "VCC", "power_in", "+5V"),
    ];

    [Fact]
    public void The_pads_give_each_pin_its_number_name_direction_and_net_in_pin_order()
    {
        List<KiCadPin> pins = KiCadComponentPins.Of(Pads, [], " U8 ");

        Assert.Equal(
            [
                new KiCadPin("1", "~RESET", "Input", "~RESET"),
                new KiCadPin("2", "GND", "Ground", "GND"),
                new KiCadPin("3", "R/~W", "Output", ""),
                new KiCadPin("10", "D0", "Input/output", "D0"),
            ],
            pins);
    }

    // A schematic's symbol names the pin before its pad does - and adds a pin no pad has.
    [Fact]
    public void The_symbol_names_a_pin_before_its_pad_and_adds_pins_no_pad_has()
    {
        List<KiCadPin> pins = KiCadComponentPins.Of(
            Pads,
            [
                new KiCadSymbolPinInfo("U8", "10", "DB0", "tri_state"),
                new KiCadSymbolPinInfo("U8", "40", "VCC", "power_in"),
                new KiCadSymbolPinInfo("U8", "2", "~", "passive"),
            ],
            "U8");

        Assert.Equal(new KiCadPin("10", "DB0", "Output", "D0"), pins.Single(pin => pin.Pin == "10"));
        Assert.Equal(new KiCadPin("40", "VCC", "Power", ""), pins.Single(pin => pin.Pin == "40"));

        // A symbol pin with no name and a passive type leaves the pad's.
        Assert.Equal(new KiCadPin("2", "GND", "Ground", "GND"), pins.Single(pin => pin.Pin == "2"));
    }

    // ###########################################################################################
    // A project that appends each pad's own number to its name ("A2_6" on pad 6) - every named pad of
    // the footprint does - has it taken off; a real name with an underscore keeps it.
    // ###########################################################################################
    [Fact]
    public void A_pad_number_appended_to_every_name_is_taken_off_but_a_real_underscore_stays()
    {
        KiCadPadInfo[] appended =
        [
            new("U18", "6", "A2_6", "input", ""),
            new("U18", "12", "GND_12", "power_in", ""),
            new("U18", "13", "~_13", "passive", ""),
        ];

        Assert.Equal(["A2", "GND", ""], KiCadComponentPins.Of(appended, [], "U18").Select(pin => pin.Name));

        KiCadPadInfo[] real =
        [
            new("U1", "1", "AUDIO_IN", "input", ""),
            new("U1", "2", "EXT_2", "input", ""),
        ];

        Assert.Equal(["AUDIO_IN", "EXT_2"], KiCadComponentPins.Of(real, [], "U1").Select(pin => pin.Name));
    }

    [Fact]
    public void A_board_without_the_component_or_a_blank_label_has_no_pins()
    {
        Assert.Empty(KiCadComponentPins.Of(Pads, [], "U99"));
        Assert.Empty(KiCadComponentPins.Of(Pads, [], " "));
        Assert.Empty(KiCadComponentPins.Of([new KiCadPadInfo("U8", " ", "", "", "")], [], "U8"));
    }

    [Theory]
    [InlineData("input", "A0", "Input")]
    [InlineData("open_collector", "~{IRQ}", "Output")]
    [InlineData("open_emitter", "X", "Output")]
    [InlineData("power_in", "VSS", "Ground")]
    [InlineData("power_in", "AGND", "Ground")]
    [InlineData("power_in", "0V", "Ground")]
    [InlineData("power_in", "VDD", "Power")]
    [InlineData("power_out", "+5V", "Power")]
    [InlineData("no_connect", "NC", "Not connected")]
    [InlineData("passive", "1", "")]
    [InlineData("unspecified", "X", "")]
    [InlineData("free", "X", "")]
    [InlineData("", "X", "")]
    [InlineData("INPUT+no_connect", "X", "Input")]
    // A library typing every pin "input" (the live 250469's): a supply pin is told by its name.
    [InlineData("input", "GND", "Ground")]
    [InlineData("input", "VDD", "Power")]
    [InlineData("passive", "+12V", "Power")]
    [InlineData("input", "9V", "Power")]
    [InlineData("input", "VPP", "Input")]
    [InlineData("no_connect", "GND", "Not connected")]
    public void A_KiCad_pin_type_is_a_direction_in_CRTs_words(string type, string name, string direction)
    {
        Assert.Equal(direction, KiCadComponentPins.Direction(type, name));
        Assert.True(direction.Length == 0 || ComponentVocabulary.PinDirections.Contains(direction));
    }

    [Theory]
    [InlineData("~{RESET}", "~RESET")]
    [InlineData("R/~{W}", "R/~W")]
    [InlineData("~", "")]
    [InlineData(" A0 ", "A0")]
    [InlineData(null, "")]
    public void A_pin_name_is_spelt_as_CRT_spells_it(string? kiCad, string crt) =>
        Assert.Equal(crt, KiCadComponentPins.PinName(kiCad));

    [Theory]
    [InlineData("/I{slash}O/Serial Bus/CNT", "CNT")]
    [InlineData("/Bus/R{slash}~{W}", "R/~W")]
    [InlineData("Net-(C12-Pad1)", "Net-(C12-Pad1)")]
    [InlineData("unconnected-(U8-Pad3)", "")]
    [InlineData("", "")]
    public void A_net_is_shown_by_its_last_part(string kiCad, string shown) =>
        Assert.Equal(shown, KiCadComponentPins.NetName(kiCad));

    // ###########################################################################################
    // "Fill from KiCad": the pins missing added in the region, the empty names and directions filled
    // in - a name or direction typed is never changed - all in pin order, with what changed counted.
    // ###########################################################################################
    [Fact]
    public void Fill_adds_what_is_missing_and_never_changes_what_was_typed()
    {
        List<KiCadPin> kiCad = KiCadComponentPins.Of(Pads, [], "U8");

        List<ComponentPinEntry> current =
        [
            new() { Region = "PAL", Pin = "10", Name = "", Direction = "" },
            new() { Region = "PAL", Pin = "1", Name = "RES", Direction = "" },
            new() { Region = "PAL", Pin = "2", Name = "Ground", Direction = "Other" },
        ];

        List<ComponentPinEntry> filled = KiCadComponentPins.Fill(current, kiCad, "PAL", out int added, out int completed);

        Assert.Equal(
            ["1:RES:Input", "2:Ground:Other", "3:R/~W:Output", "10:D0:Input/output"],
            filled.Select(pin => $"{pin.Pin}:{pin.Name}:{pin.Direction}"));
        Assert.All(filled, pin => Assert.Equal("PAL", pin.Region));
        Assert.Equal((1, 2), (added, completed));

        // Filled again, nothing more to do.
        KiCadComponentPins.Fill(filled, kiCad, "PAL", out added, out completed);
        Assert.Equal((0, 0), (added, completed));
    }

    [Fact]
    public void Each_pins_net_by_its_number_leaves_out_pins_on_no_net()
    {
        Dictionary<string, string> nets = KiCadComponentPins.NetsByPin(KiCadComponentPins.Of(Pads, [], "U8"));

        Assert.Equal(["1", "2", "10"], nets.Keys.Order(ComponentPaneEdits.PinOrder));
        Assert.Equal("D0", nets["10"]);
    }
}
