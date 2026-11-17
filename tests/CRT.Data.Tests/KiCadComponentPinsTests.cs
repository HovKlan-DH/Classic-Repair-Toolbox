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
                new KiCadPin("1", "/RESET", "Input", "/RESET"),
                new KiCadPin("2", "GND", "Ground", "GND"),
                new KiCadPin("3", "R//W", "Output", ""),
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

    // ###########################################################################################
    // A SYMBOL TYPING EVERY SIGNAL PIN "input" WAS NEVER TYPED (owner decision, 2026-11-17): no chip has no outputs, so its
    // inputs say nothing and are left out - the live 250407's 6510 and ROMs, the 250469's every chip. A supply, a ground and
    // an unconnected pin keep theirs; one output anywhere on the symbol and every type stands; and so do inputs that are
    // fewer than half its pins, the rest left unspecified - a symbol that typed only its ~RESET typed that.
    // ###########################################################################################
    [Fact]
    public void A_symbol_typing_every_signal_pin_input_gives_those_pins_no_direction()
    {
        KiCadPadInfo[] untyped =
        [
            new("U7", "1", "PHI0", "input", ""),
            new("U7", "7", "A0", "input", ""),
            new("U7", "8", "A1", "input", ""),
            new("U7", "21", "GND", "input", ""),
            new("U7", "35", "", "no_connect", ""),
            new("U7", "40", "VCC", "power_in", ""),
        ];

        Assert.Equal(
            ["1|", "7|", "8|", "21|Ground", "35|Not connected", "40|Power"],
            KiCadComponentPins.Of(untyped, [], "U7").Select(pin => $"{pin.Pin}|{pin.Direction}"));

        KiCadPadInfo[] typed = [.. untyped.Take(1), new("U7", "7", "A0", "tri_state", ""), .. untyped.Skip(2)];

        Assert.Equal(
            ["1|Input", "7|Output", "8|Input", "21|Ground", "35|Not connected", "40|Power"],
            KiCadComponentPins.Of(typed, [], "U7").Select(pin => $"{pin.Pin}|{pin.Direction}"));

        KiCadPadInfo[] partly = [.. untyped.Take(2), new("U7", "8", "A1", "unspecified", ""), .. untyped.Skip(3)];

        Assert.Equal(
            ["1|Input", "7|Input", "8|", "21|Ground", "35|Not connected", "40|Power"],
            KiCadComponentPins.Of(partly, [], "U7").Select(pin => $"{pin.Pin}|{pin.Direction}"));
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

    // An active-low pin is written with "/" before it, as the published boards write theirs (owner decision, 2026-11-17).
    [Theory]
    [InlineData("~{RESET}", "/RESET")]
    [InlineData("R/~{W}", "R//W")]
    [InlineData("~", "")]
    [InlineData(" A0 ", "A0")]
    [InlineData(null, "")]
    public void A_pin_name_is_spelt_as_CRT_spells_it(string? kiCad, string crt) =>
        Assert.Equal(crt, KiCadComponentPins.PinName(kiCad));

    [Theory]
    [InlineData("/I{slash}O/Serial Bus/CNT", "CNT")]
    [InlineData("/Bus/R{slash}~{W}", "R//W")]
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
            ["1:RES:Input", "2:Ground:Other", "3:R//W:Output", "10:D0:Input/output"],
            filled.Select(pin => $"{pin.Pin}:{pin.Name}:{pin.Direction}"));
        Assert.All(filled, pin => Assert.Equal("PAL", pin.Region));
        Assert.Equal((1, 2), (added, completed));

        // Filled again, nothing more to do.
        KiCadComponentPins.Fill(filled, kiCad, "PAL", out added, out completed);
        Assert.Equal((0, 0), (added, completed));
    }

    // WouldFill says whether Fill would add or fill in anything - "Fill pins from datasheet" is offered only then - and
    // changes nothing it is handed.
    [Fact]
    public void Would_fill_says_whether_Fill_would_add_or_fill_in_anything()
    {
        List<KiCadPin> source = [new KiCadPin("1", "PHI0", "Input", string.Empty), new KiCadPin("2", "RDY", "Input", string.Empty)];
        List<ComponentPinEntry> none = [];
        List<ComponentPinEntry> nameless = [new() { Pin = "1", Name = "", Direction = "Input" }, new() { Pin = "2", Name = "RDY", Direction = "Input" }];
        List<ComponentPinEntry> whole = [new() { Pin = "1", Name = "CLK", Direction = "Output" }, new() { Pin = "2", Name = "RDY", Direction = "Input" }];

        Assert.True(KiCadComponentPins.WouldFill(none, source));
        Assert.True(KiCadComponentPins.WouldFill(nameless, source));
        Assert.False(KiCadComponentPins.WouldFill(whole, source));
        Assert.False(KiCadComponentPins.WouldFill(whole, []));

        Assert.Empty(none);
        Assert.Equal(string.Empty, nameless[0].Name);
    }

    // "Unknown" says no direction, as DatasheetPins.Settled reads it: Fill fills it in (code review, 2026-11-17 - a chip the
    // conversion left "Unknown" could never take its datasheet's directions in the pane). A direction given stays.
    [Fact]
    public void Fill_fills_in_a_direction_that_says_Unknown()
    {
        List<ComponentPinEntry> current =
        [
            new() { Pin = "1", Name = "PHI0", Direction = "Unknown" },
            new() { Pin = "2", Name = "RDY", Direction = "unknown " },
            new() { Pin = "3", Name = "IRQ", Direction = "Output" },
        ];
        List<KiCadPin> source = [new KiCadPin("1", "PHI0", "Input", ""), new KiCadPin("2", "RDY", "Input", ""), new KiCadPin("3", "IRQ", "Input", "")];

        List<ComponentPinEntry> filled = KiCadComponentPins.Fill(current, source, string.Empty, out int added, out int completed);

        Assert.Equal(["1|Input", "2|Input", "3|Output"], filled.Select(pin => $"{pin.Pin}|{pin.Direction}"));
        Assert.Equal((0, 2), (added, completed));
        Assert.True(KiCadComponentPins.WouldFill(current, source));
    }

    [Fact]
    public void Each_pins_net_by_its_number_leaves_out_pins_on_no_net()
    {
        Dictionary<string, string> nets = KiCadComponentPins.NetsByPin(KiCadComponentPins.Of(Pads, [], "U8"));

        Assert.Equal(["1", "2", "10"], nets.Keys.Order(ComponentPaneEdits.PinOrder));
        Assert.Equal("D0", nets["10"]);
    }

    // ###########################################################################################
    // A new oscilloscope baseline's signal name comes with its pin (owner request, 2026-11-02): the
    // pinout's, as somebody may have worked on it, else the board's KiCad data's.
    // ###########################################################################################
    [Fact]
    public void A_pins_name_is_the_pinouts_else_the_KiCad_datas()
    {
        ComponentPinEntry[] pinout =
        [
            new() { Pin = "1", Name = "My name" },
            new() { Pin = "2", Name = " " },
        ];

        KiCadPin[] kiCad =
        [
            new("1", "RESET", "Input", "/RESET"),
            new("2", "GND", "Ground", "GND"),
            new("3", "D0", "Input/output", "/D0"),
        ];

        Assert.Equal("My name", KiCadComponentPins.NameOfPin(pinout, kiCad, " 1 "));
        Assert.Equal("GND", KiCadComponentPins.NameOfPin(pinout, kiCad, "2"));
        Assert.Equal("D0", KiCadComponentPins.NameOfPin(pinout, kiCad, "3 [64K]"));
        Assert.Equal(string.Empty, KiCadComponentPins.NameOfPin(pinout, kiCad, "4"));
        Assert.Equal(string.Empty, KiCadComponentPins.NameOfPin(pinout, kiCad, ""));
    }
}
