using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The pinouts the data tree's conversion makes of a board's scope baselines and KiCad data (owner
// decision, 2026-11-06 - a temporary tool on the "Pinout" page until then, BaselinePinMigration, whose
// release guard went with the tool). DataTreeConversionTests holds the conversion running it.
// ###########################################################################################
public sealed class ConversionPinoutsTests
{
    // ###########################################################################################
    // A USER'S EXCEL-ERA DRAFT TAKES ITS PUBLISHED BOARD'S PINOUTS (DraftConversion, 2026-11-06): a
    // component with no pins gets the published board's pins for its label, a row with no package the
    // published row's of the same region - while pins the draft has, a component the draft does not
    // have and a package the draft names stay as they are.
    // ###########################################################################################
    [Fact]
    public void A_draft_takes_the_published_pinouts_only_where_it_has_none()
    {
        BoardData published = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Region = "PAL", Category = "IC", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U1", Region = "NTSC", Category = "IC", Package = "DIP-28" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", Package = "DIP-14" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", Package = "DIP-8" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "VSS", Direction = "Ground" },
                new ComponentPinEntry { BoardLabel = "U2", Pin = "1", Name = "1A", Direction = "Input" },
                new ComponentPinEntry { BoardLabel = "U9", Pin = "1", Name = "X", Direction = "Unknown" },
            ],
        };

        BoardData draft = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "u1", Region = "PAL", Category = "IC" },
                new ComponentEntry { BoardLabel = "U1", Region = "NTSC", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC", Package = "SOIC-14" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "U2", Pin = "1", Name = "Typed" }],
        };

        BoardData carried = ConversionPinouts.FromPublished(draft, published);

        Assert.Equal(["u1|PAL|DIP-40", "U1|NTSC|DIP-28", "U2||SOIC-14"], carried.Components.Select(row => $"{row.BoardLabel}|{row.Region}|{row.Package}"));
        Assert.Equal(["U2|1|Typed", "U1|1|VSS"], carried.ComponentPins.Select(row => $"{row.BoardLabel}|{row.Pin}|{row.Name}"));

        Assert.Same(draft, ConversionPinouts.FromPublished(draft, null));
        Assert.Same(carried, ConversionPinouts.FromPublished(carried, published));
    }

    private static ComponentImageEntry Baseline(string label, string region, string pin, string name) =>
        new() { BoardLabel = label, Region = region, Pin = pin, Name = name, File = $"Scope baseline/{label}_{pin}.png", TimeDiv = "1ms", VoltsDiv = "1V" };

    private static BoardData Board(string category, IEnumerable<ComponentImageEntry> baselines, params ComponentPinEntry[] pins) => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U1", Category = category }, new ComponentEntry { BoardLabel = "U2", Category = "IC" }],
        ComponentImages = [.. baselines],
        ComponentPins = [.. pins],
    };

    // Every pin of a chip, as both regions' baselines say them.
    private static IEnumerable<ComponentImageEntry> BothRegions(int pins) =>
        Enumerable.Range(1, pins).SelectMany(pin => new[]
        {
            Baseline("U1", "PAL", pin.ToString(), $"P{pin}"),
            Baseline("U1", "NTSC", pin.ToString(), $"P{pin}"),
        });

    // ###########################################################################################
    // THE WHOLE BOARD TAKES ITS KiCad DATA TOO (owner request, 2026-11-03: "it only fills in details on
    // ICs. I do think it should fill in details for all components having KiCad data, and ideally guess
    // their 'package'"): a resistor with no baseline gets its pins and its package, what is typed stays, a
    // component the board does not have is left out, and a second run has nothing to do.
    // ###########################################################################################
    private static KiCadPinoutSource KiCadWith(params (string Label, string Package)[] components) => new(
        components.Select(component => new KiCadBoardComponent(component.Label, string.Empty, "Misc", component.Package, string.Empty)).ToList(),
        _ => [new KiCadPin("1", "A", "Passive", "NET1"), new KiCadPin("2", "B", "Passive", "NET2")]);

    [Fact]
    public void The_whole_board_fills_in_every_component_the_KiCad_data_has_and_guesses_its_package()
    {
        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "R1", Category = "Resistor" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor", Package = "Radial" },
            ],
            ComponentPins = [new ComponentPinEntry { BoardLabel = "C1", Pin = "1", Name = "Typed" }],
        };

        BoardData made = ConversionPinouts.ApplyToBoard(board, out int components, out _, KiCadWith(("R1", "Axial"), ("C1", "Disc"), ("X9", "HC-49")));

        Assert.Equal(2, components);

        Assert.Equal("Axial", made.Components.Single(row => row.BoardLabel == "R1").Package);
        Assert.Equal(["1|A", "2|B"], made.ComponentPins.Where(pin => pin.BoardLabel == "R1").Select(pin => $"{pin.Pin}|{pin.Name}"));

        // What it had stays - the case chosen and the name typed - and what it lacked is added.
        Assert.Equal("Radial", made.Components.Single(row => row.BoardLabel == "C1").Package);
        Assert.Equal(["1|Typed", "2|B"], made.ComponentPins.Where(pin => pin.BoardLabel == "C1").Select(pin => $"{pin.Pin}|{pin.Name}"));

        Assert.DoesNotContain(made.Components, row => row.BoardLabel == "X9");

        Assert.Same(made, ConversionPinouts.ApplyToBoard(made, out int again, out _, KiCadWith(("R1", "Axial"), ("C1", "Disc"), ("X9", "HC-49"))));
        Assert.Equal(0, again);
    }

    // ###########################################################################################
    // WHERE CRT HOLDS NO DATASHEET, KiCad's PIN TYPE GOES (DatasheetPins' order): it replaces the "Unknown" a baseline's pin
    // got - the conversion's own - and fills a pin no baseline made; where KiCad's type says nothing, a baseline's pin stays
    // "Unknown". (A symbol typing every signal pin "input" says nothing either - KiCadComponentPinsTests.)
    // ###########################################################################################
    [Fact]
    public void KiCads_pin_type_replaces_the_unknown_a_baselines_pin_got_where_CRT_holds_no_datasheet()
    {
        BoardData board = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U9", Category = "IC", TechnicalNameOrValue = "XC9572" }],
            ComponentImages = [Baseline("U9", string.Empty, "1", "CLK"), Baseline("U9", string.Empty, "2", "RDY"), Baseline("U9", string.Empty, "30", "D0")],
        };

        var kiCad = new KiCadPinoutSource(
            [new KiCadBoardComponent("U9", "XC9572", "IC", "DIP-40", string.Empty)],
            _ => [new KiCadPin("1", "CLK", "Input", "CLK"), new KiCadPin("2", "RDY", "", "RDY"), new KiCadPin("7", "A0", "Output", "A0"), new KiCadPin("30", "D0", "Input/output", "D0")]);

        BoardData made = ConversionPinouts.ApplyToBoard(board, out _, out _, kiCad);

        Assert.Equal(
            ["1|CLK|Input", "2|RDY|Unknown", "7|A0|Output", "30|D0|Input/output"],
            made.ComponentPins.Where(pin => pin.BoardLabel == "U9").Select(pin => $"{pin.Pin}|{pin.Name}|{pin.Direction}"));
    }

    // ###########################################################################################
    // THE DATASHEET COMES BEFORE KiCad (owner decision, 2026-11-17: "OK to what you suggest for 1-4" - DatasheetPins): the
    // 250407's KiCad project types its 556's DISCHARGE and its 6510's address lines as inputs; the datasheets' directions win.
    // ###########################################################################################
    [Fact]
    public void The_datasheet_comes_before_KiCads_pin_type()
    {
        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U20", Category = "IC", TechnicalNameOrValue = "556", Package = "DIP-14" },
                new ComponentEntry { BoardLabel = "U7", Category = "IC", TechnicalNameOrValue = "6510", Package = "DIP-40" },
            ],
            ComponentImages = [Baseline("U20", string.Empty, "1", "1DISCHARGE"), Baseline("U20", string.Empty, "5", "1OUTPUT"), Baseline("U7", string.Empty, "7", "A0")],
        };

        var kiCad = new KiCadPinoutSource(
            [new KiCadBoardComponent("U20", "556", "IC", "DIP-14", string.Empty), new KiCadBoardComponent("U7", "6510", "IC", "DIP-40", string.Empty)],
            label => label == "U20"
                ? [new KiCadPin("1", "1DISCHARGE", "Input", "N1")]
                : [new KiCadPin("7", "A0", "Input", "A0"), new KiCadPin("30", "D0", "Input", "D0")]);

        BoardData made = ConversionPinouts.ApplyToBoard(board, out _, out _, kiCad);

        Assert.Equal(
            ["U20|1|Output", "U20|5|Output", "U7|7|Output", "U7|30|Input/output"],
            made.ComponentPins.Select(pin => $"{pin.BoardLabel}|{pin.Pin}|{pin.Direction}"));
    }

    // ###########################################################################################
    // A CHIP CRT HOLDS THE DATASHEET OF TAKES ITS DATASHEET'S DIRECTIONS (owner request, 2026-11-17: "I would rather then
    // convert the data, so it also will show correctly in the text part of it, in the pin out page (for "Direction" in the
    // table) ... make sure it also gets done when we go-live"; and "OK to what you suggest for 1-4"): a logic chip's and
    // another IC's every pin, whatever the baselines' names or KiCad said - but the board's own "Not connected" stays, and so
    // does what the board says of a pin the datasheet leaves unconnected (a ROM's pin 1 wired to +5V). A chip the board says
    // something against the datasheet of, or one CRT does not know, is left alone. Run again, there is nothing more to do.
    // ###########################################################################################
    [Fact]
    public void A_chip_CRT_holds_the_datasheet_of_takes_its_directions_but_what_the_board_must_keep()
    {
        ComponentPinEntry Pin(string label, int pin, string name, string direction) =>
            new() { BoardLabel = label, Pin = pin.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = name, Direction = direction };

        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U20", Category = "IC", TechnicalNameOrValue = "556", Package = "DIP-14" },
                new ComponentEntry { BoardLabel = "U27", Category = "IC", TechnicalNameOrValue = "74LS08", Package = "DIP-16" },
                new ComponentEntry { BoardLabel = "U7", Category = "IC", TechnicalNameOrValue = "6510", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U4", Category = "IC", TechnicalNameOrValue = "251913-01", Package = "DIP-28" },
                new ComponentEntry { BoardLabel = "U9", Category = "IC", TechnicalNameOrValue = "XC9572" },
            ],
            ComponentPins =
            [
                Pin("U20", 1, "1DISCHARGE", "Input"),
                Pin("U20", 2, "1THRESHOLD", ""),
                Pin("U20", 5, "1OUTPUT", "Unknown"),
                Pin("U20", 7, "GND", "Power"),
                Pin("U20", 8, "2TRIGGER", "unknown"),
                Pin("U20", 9, "NC", "Not connected"),
                Pin("U20", 14, "VCC", "Unknown"),
                Pin("U27", 1, "1A", "Unknown"),
                Pin("U7", 1, "PHI0", "Unknown"),
                Pin("U7", 7, "A0", "Input"),
                Pin("U4", 1, "VCC", "Power"),
                Pin("U4", 11, "D0", "Input"),
                Pin("U9", 1, "CLK", "Unknown"),
            ],
        };

        // Counted apart from the pinouts made: these chips only took their datasheet's directions (code review, 2026-11-17).
        BoardData made = ConversionPinouts.ApplyToBoard(board, out int components, out int directed);

        Assert.Equal((0, 3), (components, directed));
        Assert.Equal(
            ["1|Output", "2|Input", "5|Output", "7|Ground", "8|Input", "9|Not connected", "14|Power"],
            made.ComponentPins.Where(pin => pin.BoardLabel == "U20").Select(pin => $"{pin.Pin}|{pin.Direction}"));
        Assert.Equal(["1|Input", "7|Output"], made.ComponentPins.Where(pin => pin.BoardLabel == "U7").Select(pin => $"{pin.Pin}|{pin.Direction}"));
        Assert.Equal(["1|Power", "11|Output"], made.ComponentPins.Where(pin => pin.BoardLabel == "U4").Select(pin => $"{pin.Pin}|{pin.Direction}"));

        // A 74LS08 in a DIP-16 is no 74LS08, and an XC9572 no chip CRT holds the datasheet of: both as they were.
        Assert.Equal("Unknown", made.ComponentPins.Single(pin => pin.BoardLabel == "U27").Direction);
        Assert.Equal("Unknown", made.ComponentPins.Single(pin => pin.BoardLabel == "U9").Direction);

        Assert.Same(made, ConversionPinouts.ApplyToBoard(made, out int again, out int directedAgain));
        Assert.Equal((0, 0), (again, directedAgain));
    }

    // The datasheet's directions go to each region's own pins while a chip's pins are told apart per region.
    [Fact]
    public void A_logic_chips_regional_pins_take_the_datasheets_directions_in_each_region()
    {
        BoardData board = new()
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U27", Category = "IC", TechnicalNameOrValue = "74LS08", Region = "PAL" },
                new ComponentEntry { BoardLabel = "U27", Category = "IC", TechnicalNameOrValue = "74LS08", Region = "NTSC" },
            ],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U27", Region = "PAL", Pin = "3", Name = "1Y", Direction = "Unknown" },
                new ComponentPinEntry { BoardLabel = "U27", Region = "NTSC", Pin = "3", Name = "BA", Direction = "Unknown" },
            ],
        };

        BoardData made = ConversionPinouts.ApplyToBoard(board, out _, out _);

        Assert.Equal(["NTSC|3|Output", "PAL|3|Output"], made.ComponentPins.Select(pin => $"{pin.Region}|{pin.Pin}|{pin.Direction}").Order());
    }

    // A component whose pins are told apart per region gets the KiCad pins in each of its regions.
    [Fact]
    public void Pins_told_apart_per_region_are_filled_in_each_region()
    {
        BoardData board = new()
        {
            Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
            ComponentPins =
            [
                new ComponentPinEntry { BoardLabel = "U1", Region = "PAL", Pin = "1", Name = "P" },
                new ComponentPinEntry { BoardLabel = "U1", Region = "NTSC", Pin = "1", Name = "N" },
            ],
        };

        BoardData made = ConversionPinouts.ApplyToBoard(board, out _, out _, KiCadWith(("U1", string.Empty)));

        Assert.Equal(["PAL|1|P", "PAL|2|B"], made.ComponentPins.Where(pin => pin.Region == "PAL").Select(pin => $"{pin.Region}|{pin.Pin}|{pin.Name}"));
        Assert.Equal(["NTSC|1|N", "NTSC|2|B"], made.ComponentPins.Where(pin => pin.Region == "NTSC").Select(pin => $"{pin.Region}|{pin.Pin}|{pin.Name}"));
    }

    // 323 of the 329 components with baselines of both regions, on the 14 boards.
    [Fact]
    public void The_same_pins_in_every_region_are_one_set_for_every_region()
    {
        BoardData made = ConversionPinouts.Apply(Board("IC", BothRegions(8)), "U1");

        List<ComponentPinEntry> pins = made.ComponentPins;

        Assert.Equal(8, pins.Count);
        Assert.All(pins, pin => Assert.Equal(string.Empty, pin.Region));
        Assert.All(pins, pin => Assert.Equal("U1", pin.BoardLabel));
        Assert.Equal("1", pins[0].Pin);
        Assert.Equal("P1", pins[0].Name);
    }

    // A name starting with a digit is no pin number: "1A" and "1B" were both pin 1, and the second dropped.
    [Fact]
    public void Pins_named_with_a_leading_digit_stay_apart()
    {
        BoardData made = ConversionPinouts.Apply(Board("IC", [Baseline("U1", "", "1A", "In"), Baseline("U1", "", "1B", "Out")]), "U1");

        Assert.Equal(["1A", "1B"], made.ComponentPins.Select(pin => pin.Pin));
    }

    [Fact]
    public void Pins_that_differ_between_the_regions_stay_per_region()
    {
        BoardData made = ConversionPinouts.Apply(
            Board("IC", [Baseline("U1", "PAL", "1", "D6"), Baseline("U1", "NTSC", "1", "D7")]),
            "U1");

        Assert.Equal(["NTSC", "PAL"], made.ComponentPins.Select(pin => pin.Region));
        Assert.Equal(["D7", "D6"], made.ComponentPins.Select(pin => pin.Name));
    }

    // The VIC-II's "SYNC/LUM explanation": a second picture of pin 15, not its name (six boards).
    [Fact]
    public void A_baseline_for_every_region_beside_regional_ones_says_nothing_about_the_pin()
    {
        BoardData made = ConversionPinouts.Apply(
            Board("IC", [Baseline("U1", "PAL", "15", "S/LUM"), Baseline("U1", "NTSC", "15", "S/LUM"), Baseline("U1", "", "15", "SYNC/LUM explanation")]),
            "U1");

        ComponentPinEntry pin = Assert.Single(made.ComponentPins);

        Assert.Equal("S/LUM", pin.Name);
        Assert.Equal(string.Empty, pin.Region);
    }

    [Fact]
    public void Baselines_of_no_region_make_pins_of_no_region()
    {
        BoardData made = ConversionPinouts.Apply(Board("IC", [Baseline("U1", "", "1", "PA3")]), "U1");

        Assert.Equal(string.Empty, Assert.Single(made.ComponentPins).Region);
    }

    [Theory]
    [InlineData("VCC", "Power")]
    [InlineData("+5V", "Power")]
    [InlineData("GND", "Ground")]
    [InlineData("VSS", "Ground")]
    [InlineData("NC", "Not connected")]
    [InlineData("A0", "Unknown")]
    public void The_direction_is_told_from_the_name_where_it_can_be(string name, string direction)
    {
        BoardData made = ConversionPinouts.Apply(Board("IC", [Baseline("U1", "", "1", name)]), "U1");

        Assert.Equal(direction, Assert.Single(made.ComponentPins).Direction);
    }

    // As "Fill from KiCad" does.
    [Fact]
    public void What_was_typed_stays_and_only_what_is_empty_is_filled_in()
    {
        BoardData made = ConversionPinouts.Apply(
            Board(
                "IC",
                [Baseline("U1", "", "1", "A0"), Baseline("U1", "", "2", "A1"), Baseline("U1", "", "3", "A2")],
                new ComponentPinEntry { BoardLabel = "U1", Pin = "1", Name = "Typed", Direction = "Input" },
                new ComponentPinEntry { BoardLabel = "U1", Pin = "2" }),
            "U1");

        Assert.Equal(["1", "2", "3"], made.ComponentPins.Select(pin => pin.Pin));
        Assert.Equal(["Typed", "A1", "A2"], made.ComponentPins.Select(pin => pin.Name));
        Assert.Equal(["Input", "Unknown", "Unknown"], made.ComponentPins.Select(pin => pin.Direction));
    }

    // "2 [64K]" is how four published baselines write pin 2 - an exact "2" says its name, not that row.
    [Fact]
    public void A_pin_written_with_a_note_is_that_pin_and_an_exact_row_names_it()
    {
        BoardData made = ConversionPinouts.Apply(
            Board("IC", [Baseline("U1", "", "2 [64K]", "A7 in 64K"), Baseline("U1", "", "2", "A7")]),
            "U1");

        ComponentPinEntry pin = Assert.Single(made.ComponentPins);

        Assert.Equal("2", pin.Pin);
        Assert.Equal("A7", pin.Name);
    }

    [Theory]
    [InlineData("IC", 40, "DIP-40")]
    [InlineData("ic", 8, "DIP-8")]
    [InlineData("IC", 7, "")]
    [InlineData("IC", 2, "")]
    [InlineData("Misc", 8, "")] // An RF modulator's eight pins are no DIP-8.
    public void The_package_is_set_only_when_it_is_no_guess(string category, int pins, string package)
    {
        BoardData made = ConversionPinouts.Apply(Board(category, BothRegions(pins)), "U1");

        Assert.Equal(package, made.Components[0].Package);
    }

    // Holes say the real case may be bigger - the Amstrad's 30 baselines of a 40-pin chip.
    [Fact]
    public void A_chip_with_a_pin_no_baseline_covers_gets_no_package()
    {
        BoardData made = ConversionPinouts.Apply(
            Board("IC", [Baseline("U1", "", "1", "A"), Baseline("U1", "", "2", "B"), Baseline("U1", "", "3", "C"), Baseline("U1", "", "6", "D")]),
            "U1");

        Assert.Equal(string.Empty, made.Components[0].Package);
        Assert.Equal(4, made.ComponentPins.Count);
    }

    [Fact]
    public void A_package_already_picked_stays()
    {
        BoardData board = Board("IC", BothRegions(8));
        board.Components[0] = new ComponentEntry { BoardLabel = "U1", Category = "IC", Package = "SIP" };

        Assert.Equal("SIP", ConversionPinouts.Apply(board, "U1").Components[0].Package);
    }

    // A component with a row of its own per region gets a package on each.
    [Fact]
    public void Each_of_a_components_regional_rows_gets_the_package()
    {
        BoardData board = Board("IC", BothRegions(8));
        board.Components.RemoveAt(0);
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "PAL" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "NTSC" });

        BoardData made = ConversionPinouts.Apply(board, "U1");

        Assert.All(made.Components.Where(row => row.BoardLabel == "U1"), row => Assert.Equal("DIP-8", row.Package));
    }

    [Fact]
    public void Nothing_to_make_is_the_same_board()
    {
        BoardData none = Board("IC", []);
        BoardData made = ConversionPinouts.Apply(Board("IC", [Baseline("U1", "", "1", "A0")]), "U1");

        Assert.Same(none, ConversionPinouts.Apply(none, "U1"));
        Assert.Same(made, ConversionPinouts.Apply(made, "U1"));
        Assert.Same(none, ConversionPinouts.Apply(none, "U99"));
    }

    [Fact]
    public void The_whole_board_is_every_component_with_a_baseline_and_says_how_many()
    {
        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry { BoardLabel = "U1", Category = "IC" },
                new ComponentEntry { BoardLabel = "U2", Category = "IC" },
                new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
            ],
            ComponentImages = [Baseline("U1", "", "1", "A"), Baseline("U2", "", "1", "B"), new ComponentImageEntry { BoardLabel = "C1", Name = "Top", File = "c.png" }],
        };

        BoardData made = ConversionPinouts.ApplyToBoard(board, out int components, out _);

        Assert.Equal(2, components);
        Assert.Equal(["U1", "U2"], made.ComponentPins.Select(pin => pin.BoardLabel));

        ConversionPinouts.ApplyToBoard(made, out int again, out _);
        Assert.Equal(0, again);
    }
}
