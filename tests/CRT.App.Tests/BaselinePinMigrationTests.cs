using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** TEMPORARY - THIS FILE AND THE TOOL GO BEFORE CRT 4.0 GOES LIVE (owner request, 2026-11-01). ***
//
// The tool making a component's pinout of its scope baselines. In CRT.App.Tests, not with the other
// CRT.Data tests, for the first test: it fails a release version while the tool is still there, which
// stops the release workflow. See .claude/CLAUDE.md, "Not done yet".
// ###########################################################################################
public sealed class BaselinePinMigrationTests
{
    [Fact]
    public void A_release_never_ships_with_the_tool_making_pinouts_of_scope_baselines()
    {
        string version = AppConfig.AppDisplayVersionString.Split('+')[0].Trim();

        Assert.True(
            version.Contains('-'),
            $"CRT {version} is a release, but the temporary tool making pinouts of scope baselines is still in it. Delete " +
            "BaselinePinMigration, its panel on the \"Pinout\" page (ComponentPane.Pages.cs, ComponentPane.axaml, the Migration* " +
            "words of ComponentPaneWording, DraftStepWording.PinsFromBaselines*) and this test.");
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

        BoardData made = BaselinePinMigration.ApplyToBoard(board, out int components, KiCadWith(("R1", "Axial"), ("C1", "Disc"), ("X9", "HC-49")));

        Assert.Equal(2, components);

        Assert.Equal("Axial", made.Components.Single(row => row.BoardLabel == "R1").Package);
        Assert.Equal(["1|A", "2|B"], made.ComponentPins.Where(pin => pin.BoardLabel == "R1").Select(pin => $"{pin.Pin}|{pin.Name}"));

        // What it had stays - the case chosen and the name typed - and what it lacked is added.
        Assert.Equal("Radial", made.Components.Single(row => row.BoardLabel == "C1").Package);
        Assert.Equal(["1|Typed", "2|B"], made.ComponentPins.Where(pin => pin.BoardLabel == "C1").Select(pin => $"{pin.Pin}|{pin.Name}"));

        Assert.DoesNotContain(made.Components, row => row.BoardLabel == "X9");

        Assert.Same(made, BaselinePinMigration.ApplyToBoard(made, out int again, KiCadWith(("R1", "Axial"), ("C1", "Disc"), ("X9", "HC-49"))));
        Assert.Equal(0, again);
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

        BoardData made = BaselinePinMigration.ApplyToBoard(board, out _, KiCadWith(("U1", string.Empty)));

        Assert.Equal(["PAL|1|P", "PAL|2|B"], made.ComponentPins.Where(pin => pin.Region == "PAL").Select(pin => $"{pin.Region}|{pin.Pin}|{pin.Name}"));
        Assert.Equal(["NTSC|1|N", "NTSC|2|B"], made.ComponentPins.Where(pin => pin.Region == "NTSC").Select(pin => $"{pin.Region}|{pin.Pin}|{pin.Name}"));
    }

    // 323 of the 329 components with baselines of both regions, on the 14 boards.
    [Fact]
    public void The_same_pins_in_every_region_are_one_set_for_every_region()
    {
        BoardData made = BaselinePinMigration.Apply(Board("IC", BothRegions(8)), "U1");

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
        BoardData made = BaselinePinMigration.Apply(Board("IC", [Baseline("U1", "", "1A", "In"), Baseline("U1", "", "1B", "Out")]), "U1");

        Assert.Equal(["1A", "1B"], made.ComponentPins.Select(pin => pin.Pin));
    }

    [Fact]
    public void Pins_that_differ_between_the_regions_stay_per_region()
    {
        BoardData made = BaselinePinMigration.Apply(
            Board("IC", [Baseline("U1", "PAL", "1", "D6"), Baseline("U1", "NTSC", "1", "D7")]),
            "U1");

        Assert.Equal(["NTSC", "PAL"], made.ComponentPins.Select(pin => pin.Region));
        Assert.Equal(["D7", "D6"], made.ComponentPins.Select(pin => pin.Name));
    }

    // The VIC-II's "SYNC/LUM explanation": a second picture of pin 15, not its name (six boards).
    [Fact]
    public void A_baseline_for_every_region_beside_regional_ones_says_nothing_about_the_pin()
    {
        BoardData made = BaselinePinMigration.Apply(
            Board("IC", [Baseline("U1", "PAL", "15", "S/LUM"), Baseline("U1", "NTSC", "15", "S/LUM"), Baseline("U1", "", "15", "SYNC/LUM explanation")]),
            "U1");

        ComponentPinEntry pin = Assert.Single(made.ComponentPins);

        Assert.Equal("S/LUM", pin.Name);
        Assert.Equal(string.Empty, pin.Region);
    }

    [Fact]
    public void Baselines_of_no_region_make_pins_of_no_region()
    {
        BoardData made = BaselinePinMigration.Apply(Board("IC", [Baseline("U1", "", "1", "PA3")]), "U1");

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
        BoardData made = BaselinePinMigration.Apply(Board("IC", [Baseline("U1", "", "1", name)]), "U1");

        Assert.Equal(direction, Assert.Single(made.ComponentPins).Direction);
    }

    // As "Fill from KiCad" does.
    [Fact]
    public void What_was_typed_stays_and_only_what_is_empty_is_filled_in()
    {
        BoardData made = BaselinePinMigration.Apply(
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
        BoardData made = BaselinePinMigration.Apply(
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
        BoardData made = BaselinePinMigration.Apply(Board(category, BothRegions(pins)), "U1");

        Assert.Equal(package, made.Components[0].Package);
    }

    // Holes say the real case may be bigger - the Amstrad's 30 baselines of a 40-pin chip.
    [Fact]
    public void A_chip_with_a_pin_no_baseline_covers_gets_no_package()
    {
        BoardData made = BaselinePinMigration.Apply(
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

        Assert.Equal("SIP", BaselinePinMigration.Apply(board, "U1").Components[0].Package);
    }

    // A component with a row of its own per region gets a package on each.
    [Fact]
    public void Each_of_a_components_regional_rows_gets_the_package()
    {
        BoardData board = Board("IC", BothRegions(8));
        board.Components.RemoveAt(0);
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "PAL" });
        board.Components.Add(new ComponentEntry { BoardLabel = "U1", Category = "IC", Region = "NTSC" });

        BoardData made = BaselinePinMigration.Apply(board, "U1");

        Assert.All(made.Components.Where(row => row.BoardLabel == "U1"), row => Assert.Equal("DIP-8", row.Package));
    }

    [Fact]
    public void Nothing_to_make_is_the_same_board()
    {
        BoardData none = Board("IC", []);
        BoardData made = BaselinePinMigration.Apply(Board("IC", [Baseline("U1", "", "1", "A0")]), "U1");

        Assert.Same(none, BaselinePinMigration.Apply(none, "U1"));
        Assert.False(BaselinePinMigration.CanRun(none, "U1"));
        Assert.Same(made, BaselinePinMigration.Apply(made, "U1"));
        Assert.False(BaselinePinMigration.CanRun(made, "U1"));
        Assert.Same(none, BaselinePinMigration.Apply(none, "U99"));
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

        BoardData made = BaselinePinMigration.ApplyToBoard(board, out int components);

        Assert.Equal(2, components);
        Assert.Equal(["U1", "U2"], made.ComponentPins.Select(pin => pin.BoardLabel));

        BaselinePinMigration.ApplyToBoard(made, out int again);
        Assert.Equal(0, again);
    }
}
