using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadComponentTable - a board's components from its loaded KiCad data (CRT 4.0 step 3): the first
// PCB file's footprints with their names, values and pads, the schematic symbols' datasheets, and
// each added component's pins the ones "Fill from KiCad" adds.
// ###########################################################################################
public sealed class KiCadComponentTableTests
{
    // ###########################################################################################
    // HOW MANY COMPONENTS EACH VIEW SHOWS (owner request, 2026-11-01): a side of the board by marking's
    // own rule - a through-hole part on both, a surface part on its own side, no mounting hole - and a
    // sheet by the parts drawn on it, each once however many units it has, no power symbol.
    // ###########################################################################################
    [Fact]
    public void Each_view_counts_the_components_it_shows()
    {
        static KiCadPcbPad Pad(string layer) => new() { Number = "1", Layers = [layer] };

        var root = new KiCadProjectRoot
        {
            Project = new KiCadProjectInfo
            {
                Views =
                [
                    new KiCadProjectView { DisplayName = "Board - PCB Top", Type = "pcb_top", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "Board - PCB Bottom", Type = "pcb_bottom", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "Main", Type = "schematic", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "Nowhere", Type = "schematic", SourceIndex = 9 },
                ],
            },
            Pcb =
            [
                new KiCadPcb
                {
                    Footprints =
                    [
                        new KiCadPcbFootprint { Reference = "U1", Pads = [Pad("*.Cu")] },
                        new KiCadPcbFootprint { Reference = "C1", Pads = [Pad("F.Cu")] },
                        new KiCadPcbFootprint { Reference = "C2", Pads = [Pad("B.Cu")] },
                        new KiCadPcbFootprint { Reference = "H1", FootprintName = "MountingHole:MountingHole_3mm", Pads = [Pad("*.Cu")] },
                    ],
                },
            ],
            Schematics =
            [
                new KiCadSchematic
                {
                    Symbols =
                    [
                        new KiCadSchematicSymbol { Reference = "U1" },
                        new KiCadSchematicSymbol { Reference = "U1" },
                        new KiCadSchematicSymbol { Reference = "#PWR01" },
                        new KiCadSchematicSymbol { Reference = "Y1" },
                    ],
                },
            ],
        };

        Dictionary<string, int> counts = KiCadComponentTable.CountsByView(root);

        Assert.Equal(2, counts["board - pcb top"]);
        Assert.Equal(2, counts["Board - PCB Bottom"]);
        Assert.Equal(2, counts["Main"]);
        Assert.Equal(0, counts["Nowhere"]);

        // Y1 has no footprint: only a sheet's picture can be marked from it, by its symbol.
        Assert.Equal(["Y1"], KiCadComponentTable.OnlyInSchematic(root));

        // So it counts among what KiCad can mark, once, with those on the PCB (2026-11-04: "need marking by
        // hand" until then, for what a sheet's alignment now marks).
        Assert.Equal(KiCadComponentTable.Of(root).Count + 1, BoardKiCadFacts.Of(root).Markable);

        Assert.Empty(KiCadComponentTable.CountsByView(null));
        Assert.Empty(KiCadComponentTable.OnlyInSchematic(null));
    }

    private static KiCadPcbPad Pad(string number, double x, double y, string net = "", string function = "") => new()
    {
        Number = number,
        AbsoluteCenter = new KiCadPoint2D { X = x, Y = y },
        Net = net.Length == 0 ? null : new KiCadNetRef { Id = net, Name = net },
        PinFunction = function,
    };

    private static KiCadProjectRoot Root() => new()
    {
        Pcb =
        [
            new KiCadPcb
            {
                Footprints =
                [
                    // A 24-pin DIP whose name gives no width: pin 1 and pin 24 lie 15.24 mm apart.
                    new KiCadPcbFootprint
                    {
                        Reference = "U4",
                        Value = "2364A",
                        FootprintName = "KU-14194HB-RevB-KiCad:DIP24",
                        Pads = [.. Enumerable.Range(1, 24).Select(number => number <= 12
                            ? Pad($"{number}", 0, number * 2.54, number == 1 ? "/A7" : "", number == 1 ? "A7" : "")
                            : Pad($"{number}", 15.24, (25 - number) * 2.54))],
                    },
                    new KiCadPcbFootprint
                    {
                        Reference = "C3",
                        Value = ".1",
                        FootprintName = "KU-14194HB-RevB-KiCad:CAP_P5.08mm",
                        Datasheet = "~",
                        Pads = [Pad("1", 0, 0), Pad("2", 5.08, 0)],
                    },
                ],
            },
            new KiCadPcb { Footprints = [new KiCadPcbFootprint { Reference = "U99", Pads = [Pad("1", 0, 0)] }] },
        ],
        Schematics =
        [
            new KiCadSchematic
            {
                Symbols =
                [
                    new KiCadSchematicSymbol
                    {
                        Reference = "U4",
                        Value = "2364",
                        PropertiesDetailed = [new KiCadSchematicSymbolPropertyDetailed { Name = "Datasheet", Value = "https://archive.org/download/mos_2364_rom_feb_1980/mos_2364_rom_feb_1980.pdf" }],
                    },
                ],
            },
        ],
    };

    [Fact]
    public void The_first_PCBs_footprints_are_the_components_with_their_packages_and_datasheets()
    {
        Assert.Equal(
            [
                new KiCadBoardComponent("C3", ".1", "Capacitor", "", ""),
                new KiCadBoardComponent("U4", "2364A", "IC", "DIP-24 wide", "https://archive.org/download/mos_2364_rom_feb_1980/mos_2364_rom_feb_1980.pdf"),
            ],
            KiCadComponentTable.Of(Root()));

        Assert.Empty(KiCadComponentTable.Of(null));
    }

    // The chip's pins come with it, named as KiCad names them - its net never stored.
    [Fact]
    public void An_added_chip_takes_its_pins_from_KiCad()
    {
        BoardData board = KiCadComponentTable.WithComponents(new BoardData(), Root(), out KiCadComponentsAdded added);

        Assert.Equal(new KiCadComponentsAdded(2, 1, 1), added);
        Assert.Equal(24, board.ComponentPins.Count(row => row.BoardLabel == "U4"));
        Assert.Equal(("1", "A7"), (board.ComponentPins[0].Pin, board.ComponentPins[0].Name));
        Assert.Equal(["Datasheet for 2364A"], board.ComponentLinks.Select(row => row.Name));
    }
}
