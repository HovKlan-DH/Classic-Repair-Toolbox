using System;
using System.Collections.Generic;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Following one net across a SHEET BOUNDARY (owner report, 2026-09-24).
//
// *** THE REPORTED BUG. *** Selecting CN9 on the C128 Open128 board highlighted every trace
// except two. Those two were the userport CNT1/SP1 lines, and the cause was that a hierarchical
// KiCad design renames a net at every sheet boundary:
//
//     the Userport sheet draws it labelled ...... CNT1
//     the PCB calls that copper ................. /I{slash}O/Serial Bus/CNT
//
// because on the parent I/O sheet, the Userport sheet symbol's CNT1 pin and the Serial Bus sheet
// symbol's CNT pin are joined by a run of bare wire. Component selection reads net names off the
// PCB pads, so it looked up "CNT" on a sheet that only knows "CNT1" and found nothing - while
// hovering the same wire still highlighted it, because hover hit-tests geometry and never does
// this lookup. That asymmetry is exactly what the project owner noticed.
//
// The fixture below is that real topology in miniature: two child sheets whose differently named
// pins are joined by an UNLABELLED wire run on the parent. An earlier implementation keyed the
// alias on the parent's label for the wire and passed every test I wrote for it while still
// failing the real board, so the absence of a label here is deliberate and load-bearing.
// ###########################################################################################
public sealed class KiCadSheetNetAliasesTests
{
    // ------------------------------------------------------------------ fixture helpers

    private static KiCadSchematicPathItem Wire(double x1, double y1, double x2, double y2) =>
        new()
        {
            Type = "wire",
            Points =
            [
                new KiCadPoint2D { X = x1, Y = y1 },
                new KiCadPoint2D { X = x2, Y = y2 },
            ],
        };

    private static KiCadSchematicSheetPin Pin(string name, double x, double y) =>
        new()
        {
            Name = name,
            NormalizedName = name,
            At = new KiCadPoint2DAngle { X = x, Y = y },
        };

    private static KiCadSchematicLabel Label(string text, double x, double y) =>
        new()
        {
            Type = "label",
            Text = text,
            NormalizedText = text,
            At = new KiCadPoint2DAngle { X = x, Y = y },
        };

    // The real shape: parent at index 0, two children at 1 and 2, joined by bare wire.
    private static List<KiCadSchematic> TwoChildrenJoinedByWire(
        string leftPinName,
        string rightPinName,
        IEnumerable<KiCadSchematicLabel>? parentLabels = null)
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            DisplayName = "Parent",
            Wires =
            {
                KiCadSheetNetAliasesTests.Wire(10, 10, 20, 10),
                KiCadSheetNetAliasesTests.Wire(20, 10, 30, 10),
            },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Left",
                    FileName = "left.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin(leftPinName, 10, 10) },
                },
                new KiCadSchematicChildSheet
                {
                    SheetName = "Right",
                    FileName = "right.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin(rightPinName, 30, 10) },
                },
            },
        };

        foreach (var label in parentLabels ?? [])
        {
            parent.Labels.Local.Add(label);
        }

        return
        [
            parent,
            new KiCadSchematic { Filename = "left.kicad_sch", DisplayName = "Left" },
            new KiCadSchematic { Filename = "right.kicad_sch", DisplayName = "Right" },
        ];
    }

    // ------------------------------------------------------------------ the reported bug

    // ###########################################################################################
    // *** THE REPORTED BUG, AS A TEST. *** Each child learns the other's name for the shared net,
    // so a PCB net named after one sheet resolves on the other.
    //
    // This fails against the pre-fix version, which had no alias table at all.
    // ###########################################################################################
    [Fact]
    public void Each_child_sheet_learns_the_OTHER_sheets_name_for_a_shared_net()
    {
        var aliases = KiCadSheetNetAliases.Build(
            KiCadSheetNetAliasesTests.TwoChildrenJoinedByWire("CNT1", "CNT"));

        // Sheet 1 (Left) draws CNT1 and must now also answer to CNT.
        Assert.Equal("CNT1", aliases[1]["CNT"]);

        // And symmetrically.
        Assert.Equal("CNT", aliases[2]["CNT1"]);
    }

    // ###########################################################################################
    // *** THE WIRE RUN CARRIES NO LABEL, WHICH IS THE WHOLE POINT. ***
    //
    // The first implementation of this class resolved the connection by looking for the parent's
    // label on the wire. On the real board that run has none - the two sheet pins are simply wired
    // together - so it found nothing and the bug survived the "fix". A test with a labelled wire
    // passes either way and proves nothing; this one is the one that matters.
    // ###########################################################################################
    [Fact]
    public void An_UNLABELLED_wire_run_still_connects_the_two_pins()
    {
        var schematics = KiCadSheetNetAliasesTests.TwoChildrenJoinedByWire("SP1", "SP");

        Assert.Empty(schematics[0].Labels.Local);
        Assert.Empty(schematics[0].Labels.Global);
        Assert.Empty(schematics[0].Labels.Hierarchical);

        var aliases = KiCadSheetNetAliases.Build(schematics);

        Assert.Equal("SP1", aliases[1]["SP"]);
    }

    // A label on the run is a legitimate extra name for the same net, so it is collected too - it
    // is simply not required for the pins to be connected.
    [Fact]
    public void A_label_on_the_run_becomes_an_alias_as_well()
    {
        var aliases = KiCadSheetNetAliases.Build(
            KiCadSheetNetAliasesTests.TwoChildrenJoinedByWire(
                "CNT1",
                "CNT",
                [KiCadSheetNetAliasesTests.Label("USERPORT_CNT", 20, 10)]));

        Assert.Equal("CNT1", aliases[1]["USERPORT_CNT"]);
        Assert.Equal("CNT1", aliases[1]["CNT"]);
    }

    // ------------------------------------------------------------------ not connected

    // ###########################################################################################
    // Two pins that are NOT on one run must not learn each other's names. Without this the class
    // would alias every pin on a sheet to every other, and selecting one component would light up
    // unrelated copper - far worse than the bug being fixed.
    // ###########################################################################################
    [Fact]
    public void Pins_on_SEPARATE_runs_learn_nothing_from_each_other()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires =
            {
                KiCadSheetNetAliasesTests.Wire(10, 10, 20, 10),
                KiCadSheetNetAliasesTests.Wire(10, 50, 20, 50),   // a different run entirely
            },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Left",
                    FileName = "left.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("A1", 10, 10) },
                },
                new KiCadSchematicChildSheet
                {
                    SheetName = "Right",
                    FileName = "right.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("B1", 10, 50) },
                },
            },
        };

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "left.kicad_sch" },
            new KiCadSchematic { Filename = "right.kicad_sch" },
        ]);

        Assert.Empty(aliases);
    }

    // Two pins with the SAME name are one name, not two, so there is no alias to record.
    [Fact]
    public void Identically_named_pins_produce_no_alias()
    {
        var aliases = KiCadSheetNetAliases.Build(
            KiCadSheetNetAliasesTests.TwoChildrenJoinedByWire("RESET", "RESET"));

        Assert.Empty(aliases);
    }

    // ------------------------------------------------------------------ ambiguity

    // ###########################################################################################
    // *** THE GUARD THAT MATTERS MOST. *** Highlighting the WRONG net is worse than highlighting
    // none: someone at a bench would trust it and probe the wrong pin.
    //
    // Here one child reaches the run through TWO differently named pins, so there is no single
    // "this sheet's name" for the net. Nothing is recorded for that child.
    // ###########################################################################################
    [Fact]
    public void A_child_reaching_one_run_through_TWO_pins_records_nothing_for_that_child()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires =
            {
                KiCadSheetNetAliasesTests.Wire(10, 10, 20, 10),
                KiCadSheetNetAliasesTests.Wire(20, 10, 30, 10),
                KiCadSheetNetAliasesTests.Wire(10, 10, 10, 20),
            },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Left",
                    FileName = "left.kicad_sch",
                    Pins =
                    {
                        KiCadSheetNetAliasesTests.Pin("A1", 10, 10),
                        KiCadSheetNetAliasesTests.Pin("A2", 10, 20),
                    },
                },
                new KiCadSchematicChildSheet
                {
                    SheetName = "Right",
                    FileName = "right.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("B1", 30, 10) },
                },
            },
        };

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "left.kicad_sch" },
            new KiCadSchematic { Filename = "right.kicad_sch" },
        ]);

        // Nothing for the ambiguous child...
        Assert.False(aliases.TryGetValue(1, out var left) && left.Count > 0);

        // ...but the unambiguous one still learns what it can, since its own name is not in doubt.
        Assert.Equal("B1", aliases[2]["A1"]);
    }

    // ###########################################################################################
    // Two different nets claiming ONE alias on the same sheet cancel each other out. The entry is
    // emptied rather than letting the first or last writer win, because either choice would send a
    // highlight to the wrong wire half the time.
    // ###########################################################################################
    [Fact]
    public void An_alias_claimed_by_TWO_different_nets_is_dropped()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires =
            {
                KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10),
                KiCadSheetNetAliasesTests.Wire(10, 50, 30, 50),
            },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Left",
                    FileName = "left.kicad_sch",
                    Pins =
                    {
                        KiCadSheetNetAliasesTests.Pin("OWN_A", 10, 10),
                        KiCadSheetNetAliasesTests.Pin("OWN_B", 10, 50),
                    },
                },
                new KiCadSchematicChildSheet
                {
                    SheetName = "Right",
                    FileName = "right.kicad_sch",
                    Pins =
                    {
                        // The SAME name on two unrelated runs.
                        KiCadSheetNetAliasesTests.Pin("SHARED", 30, 10),
                        KiCadSheetNetAliasesTests.Pin("SHARED", 30, 50),
                    },
                },
            },
        };

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "left.kicad_sch" },
            new KiCadSchematic { Filename = "right.kicad_sch" },
        ]);

        // Emptied, which the loader treats as "no alias" - never resolved to one of the two.
        Assert.True(string.IsNullOrEmpty(aliases[1]["SHARED"]));
    }

    // ------------------------------------------------------------------ edges

    [Fact]
    public void A_flat_design_with_no_child_sheets_produces_no_aliases()
    {
        // Single-sheet boards were unaffected by the bug, and must stay unaffected by the fix.
        var aliases = KiCadSheetNetAliases.Build(
            [new KiCadSchematic { Filename = "flat.kicad_sch" }]);

        Assert.Empty(aliases);
    }

    [Fact]
    public void An_empty_or_null_project_is_handled()
    {
        Assert.Empty(KiCadSheetNetAliases.Build([]));
        Assert.Empty(KiCadSheetNetAliases.Build(null!));
    }

    [Fact]
    public void A_child_sheet_whose_file_is_not_loaded_is_skipped()
    {
        // A project can reference a sheet file that is not on disk; it must not throw.
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Missing",
                    FileName = "nowhere.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("X", 10, 10) },
                },
            },
        };

        Assert.Empty(KiCadSheetNetAliases.Build([parent]));
    }

    // A pin sitting a hair off the wire endpoint still counts - KiCad coordinates are exact in
    // practice but not bit-exact once a sheet symbol has been rotated.
    [Fact]
    public void A_pin_within_tolerance_of_the_wire_end_still_connects()
    {
        var schematics = KiCadSheetNetAliasesTests.TwoChildrenJoinedByWire("CNT1", "CNT");

        schematics[0].ChildSheets[0].Pins[0] = new KiCadSchematicSheetPin
        {
            Name = "CNT1",
            NormalizedName = "CNT1",
            At = new KiCadPoint2DAngle { X = 10.2, Y = 10.1 },
        };

        var aliases = KiCadSheetNetAliases.Build(schematics);

        Assert.Equal("CNT1", aliases[1]["CNT"]);
    }

    // ------------------------------------------------------------------ labels partway along a wire

    // ###########################################################################################
    // *** A LABEL ATTACHES ANYWHERE ALONG A WIRE, NOT ONLY AT ITS ENDS (owner report,
    // 2026-09-24, the "I/O; Cassette Interface" sheet). ***
    //
    // On the real C128 I/O sheet the cassette WRITE pin sits at x=210.82, its wire runs back to
    // x=200.66, and the "P3" label naming that net sits at x=205.74 - the MIDDLE of the segment.
    // The PCB then calls that copper "P3", a name the cassette sheet has never heard of.
    //
    // The first version of this class walked endpoint to endpoint, so it reached the wire and still
    // found no label. That is why the cassette sheet reported four named nets and left the rest
    // dark. These coordinates are the real ones, scaled to the fixture.
    // ###########################################################################################
    [Fact]
    public void A_label_PARTWAY_ALONG_the_wire_names_the_net()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            // One long segment, exactly like the real sheet.
            Wires = { KiCadSheetNetAliasesTests.Wire(200.66, 78.74, 210.82, 78.74) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Casette Interface",
                    FileName = "cassette.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("WRITE", 210.82, 78.74) },
                },
            },
        };

        // The label sits mid-segment, touching neither end.
        parent.Labels.Local.Add(KiCadSheetNetAliasesTests.Label("P3", 205.74, 78.74));

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "cassette.kicad_sch" },
        ]);

        // The cassette sheet calls it WRITE; the PCB calls it P3. Both must now resolve.
        Assert.Equal("WRITE", aliases[1]["P3"]);
    }

    // ###########################################################################################
    // *** ONE PIN IS ENOUGH WHEN A LABEL NAMES THE NET. ***
    //
    // The original rule required two sheet pins on a run, which is the CNT1/CNT shape. The cassette
    // lines are the other shape: a single child pin wired to a label on the parent. Requiring two
    // pins missed every net drawn that way, which on the C128 is most of that sheet.
    // ###########################################################################################
    [Fact]
    public void A_run_with_ONE_pin_and_a_label_still_produces_an_alias()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Child",
                    FileName = "child.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("MOTOR", 10, 10) },
                },
            },
        };

        parent.Labels.Local.Add(KiCadSheetNetAliasesTests.Label("P5", 30, 10));

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "child.kicad_sch" },
        ]);

        Assert.Equal("MOTOR", aliases[1]["P5"]);
    }

    // A lone pin on an unlabelled run names nothing, so nothing is recorded - the guard against
    // inventing an alias out of an anonymous wire.
    [Fact]
    public void A_run_with_ONE_pin_and_NO_label_produces_nothing()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Child",
                    FileName = "child.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("MOTOR", 10, 10) },
                },
            },
        };

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "child.kicad_sch" },
        ]);

        Assert.False(aliases.TryGetValue(1, out var child) && child.Count > 0);
    }

    // ###########################################################################################
    // A point NEAR a segment but genuinely off it must not attach. The on-segment test measures
    // perpendicular distance to the segment and clamps to its ends, so a label parked well clear of
    // the wire - which happens constantly on a busy sheet - cannot capture that net.
    // ###########################################################################################
    [Fact]
    public void A_label_sitting_AWAY_from_the_wire_does_not_name_it()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Child",
                    FileName = "child.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("MOTOR", 10, 10) },
                },
            },
        };

        // Mid-span horizontally, but five millimetres off the wire.
        parent.Labels.Local.Add(KiCadSheetNetAliasesTests.Label("ELSEWHERE", 20, 15));

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "child.kicad_sch" },
        ]);

        Assert.False(aliases.TryGetValue(1, out var child) && child.Count > 0);
    }

    // And a point beyond the END of a segment does not attach either - the clamp, without which a
    // label anywhere on the infinite line through the wire would capture it.
    [Fact]
    public void A_label_BEYOND_the_end_of_the_wire_does_not_name_it()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 30, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Child",
                    FileName = "child.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("MOTOR", 10, 10) },
                },
            },
        };

        // Collinear with the wire, but well past its right-hand end.
        parent.Labels.Local.Add(KiCadSheetNetAliasesTests.Label("PAST_END", 60, 10));

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "child.kicad_sch" },
        ]);

        Assert.False(aliases.TryGetValue(1, out var child) && child.Count > 0);
    }

    // A pin landing mid-wire rather than at an end is the same rule seen from the other side, and
    // happens wherever a sheet symbol is dropped onto an existing run.
    [Fact]
    public void A_PIN_partway_along_the_wire_still_joins_the_run()
    {
        var parent = new KiCadSchematic
        {
            Filename = "parent.kicad_sch",
            Wires = { KiCadSheetNetAliasesTests.Wire(10, 10, 50, 10) },
            ChildSheets =
            {
                new KiCadSchematicChildSheet
                {
                    SheetName = "Left",
                    FileName = "left.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("A1", 10, 10) },
                },
                new KiCadSchematicChildSheet
                {
                    SheetName = "Right",
                    FileName = "right.kicad_sch",
                    Pins = { KiCadSheetNetAliasesTests.Pin("B1", 30, 10) },
                },
            },
        };

        var aliases = KiCadSheetNetAliases.Build(
        [
            parent,
            new KiCadSchematic { Filename = "left.kicad_sch" },
            new KiCadSchematic { Filename = "right.kicad_sch" },
        ]);

        Assert.Equal("A1", aliases[1]["B1"]);
        Assert.Equal("B1", aliases[2]["A1"]);
    }
}
