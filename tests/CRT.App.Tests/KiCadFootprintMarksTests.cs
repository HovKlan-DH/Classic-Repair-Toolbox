using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadFootprintMarks - a component's rectangle on a picture aligned with KiCad, from its footprint
// (CRT 4.0 step 3, slice 3c: "Mark all automatically draws each component's rectangle from its
// footprint"). The numbers are worked out by hand: the board's pads span 9..51 by 9..31 mm, and the
// picture is 420 x 220 pixels, so an unmoved calibration is 10 pixels a millimetre.
// ###########################################################################################
public sealed class KiCadFootprintMarksTests
{
    private static readonly KiCadCalibrationValues Unmoved = new(1, 1, 0, 0, false, false);

    private static KiCadPcbPad Pad(double x, double y, double size, params string[] layers) => new()
    {
        Number = "1",
        AbsoluteCenter = new KiCadPoint2D { X = x, Y = y },
        Size = new KiCadSize2D { X = size, Y = size },
        Layers = [.. layers],
    };

    // U1 and U2 through the board; Q1 on the bottom's copper alone.
    private static KiCadPcb Pcb() => new()
    {
        Footprints =
        [
            new KiCadPcbFootprint { Reference = "U1", Pads = [Pad(10, 10, 2, "*.Cu"), Pad(20, 10, 2, "*.Cu")] },
            new KiCadPcbFootprint { Reference = "U2", Pads = [Pad(30, 20, 2, "*.Cu"), Pad(40, 20, 2, "*.Cu")] },
            new KiCadPcbFootprint { Reference = "Q1", Pads = [Pad(50, 30, 2, "B.Cu", "B.Paste", "B.Mask")] },
        ],
    };

    private static BoardData Board() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", Category = "IC" },
            new ComponentEntry { BoardLabel = "u2", Category = "IC", Region = "PAL" },
            new ComponentEntry { BoardLabel = "U2", Category = "IC", Region = "NTSC" },
            new ComponentEntry { BoardLabel = "Q1", Category = "Transistor" },
            new ComponentEntry { BoardLabel = "C1", Category = "Capacitor" },
        ],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Bottom", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" },
        ],
    };

    // The Schematics tab's own mapping: across the world bounds 0..1, mirrored, scaled, then moved.
    [Fact]
    public void A_world_point_maps_onto_the_picture_as_the_copper_is_drawn()
    {
        var bounds = new Rect(0, 0, 100, 50);

        Assert.Equal(new Point(100, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, Unmoved, 1000, 500));
        Assert.Equal(new Point(900, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, new KiCadCalibrationValues(1, 1, 0, 0, true, false), 1000, 500));
        Assert.Equal(new Point(70, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, new KiCadCalibrationValues(0.5, 0.5, 20, 25, false, false), 1000, 500));
    }

    [Fact]
    public void The_world_bounds_are_everything_copper()
    {
        Assert.Equal(new Rect(9, 9, 42, 22), KiCadFootprintMarks.PcbWorldBounds(Pcb()));
        Assert.Equal(default, KiCadFootprintMarks.PcbWorldBounds(new KiCadPcb()));
    }

    // ###########################################################################################
    // A footprint's rectangle is its pads, each turned by its own angle, with half a millimetre
    // around - and at least two across, or an axial resistor would be a line.
    // ###########################################################################################
    [Fact]
    public void A_footprint_is_its_pads_with_a_margin_and_at_least_two_millimetres_across()
    {
        Assert.Equal(new Rect(28.5, 18.5, 13, 3), KiCadFootprintMarks.FootprintWorldRect(Pcb().Footprints[1]));

        // A pad 4 x 1 turned a quarter: 1 across and 4 down.
        var turned = new KiCadPcbFootprint
        {
            Pads = [new KiCadPcbPad { AbsoluteCenter = new KiCadPoint2D { X = 0, Y = 0 }, Size = new KiCadSize2D { X = 4, Y = 1 }, RotationDegrees = 90 }],
        };

        Rect turnedRect = KiCadFootprintMarks.FootprintWorldRect(turned)!.Value;
        Assert.Equal(2, turnedRect.Width, 6);
        Assert.Equal(5, turnedRect.Height, 6);

        // An axial resistor's thin pads: widened about their middle to two millimetres.
        var axial = new KiCadPcbFootprint { Pads = [Pad(0, 0, 0.5), Pad(10, 0, 0.5)] };
        Assert.Equal(new Rect(-0.75, -1, 11.5, 2), KiCadFootprintMarks.FootprintWorldRect(axial));

        Assert.Null(KiCadFootprintMarks.FootprintWorldRect(new KiCadPcbFootprint { Pads = [new KiCadPcbPad { Number = "1" }] }));
    }

    [Fact]
    public void A_through_hole_part_is_seen_from_both_sides_and_a_surface_part_from_its_own()
    {
        KiCadPcb pcb = Pcb();

        Assert.True(KiCadFootprintMarks.IsOnSide(pcb.Footprints[0], "F.Cu"));
        Assert.True(KiCadFootprintMarks.IsOnSide(pcb.Footprints[0], "B.Cu"));
        Assert.False(KiCadFootprintMarks.IsOnSide(pcb.Footprints[2], "F.Cu"));
        Assert.True(KiCadFootprintMarks.IsOnSide(pcb.Footprints[2], "B.Cu"));
    }

    // ###########################################################################################
    // On the picture: whole pixels, cut to the picture; a mirrored calibration turns the corners
    // round and gives the same rectangle where it lands; one wholly off the picture is none.
    // ###########################################################################################
    [Fact]
    public void A_rectangle_lands_in_whole_pixels_cut_to_the_picture()
    {
        var bounds = new Rect(9, 9, 42, 22);

        Assert.Equal(new PixelRect(195, 95, 130, 30), KiCadFootprintMarks.OnPicture(new Rect(28.5, 18.5, 13, 3), bounds, Unmoved, 420, 220));
        Assert.Equal(new PixelRect(95, 95, 130, 30), KiCadFootprintMarks.OnPicture(new Rect(28.5, 18.5, 13, 3), bounds, new KiCadCalibrationValues(1, 1, 0, 0, true, false), 420, 220));
        Assert.Equal(new PixelRect(395, 195, 25, 25), KiCadFootprintMarks.OnPicture(new Rect(48.5, 28.5, 3, 3), bounds, Unmoved, 420, 220));
        Assert.Null(KiCadFootprintMarks.OnPicture(new Rect(100, 100, 3, 3), bounds, Unmoved, 420, 220));
        Assert.Null(KiCadFootprintMarks.OnPicture(new Rect(28.5, 18.5, 13, 3), bounds, Unmoved, 0, 220));
    }

    // ###########################################################################################
    // The marks: every component of the board seen from the picture's side with no mark on this
    // picture yet - one per label whatever its regions, under the board's own spelling; never one
    // already drawn, nor one KiCad has no footprint for.
    // ###########################################################################################
    [Fact]
    public void Only_the_components_not_marked_yet_get_a_mark_from_the_side_seen()
    {
        List<ComponentHighlightEntry> top = KiCadFootprintMarks.For(Board(), " Top ", Pcb(), bottom: false, Unmoved, 420, 220);

        Assert.Equal(
            [("Top", "u2", "195", "95", "130", "30")],
            top.Select(row => (row.SchematicName, row.BoardLabel, row.X, row.Y, row.Width, row.Height)));

        List<ComponentHighlightEntry> bottom = KiCadFootprintMarks.For(Board(), "Bottom", Pcb(), bottom: true, Unmoved, 420, 220);

        Assert.Equal(["u2", "Q1"], bottom.Select(row => row.BoardLabel));
    }

    // ###########################################################################################
    // The plan for a picture: the PCB its CAD name shows and from which side, its alignment and its
    // size - or the first thing missing, in the contributor's words: no such picture, no CAD name,
    // not aligned, no KiCad data, a view KiCad does not have, a schematic sheet, an unread image.
    // ###########################################################################################
    [Fact]
    public void Marking_is_planned_from_the_pictures_view_alignment_and_size_or_says_what_is_missing()
    {
        BoardData board = Board().WithSchematics(
        [
            new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
            new BoardSchematicEntry { SchematicName = "Bottom", SchematicImageFile = "Images/bottom.png", CadName = "PCB bottom" },
            new BoardSchematicEntry { SchematicName = "Sheet", SchematicImageFile = "Images/sheet.png", CadName = "Main sheet" },
            new BoardSchematicEntry { SchematicName = "Photo", SchematicImageFile = "Images/photo.jpg" },
            new BoardSchematicEntry { SchematicName = "Other", SchematicImageFile = "Images/other.png", CadName = "Not a view" },
        ]);

        KiCadCalibrationEntry[] calibrations =
        [
            new() { SchematicName = "top", CadName = "PCB top", ScaleX = 1, ScaleY = 1, MirrorX = true },
            new() { SchematicName = "Bottom", CadName = "PCB bottom", ScaleX = 1, ScaleY = 1 },
            new() { SchematicName = "Sheet", CadName = "Main sheet", ScaleX = 1, ScaleY = 1 },
            new() { SchematicName = "Other", CadName = "Not a view", ScaleX = 1, ScaleY = 1 },
        ];

        var content = new BoardContent(board, calibrations, []);
        KiCadPcb pcb = Pcb();
        var root = new KiCadProjectRoot
        {
            Pcb = [pcb],
            Project = new KiCadProjectInfo
            {
                Views =
                [
                    new KiCadProjectView { DisplayName = "PCB top", Type = "pcb_top", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "PCB bottom", Type = "pcb_bottom", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "Main sheet", Type = "schematic", SourceIndex = 0 },
                ],
            },
        };

        KiCadMarkPlan top = KiCadFootprintMarks.Plan(content, " Top ", root, (420, 220));
        Assert.Null(top.Problem);
        Assert.Same(pcb, top.Pcb);
        Assert.False(top.Bottom);
        Assert.True(top.Calibration.MirrorX);
        Assert.Equal(["u2"], top.Marks(board).Select(row => row.BoardLabel));

        Assert.True(KiCadFootprintMarks.Plan(content, "Bottom", root, (420, 220)).Bottom);

        Assert.Equal(PictureToolWording.PictureNotFound("Nowhere"), KiCadFootprintMarks.Plan(content, "Nowhere", root, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.NoKiCadView("Photo"), KiCadFootprintMarks.Plan(content, "Photo", root, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.NotABoardView("Sheet"), KiCadFootprintMarks.Plan(content, "Sheet", root, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.NoKiCadView("Other"), KiCadFootprintMarks.Plan(content, "Other", root, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.KiCadNotRead, KiCadFootprintMarks.Plan(content, "Top", null, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.ImageMissing("Top"), KiCadFootprintMarks.Plan(content, "Top", root, null).Problem);
        Assert.Equal(
            PictureToolWording.NotAligned("Top"),
            KiCadFootprintMarks.Plan(new BoardContent(board, [], []), "Top", root, (420, 220)).Problem);
        Assert.Empty(KiCadFootprintMarks.Plan(content, "Photo", root, (420, 220)).Marks(board));
    }

    // The marks go after the picture's own, so a picture's rows stay together.
    [Fact]
    public void The_marks_go_after_the_pictures_own()
    {
        BoardData board = Board();
        List<ComponentHighlightEntry> marks = KiCadFootprintMarks.For(board, "Top", Pcb(), bottom: false, Unmoved, 420, 220);

        BoardData marked = KiCadFootprintMarks.WithMarks(board, marks);

        Assert.Equal([("Top", "U1"), ("Top", "u2"), ("Bottom", "U1")], marked.ComponentHighlights.Select(row => (row.SchematicName, row.BoardLabel)));
        Assert.Same(board, KiCadFootprintMarks.WithMarks(board, []));
        Assert.Equal(2, board.ComponentHighlights.Count);
    }
}
