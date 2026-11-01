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

    private static KiCadMarkPlan PlanOfTop(KiCadCalibrationValues alignment) => new("Top", Pcb(), false, alignment, 420, 220, null);

    // What KiCad draws for the board's components on "Top" under `alignment`, by label.
    private static Dictionary<string, ComponentHighlightEntry> Drawn(BoardData board, KiCadCalibrationValues alignment) =>
        KiCadFootprintMarks.For(board, "Top", Pcb(), false, alignment, 420, 220, skipMarked: false).ToDictionary(mark => mark.BoardLabel);

    private static BoardData WithHighlights(params ComponentHighlightEntry[] marks) => new()
    {
        Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }, new ComponentEntry { BoardLabel = "U2", Category = "IC" }],
        ComponentHighlights = [.. marks],
    };

    private static readonly KiCadCalibrationValues Shifted = new(1, 1, 30, 10, false, false);

    // ###########################################################################################
    // *** ALIGNED AGAIN, THE MARKS KiCad DREW FOLLOW (owner request, 2026-11-01). *** A mark exactly where
    // its footprint lay under the alignment before goes where it lies now; a mark anywhere else was moved
    // or drawn by hand, and stays. Another picture's marks are not touched.
    // ###########################################################################################
    [Fact]
    public void Aligned_again_a_mark_KiCad_drew_moves_and_one_by_hand_stays()
    {
        Dictionary<string, ComponentHighlightEntry> before = Drawn(WithHighlights(), Unmoved);
        var byHand = new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U2", X = "1", Y = "2", Width = "3", Height = "4" };
        var elsewhere = new ComponentHighlightEntry { SchematicName = "Bottom", BoardLabel = "U1", X = before["U1"].X, Y = before["U1"].Y, Width = before["U1"].Width, Height = before["U1"].Height };
        BoardData board = WithHighlights(before["U1"], byHand, elsewhere);

        BoardData result = KiCadFootprintMarks.Realigned(board, PlanOfTop(Shifted), Unmoved, out int changed);

        ComponentHighlightEntry there = Drawn(board, Shifted)["U1"];
        ComponentHighlightEntry moved = result.ComponentHighlights.Single(row => row.SchematicName == "Top" && row.BoardLabel == "U1");

        Assert.Equal((there.X, there.Y, there.Width, there.Height), (moved.X, moved.Y, moved.Width, moved.Height));
        Assert.Same(byHand, result.ComponentHighlights.Single(row => row.BoardLabel == "U2"));
        Assert.Same(elsewhere, result.ComponentHighlights.Single(row => row.SchematicName == "Bottom"));
        Assert.Equal(1, changed);
    }

    // Done pressed without moving anything, and a picture with no mark of KiCad's: nothing happens - the
    // same board, so no step of the draft's history.
    [Fact]
    public void Aligned_again_with_nothing_to_follow_is_the_same_board()
    {
        BoardData drawn = WithHighlights(Drawn(WithHighlights(), Unmoved)["U1"], Drawn(WithHighlights(), Unmoved)["U2"]);
        Assert.Same(drawn, KiCadFootprintMarks.Realigned(drawn, PlanOfTop(Unmoved), Unmoved, out int unchanged));
        Assert.Equal(0, unchanged);

        BoardData allByHand = WithHighlights(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "2", Width = "3", Height = "4" });
        Assert.Same(allByHand, KiCadFootprintMarks.Realigned(allByHand, PlanOfTop(Shifted), Unmoved, out int none));
        Assert.Equal(0, none);
    }

    // ###########################################################################################
    // A footprint the bad alignment left off the picture gets its mark now, and one the new alignment puts
    // off it loses its own - each counted as a change. The first alignment here is 1000 pixels off, so
    // only a mark for U1, drawn when it still landed, is KiCad's.
    // ###########################################################################################
    [Fact]
    public void Aligned_again_a_footprint_comes_onto_or_goes_off_the_picture()
    {
        var uOffPicture = new KiCadCalibrationValues(1, 1, -180, 0, false, false);
        Dictionary<string, ComponentHighlightEntry> first = Drawn(WithHighlights(), uOffPicture);

        Assert.True(first.ContainsKey("U2"));
        Assert.False(first.ContainsKey("U1"));

        BoardData result = KiCadFootprintMarks.Realigned(WithHighlights(first["U2"]), PlanOfTop(Unmoved), uOffPicture, out int changed);

        Assert.Equal(["U1", "U2"], result.ComponentHighlights.Select(row => row.BoardLabel).Order());
        Assert.Equal(2, changed);

        Dictionary<string, ComponentHighlightEntry> drawn = Drawn(WithHighlights(), Unmoved);
        BoardData goes = KiCadFootprintMarks.Realigned(WithHighlights(drawn["U1"], drawn["U2"]), PlanOfTop(uOffPicture), Unmoved, out int went);

        Assert.DoesNotContain(goes.ComponentHighlights, row => row.BoardLabel == "U1");
        Assert.Contains(goes.ComponentHighlights, row => row.BoardLabel == "U2");
        Assert.Equal(2, went);
    }

    // ###########################################################################################
    // *** ONLY A FOOTPRINT NEWLY ON THE PICTURE GETS A MARK (code review, 2026-11-01). *** One that lay on
    // it all along with no mark had its mark deleted on purpose - a part the contributor does not want
    // shown - and aligning again must not bring it back.
    // ###########################################################################################
    [Fact]
    public void Aligned_again_a_mark_deleted_on_purpose_stays_deleted()
    {
        BoardData board = WithHighlights(Drawn(WithHighlights(), Unmoved)["U1"]);

        BoardData result = KiCadFootprintMarks.Realigned(board, PlanOfTop(Shifted), Unmoved, out int changed);

        Assert.Equal("U1", Assert.Single(result.ComponentHighlights).BoardLabel);
        Assert.Equal(1, changed);
    }

    // ###########################################################################################
    // Where each side's footprints lie (a sheet has no entry), and how many labels have a mark exactly
    // where KiCad draws it - what a card's thumbnail tells automatic marks from those by hand with.
    // ###########################################################################################
    [Fact]
    public void A_mark_where_its_footprint_lies_counts_as_automatic()
    {
        var root = new KiCadProjectRoot
        {
            Project = new KiCadProjectInfo
            {
                Views =
                [
                    new KiCadProjectView { DisplayName = "PCB top", Type = "pcb_top", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "PCB bottom", Type = "pcb_bottom", SourceIndex = 0 },
                    new KiCadProjectView { DisplayName = "Sheet 1", Type = "schematic", SourceIndex = 0 },
                ],
            },
            Pcb = [Pcb()],
        };

        Dictionary<string, KiCadViewFootprints> views = KiCadFootprintMarks.ViewFootprints(root);

        Assert.Equal(["PCB bottom", "PCB top"], views.Keys.Order());
        Assert.Equal(["U1", "U2"], views["PCB top"].ByLabel.Keys.Order());
        Assert.Equal(["Q1", "U1", "U2"], views["pcb BOTTOM"].ByLabel.Keys.Order());
        Assert.Empty(KiCadFootprintMarks.ViewFootprints(null));

        ComponentHighlightEntry u1 = Drawn(WithHighlights(), Unmoved)["U1"];
        Rect Rectangle(ComponentHighlightEntry mark) => new(double.Parse(mark.X), double.Parse(mark.Y), double.Parse(mark.Width), double.Parse(mark.Height));

        var marks = new Dictionary<string, List<Rect>>(StringComparer.OrdinalIgnoreCase)
        {
            ["u1"] = [new Rect(1, 1, 1, 1), Rectangle(u1)],
            ["U2"] = [new Rect(1, 1, 1, 1)],
            ["C9"] = [new Rect(1, 1, 1, 1)],
        };

        Assert.Equal(1, KiCadFootprintMarks.CountAutomatic(marks, views["PCB top"], Unmoved, 420, 220));
        Assert.Equal(0, KiCadFootprintMarks.CountAutomatic(marks, views["PCB top"], Shifted, 420, 220));
        Assert.Equal(0, KiCadFootprintMarks.CountAutomatic(marks, views["PCB top"], Unmoved, 0, 0));
    }

    // The Schematics tab's own mapping: across the world bounds 0..1, mirrored, scaled, then moved.
    [Fact]
    public void A_world_point_maps_onto_the_picture_as_the_copper_is_drawn()
    {
        var bounds = new Rect(0, 0, 100, 50);

        Assert.Equal(new Point(100, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, Unmoved, 1000, 500));
        Assert.Equal(new Point(900, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, new KiCadCalibrationValues(1, 1, 0, 0, true, false), 1000, 500));
        Assert.Equal(new Point(70, 50), KiCadFootprintMarks.WorldToImagePixel(10, 5, bounds, new KiCadCalibrationValues(0.5, 0.5, 20, 25, false, false), 1000, 500));
    }

    // ###########################################################################################
    // A world LENGTH - a trace's width - on the picture: the mean of the two axes' scales, a mirror
    // (a negative scale) making no difference. The Schematics tab and a card's preview both draw a
    // trace this wide (code review, 2026-10-30: the preview had a copy of the tab's formula).
    // ###########################################################################################
    [Fact]
    public void A_world_length_is_the_mean_of_the_two_scales_on_the_picture()
    {
        var bounds = new Rect(0, 0, 100, 50);

        // 10 pixels per world unit across, 10 down: a 0.25 wide trace is 2.5 pixels.
        Assert.Equal(2.5, KiCadFootprintMarks.WorldLengthOnPicture(0.25, bounds, 1, 1, 1000, 500), 9);

        // 5 across (half the scale), 20 down (twice the height): the mean, 12.5 per unit.
        Assert.Equal(12.5, KiCadFootprintMarks.WorldLengthOnPicture(1, bounds, 0.5, 1, 1000, 1000), 9);
        Assert.Equal(12.5, KiCadFootprintMarks.WorldLengthOnPicture(1, bounds, -0.5, -1, 1000, 1000), 9);
    }

    // ###########################################################################################
    // A picture is found by name, capitals ignored - and of two sharing the name, by its image file
    // (code review, 2026-10-30: a card's traces were read for the first picture of its name). With no
    // file, or one neither has, the first of the name.
    // ###########################################################################################
    [Fact]
    public void A_picture_is_found_by_name_and_of_two_of_a_name_by_its_file()
    {
        var board = new BoardData
        {
            Schematics =
            [
                new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "Images/top.png", CadName = "PCB top" },
                new BoardSchematicEntry { SchematicName = " top ", SchematicImageFile = "Images/top-2.png" },
            ],
        };

        Assert.Equal("Images/top.png", KiCadFootprintMarks.PictureRow(board, "TOP")?.SchematicImageFile);
        Assert.Equal("Images/top-2.png", KiCadFootprintMarks.PictureRow(board, "Top", " images/TOP-2.png ")?.SchematicImageFile);
        Assert.Equal("Images/top.png", KiCadFootprintMarks.PictureRow(board, "Top", "Images/gone.png")?.SchematicImageFile);
        Assert.Null(KiCadFootprintMarks.PictureRow(board, "Bottom", "Images/top.png"));

        // The plan is of the picture the file names: the second, with no KiCad view, says so; the
        // first, with one, goes on to ask for its alignment.
        var content = new BoardContent(board, [], []);
        Assert.Equal(PictureToolWording.NoKiCadView("Top"), KiCadFootprintMarks.Plan(content, "Top", null, (420, 220), "Images/top-2.png").Problem);
        Assert.Equal(PictureToolWording.NotAligned("Top"), KiCadFootprintMarks.Plan(content, "Top", null, (420, 220), "Images/top.png").Problem);
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

        // A schematic sheet is told apart from every other reason - the marking that follows an
        // alignment says nothing of it (2026-11-01).
        Assert.True(KiCadFootprintMarks.Plan(content, "Sheet", root, (420, 220)).NotABoardSide);
        Assert.False(KiCadFootprintMarks.Plan(content, "Other", root, (420, 220)).NotABoardSide);
        Assert.False(KiCadFootprintMarks.Plan(new BoardContent(board, [], []), "Top", root, (420, 220)).NotABoardSide);
        Assert.False(top.NotABoardSide);
        Assert.Equal(PictureToolWording.KiCadNotRead, KiCadFootprintMarks.Plan(content, "Top", null, (420, 220)).Problem);
        Assert.Equal(PictureToolWording.ImageMissing("Top"), KiCadFootprintMarks.Plan(content, "Top", root, null).Problem);
        Assert.Equal(
            PictureToolWording.NotAligned("Top"),
            KiCadFootprintMarks.Plan(new BoardContent(board, [], []), "Top", root, (420, 220)).Problem);
        Assert.Empty(KiCadFootprintMarks.Plan(content, "Photo", root, (420, 220)).Marks(board));
    }

    // Whether a picture has any mark - capitals ignored, as highlights are filed: the marking that
    // follows an alignment leaves such a picture alone (2026-11-01).
    [Fact]
    public void A_picture_has_marks_when_any_component_is_marked_on_it()
    {
        var board = new BoardData
        {
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top ", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" }],
        };

        Assert.True(KiCadFootprintMarks.HasMarks(board, "top"));
        Assert.False(KiCadFootprintMarks.HasMarks(board, "Bottom"));
        Assert.False(KiCadFootprintMarks.HasMarks(new BoardData(), "Top"));
    }

    // The KiCad views that are a side of the board, by name - the Pictures part's cards' (2026-11-01).
    [Fact]
    public void The_board_views_are_the_sides_of_the_board_only()
    {
        var root = new KiCadProjectRoot
        {
            Project = new KiCadProjectInfo
            {
                Views =
                [
                    new KiCadProjectView { DisplayName = " PCB top ", Type = "pcb_top" },
                    new KiCadProjectView { DisplayName = "PCB bottom", Type = "PCB_BOTTOM" },
                    new KiCadProjectView { DisplayName = "Main sheet", Type = "schematic" },
                    new KiCadProjectView { DisplayName = " ", Type = "pcb_top" },
                ],
            },
        };

        Assert.Equal(["PCB top", "PCB bottom"], BoardKiCadFacts.BoardViewsOf(root));
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
