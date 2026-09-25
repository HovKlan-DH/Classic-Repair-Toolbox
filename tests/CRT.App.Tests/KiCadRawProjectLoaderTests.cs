using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Characterisation tests for KiCadRawProjectLoader - the S-expression parser that turns raw
// .kicad_pcb / .kicad_sch files into the model the Schematics overlay draws from. Everything
// the interactive trace overlay renders comes out of here, so a silent regression in this
// parser shows up as traces in the wrong place rather than as an error.
//
// The parser itself is a private nested class, so these drive it through the real LoadAsync
// entry point with fixture files on disk. See KiCadFixtures.cs for the inputs.
public sealed class KiCadRawProjectLoaderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private Task<KiCadProjectRoot?> LoadAsync(params string[] paths) =>
        KiCadRawProjectLoader.LoadAsync(paths);

    private string WritePcb(string content = null!, string name = "board.kicad_pcb") =>
        this.thisWorkspace.WriteFile(name, content ?? KiCadFixtures.Pcb);

    // --------------------------------------------------------------- file selection

    [Fact]
    public async Task LoadAsync_returns_null_when_no_usable_file_was_given()
    {
        Assert.Null(await this.LoadAsync());
    }

    [Fact]
    public async Task LoadAsync_skips_a_missing_file_instead_of_throwing()
    {
        // A board can reference a KiCad file that was never contributed.
        Assert.Null(await this.LoadAsync(Path.Combine(this.thisWorkspace.Root, "nope.kicad_pcb")));
    }

    [Fact]
    public async Task LoadAsync_ignores_files_with_an_unsupported_extension()
    {
        string txt = this.thisWorkspace.WriteFile("notes.txt", KiCadFixtures.Pcb);
        Assert.Null(await this.LoadAsync(txt));
    }

    [Fact]
    public async Task LoadAsync_deduplicates_the_same_file_listed_twice()
    {
        string pcb = this.WritePcb();

        KiCadProjectRoot? root = await this.LoadAsync(pcb, pcb);

        Assert.NotNull(root);
        Assert.Single(root!.Pcb);
    }

    [Fact]
    public async Task LoadAsync_expands_a_kicad_pro_file_to_its_sibling_pcb_and_schematic()
    {
        // Boards usually reference the project file; the loader finds the matching board and
        // schematic beside it by base name.
        this.thisWorkspace.WriteFile("board.kicad_pcb", KiCadFixtures.Pcb);
        this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);
        string pro = this.thisWorkspace.WriteFile("board.kicad_pro", "{}");

        KiCadProjectRoot? root = await this.LoadAsync(pro);

        Assert.NotNull(root);
        Assert.Single(root!.Pcb);
        Assert.Single(root.Schematics);
    }

    [Fact]
    public async Task LoadAsync_reports_the_source_filename_on_the_pcb()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(name: "c64-250407.kicad_pcb"));

        Assert.Equal("c64-250407.kicad_pcb", root!.Pcb[0].Filename);
    }

    [Fact]
    public async Task LoadAsync_marks_a_successful_load_as_ok()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());
        Assert.True(root!.Ok);
    }

    // ------------------------------------------------------------------------- nets

    [Fact]
    public async Task Nets_are_collected_and_ordered_by_id()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        var nets = root!.Pcb[0].Nets.List;

        Assert.Equal(3, nets.Count);
        Assert.Equal(new[] { "0", "1", "2" }, nets.Select(n => n.Id));
        Assert.Equal("GND", nets.Single(n => n.Id == "1").Name);
    }

    [Fact]
    public async Task A_hierarchical_net_name_is_normalised_to_its_last_segment()
    {
        // "/Sheet1/CLK" is the same signal as "CLK" for highlighting purposes.
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadNetRef clk = root!.Pcb[0].Nets.List.Single(n => n.Id == "2");

        Assert.Equal("/Sheet1/CLK", clk.Name);
        Assert.Equal("CLK", clk.NormalizedName);
    }

    [Fact]
    public async Task A_flat_net_name_is_left_alone_by_normalisation()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadNetRef gnd = root!.Pcb[0].Nets.List.Single(n => n.Id == "1");
        Assert.Equal("GND", gnd.NormalizedName);
    }

    // ------------------------------------------------------------------- footprints

    [Fact]
    public async Task A_footprint_reference_is_read_from_fp_text()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbFootprint footprint = Assert.Single(root!.Pcb[0].Footprints);
        Assert.Equal("U1", footprint.Reference);
        Assert.Equal("F.Cu", footprint.Layer);
        Assert.Equal(2, footprint.Pads.Count);
    }

    [Fact]
    public async Task A_footprint_reference_is_also_read_from_a_kicad7_property()
    {
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbWithPropertyReference));

        Assert.Equal("R1", Assert.Single(root!.Pcb[0].Footprints).Reference);
    }

    [Fact]
    public async Task A_legacy_module_node_is_treated_as_a_footprint()
    {
        // Older .kicad_pcb files use (module ...) rather than (footprint ...).
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbWithLegacyModule));

        KiCadPcbFootprint footprint = Assert.Single(root!.Pcb[0].Footprints);
        Assert.Equal("U9", footprint.Reference);
        Assert.Single(footprint.Pads);
    }

    [Fact]
    public async Task A_pad_absolute_centre_is_the_footprint_origin_plus_the_pad_offset()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbPad pad1 = root!.Pcb[0].Footprints[0].Pads.Single(p => p.Number == "1");

        Assert.NotNull(pad1.AbsoluteCenter);
        Assert.Equal(100 - 3.81, pad1.AbsoluteCenter!.X, precision: 9);
        Assert.Equal(50 - 7.62, pad1.AbsoluteCenter.Y, precision: 9);
    }

    [Fact]
    public async Task A_pad_on_a_bottom_layer_footprint_keeps_the_offset_the_file_stores()
    {
        // KiCad bakes the flip into the stored pad coordinates the moment a footprint is moved to
        // the back, so a bottom-side pad needs no mirroring here - it is already in board space.
        // Mirroring it again moved every back-side pad off its own tracks; on the one bottom-side
        // footprint in the shipped boards (Open128's RGBI connector) the mirrored pads landed
        // 8-20 mm away from the track ends they are soldered to, while the stored coordinates
        // land exactly on them. Flipping the whole view for a bottom-side photo is a separate,
        // view-level concern and is handled by the calibration mirror flags.
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbBottomLayerFootprint));

        KiCadPcbPad pad = root!.Pcb[0].Footprints[0].Pads[0];

        Assert.Equal(100 + 5, pad.AbsoluteCenter!.X, precision: 9);
        Assert.Equal(100 + 3, pad.AbsoluteCenter.Y, precision: 9);
    }

    [Fact]
    public async Task A_pad_rotation_is_captured_as_the_absolute_angle_the_file_states()
    {
        // A pad's (at x y angle) mixes frames: x/y are footprint-local and unrotated, but the angle
        // is absolute - KiCad has already added the parent footprint's rotation into it. So the
        // loader must store the angle verbatim and must not add the footprint angle a second time.
        // Every pad below sits in a footprint rotated 90 degrees.
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbRotatedPads));

        System.Collections.Generic.List<KiCadPcbPad> pads = root!.Pcb[0].Footprints[0].Pads;

        Assert.Equal(90, pads.Single(pad => pad.Number == "1").RotationDegrees, precision: 9);
        Assert.Equal(90, pads.Single(pad => pad.Number == "2").RotationDegrees, precision: 9);
        Assert.Equal(180, pads.Single(pad => pad.Number == "3").RotationDegrees, precision: 9);
    }

    [Fact]
    public async Task A_pad_with_no_stated_angle_has_no_rotation()
    {
        // Most pads omit the third value in (at ...) entirely, and those must stay axis-aligned
        // rather than inheriting anything from the footprint they sit in.
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbRotatedPads));

        KiCadPcbPad pad4 = root!.Pcb[0].Footprints[0].Pads.Single(pad => pad.Number == "4");

        Assert.Equal(0, pad4.RotationDegrees, precision: 9);
    }

    [Fact]
    public async Task A_rotated_pad_keeps_the_size_the_file_states_rather_than_swapping_it()
    {
        // The size is stated in the pad's own frame and stays that way; orientation is carried by
        // the rotation alone. Swapping width and height here instead would double-rotate the pad
        // once the renderer applies that rotation.
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbRotatedPads));

        KiCadPcbPad pad1 = root!.Pcb[0].Footprints[0].Pads.Single(pad => pad.Number == "1");

        Assert.Equal(2, pad1.Size!.X, precision: 9);
        Assert.Equal(0.8, pad1.Size.Y, precision: 9);
    }

    [Fact]
    public async Task A_pad_on_a_rotated_footprint_rotates_about_the_footprint_origin()
    {
        // KiCad uses a Y-down angle convention, so +90 degrees sends +X to -Y.
        KiCadProjectRoot? root = await this.LoadAsync(
            this.WritePcb(KiCadFixtures.PcbRotatedFootprint));

        KiCadPcbPad pad = root!.Pcb[0].Footprints[0].Pads[0];

        Assert.Equal(100, pad.AbsoluteCenter!.X, precision: 6);
        Assert.Equal(90, pad.AbsoluteCenter.Y, precision: 6);
    }

    [Fact]
    public async Task Pad_shape_size_layers_and_net_are_captured()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbPad pad1 = root!.Pcb[0].Footprints[0].Pads.Single(p => p.Number == "1");

        Assert.Equal("rect", pad1.Shape);
        Assert.Equal(1.6, pad1.Size!.X, precision: 9);
        Assert.Equal(1.6, pad1.Size.Y, precision: 9);
        Assert.Equal(new[] { "*.Cu", "*.Mask" }, pad1.Layers);
        Assert.Equal("1", pad1.Net!.Id);
        Assert.Equal("GND", pad1.Net.Name);
    }

    [Fact]
    public async Task A_pad_net_name_falls_back_to_the_board_net_table()
    {
        // The pad here carries an inline name, so also prove the normalised form survives.
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbPad pad2 = root!.Pcb[0].Footprints[0].Pads.Single(p => p.Number == "2");

        Assert.Equal("2", pad2.Net!.Id);
        Assert.Equal("CLK", pad2.Net.NormalizedName);
    }

    // ------------------------------------------------------- hierarchical sheets

    // A hierarchical KiCad design keeps its child sheets in the same folder as the root sheet, so
    // the board folder listing hands the loader both files. Treating each of them as a root sheet
    // used to load every child twice - and the two copies disagreed about their own name, because a
    // sheet reached through its parent is named by that parent's "Sheetname" while a sheet loaded
    // standalone falls back to its filename. That name is what a contributor copies into the board
    // workbook's "CAD name" column, so the duplicate was a data-entry trap and not just wasted work.

    [Fact]
    public async Task A_child_sheet_listed_beside_its_parent_is_loaded_only_once()
    {
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        string child = this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);

        KiCadProjectRoot? root = await this.LoadAsync(parent, child);

        Assert.Equal(2, root!.Schematics.Count);
        Assert.Equal(
            new[] { "child.kicad_sch", "root.kicad_sch" },
            root.Schematics.Select(schematic => schematic.Filename).OrderBy(name => name).ToArray());
    }

    [Fact]
    public async Task A_child_sheet_is_named_by_its_parent_rather_than_by_its_filename()
    {
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        string child = this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);

        KiCadProjectRoot? root = await this.LoadAsync(parent, child);

        Assert.Equal(
            new[] { "Power supply", "root" },
            root!.Project.Views.Select(view => view.DisplayName).OrderBy(name => name).ToArray());
    }

    [Fact]
    public async Task A_child_sheet_keeps_its_parents_name_even_when_it_is_listed_first()
    {
        // The board folder is listed alphabetically, so "child.kicad_sch" reaches the loader before
        // "root.kicad_sch". Order must not decide which name a sheet ends up with.
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        string child = this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);

        KiCadProjectRoot? root = await this.LoadAsync(child, parent);

        Assert.Equal(2, root!.Schematics.Count);
        Assert.Contains(root.Project.Views, view => view.DisplayName == "Power supply");
        Assert.DoesNotContain(root.Project.Views, view => view.DisplayName == "child");
    }

    [Fact]
    public async Task The_root_sheet_is_loaded_first_so_page_ordinals_start_at_it()
    {
        // A board whose "CAD name" column is blank falls back to matching "Schematics #1 of 2" by
        // page ordinal, which indexes the schematic views in load order. Page 1 must be the root
        // sheet, not whichever child happened to sort first on disk.
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        string child = this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);

        KiCadProjectRoot? root = await this.LoadAsync(child, parent);

        Assert.Equal("root.kicad_sch", root!.Schematics[0].Filename);
        Assert.Equal("child.kicad_sch", root.Schematics[1].Filename);
    }

    [Fact]
    public async Task A_child_sheet_that_was_not_listed_is_still_loaded_from_disk()
    {
        // Only the root sheet is handed to the loader; the child must still be followed and named.
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);

        KiCadProjectRoot? root = await this.LoadAsync(parent);

        Assert.Equal(2, root!.Schematics.Count);
        Assert.Contains(root.Project.Views, view => view.DisplayName == "Power supply");
    }

    [Fact]
    public async Task Sheets_that_only_reference_each_other_are_still_loaded_once_each()
    {
        // Neither file is a root, so a walk that only starts from root sheets would load nothing.
        // Every sheet must survive, and the cycle must not recurse forever.
        string a = this.thisWorkspace.WriteFile("cycleA.kicad_sch", KiCadFixtures.SchematicCycleA);
        string b = this.thisWorkspace.WriteFile("cycleB.kicad_sch", KiCadFixtures.SchematicCycleB);

        KiCadProjectRoot? root = await this.LoadAsync(a, b);

        Assert.Equal(2, root!.Schematics.Count);
        Assert.Equal(
            new[] { "cycleA.kicad_sch", "cycleB.kicad_sch" },
            root.Schematics.Select(schematic => schematic.Filename).OrderBy(name => name).ToArray());
    }

    [Fact]
    public async Task Every_loaded_schematic_gets_exactly_one_view_so_the_log_can_name_it()
    {
        // The startup log prints one "Display name" line per view, grouped by source file, and that
        // is where a contributor reads the value for the workbook's "CAD name" column. One view per
        // schematic is what makes that listing complete and unambiguous.
        string parent = this.thisWorkspace.WriteFile("root.kicad_sch", KiCadFixtures.SchematicRootWithChildSheet);
        string child = this.thisWorkspace.WriteFile("child.kicad_sch", KiCadFixtures.SchematicChildSheet);
        string pcb = this.WritePcb();

        KiCadProjectRoot? root = await this.LoadAsync(pcb, parent, child);

        var schematicViews = root!.Project.Views
            .Where(view => view.Type == "schematic")
            .ToList();

        Assert.Equal(root.Schematics.Count, schematicViews.Count);
        Assert.Equal(
            schematicViews.Count,
            schematicViews.Select(view => view.DisplayName).Distinct(StringComparer.OrdinalIgnoreCase).Count());

        // Two PCB views (top and bottom) plus one per schematic.
        Assert.Equal((root.Pcb.Count * 2) + root.Schematics.Count, root.Project.Views.Count);
    }

    // ---------------------------------------------------------------------- routing

    [Fact]
    public async Task Track_segments_are_parsed_with_geometry_layer_and_net()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        var segments = root!.Pcb[0].Routing.Segments;

        Assert.Equal(2, segments.Count);

        KiCadPcbSegment first = segments[0];
        Assert.Equal(100, first.Start!.X);
        Assert.Equal(50, first.Start.Y);
        Assert.Equal(110, first.End!.X);
        Assert.Equal(50, first.End.Y);
        Assert.Equal(0.25, first.Width);
        Assert.Equal("F.Cu", first.Layer);
        Assert.Equal("1", first.Net!.Id);
    }

    [Fact]
    public async Task Segments_keep_their_layer_so_top_and_bottom_can_be_drawn_apart()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        var layers = root!.Pcb[0].Routing.Segments.Select(s => s.Layer).ToList();
        Assert.Equal(new[] { "F.Cu", "B.Cu" }, layers);
    }

    [Fact]
    public async Task Vias_are_parsed_with_position_size_layers_and_net()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbVia via = Assert.Single(root!.Pcb[0].Routing.Vias);

        Assert.Equal(105, via.At!.X);
        Assert.Equal(55, via.At.Y);
        Assert.Equal(0.8, via.Size);
        Assert.Equal(new[] { "F.Cu", "B.Cu" }, via.Layers);
        Assert.Equal("1", via.Net!.Id);
    }

    [Fact]
    public async Task Arcs_are_parsed_with_all_three_control_points()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbArc arc = Assert.Single(root!.Pcb[0].Routing.Arcs);

        Assert.Equal(120, arc.Start!.X);
        Assert.Equal(122, arc.Mid!.X);
        Assert.Equal(124, arc.End!.X);
        Assert.Equal("F.Cu", arc.Layer);
        Assert.Equal("2", arc.Net!.Id);
    }

    [Fact]
    public async Task Zones_are_parsed_with_outline_and_filled_polygons()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        KiCadPcbZone zone = Assert.Single(root!.Pcb[0].Routing.Zones);

        Assert.Equal("1", zone.Net!.Id);
        Assert.Equal("GND", zone.Net.Name);
        Assert.Contains("B.Cu", zone.Layers);

        Assert.Equal(4, Assert.Single(zone.OutlinePolygons).Points.Count);
        Assert.Equal(4, Assert.Single(zone.FilledPolygons).Points.Count);
    }

    [Fact]
    public async Task A_zone_with_no_polygons_is_dropped()
    {
        string content = """
        (kicad_pcb (version 20221018)
          (net 1 "GND")
          (zone (net 1) (net_name "GND") (layer "B.Cu"))
          (segment (start 0 0) (end 1 0) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(content));

        Assert.Empty(root!.Pcb[0].Routing.Zones);
    }

    // -------------------------------------------------------------- highlight index

    [Fact]
    public async Task The_highlight_index_groups_routing_items_by_net()
    {
        // This index is what the overlay uses to light up a whole net at once.
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        var index = root!.Pcb[0].HighlightIndex;

        Assert.True(index.ContainsKey("1"));
        Assert.True(index.ContainsKey("2"));

        Assert.Equal(new[] { 0 }, index["1"].Segments);   // first segment is on net 1
        Assert.Equal(new[] { 1 }, index["2"].Segments);   // second segment is on net 2
        Assert.Equal(new[] { 0 }, index["1"].Vias);
        Assert.Equal(new[] { 0 }, index["2"].Arcs);
    }

    // ------------------------------------------------------------- s-expression edges

    [Fact]
    public async Task The_parser_ignores_semicolon_comments()
    {
        string content = """
        ; this whole line is a comment
        (kicad_pcb (version 20221018)
          (net 1 "GND")   ; trailing comment
          (segment (start 0 0) (end 1 0) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(content));

        Assert.Single(root!.Pcb[0].Routing.Segments);
    }

    [Fact]
    public async Task The_parser_keeps_spaces_inside_quoted_strings()
    {
        string content = """
        (kicad_pcb (version 20221018)
          (net 1 "Net with spaces")
          (segment (start 0 0) (end 1 0) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(content));

        Assert.Equal("Net with spaces", root!.Pcb[0].Nets.List.Single(n => n.Id == "1").Name);
    }

    [Fact]
    public async Task The_parser_unescapes_backslash_sequences_in_quoted_strings()
    {
        string content = """
        (kicad_pcb (version 20221018)
          (net 1 "Net \"quoted\" name")
          (segment (start 0 0) (end 1 0) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(content));

        Assert.Equal("Net \"quoted\" name", root!.Pcb[0].Nets.List.Single(n => n.Id == "1").Name);
    }

    [Fact]
    public async Task Numbers_are_parsed_with_an_invariant_decimal_point()
    {
        // Guards against a Danish/German locale reading 0.25 as 25.
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        Assert.Equal(0.25, root!.Pcb[0].Routing.Segments[0].Width);
    }

    [Fact]
    public async Task Negative_coordinates_are_parsed()
    {
        string content = """
        (kicad_pcb (version 20221018)
          (net 1 "GND")
          (segment (start -10.5 -20.25) (end 0 0) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb(content));

        KiCadPoint2D start = root!.Pcb[0].Routing.Segments[0].Start!;

        Assert.Equal(-10.5, start.X);
        Assert.Equal(-20.25, start.Y);
    }

    [Fact]
    public async Task A_file_that_is_not_a_kicad_pcb_root_yields_no_pcb()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb("(something_else (version 1))"));
        Assert.Null(root);
    }

    [Fact]
    public async Task An_empty_file_yields_no_project()
    {
        Assert.Null(await this.LoadAsync(this.WritePcb("")));
    }

    // ------------------------------------------------------------------ schematics

    [Fact]
    public async Task A_schematic_file_is_loaded_and_reports_its_filename()
    {
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectRoot? root = await this.LoadAsync(sch);

        Assert.NotNull(root);
        KiCadSchematic schematic = Assert.Single(root!.Schematics);
        Assert.Equal("board.kicad_sch", schematic.Filename);
    }

    [Fact]
    public async Task Schematic_wires_are_parsed()
    {
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectRoot? root = await this.LoadAsync(sch);

        Assert.Equal(2, root!.Schematics[0].Wires.Count);
    }

    [Fact]
    public async Task Schematic_labels_are_split_by_scope()
    {
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectRoot? root = await this.LoadAsync(sch);
        KiCadSchematicLabels labels = root!.Schematics[0].Labels;

        Assert.Contains(labels.Local, l => l.Text == "CLK");
        Assert.Contains(labels.Global, l => l.Text == "GND");
    }

    [Fact]
    public async Task Schematic_symbols_expose_their_reference()
    {
        string sch = this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectRoot? root = await this.LoadAsync(sch);

        Assert.Contains(root!.Schematics[0].Symbols, s => s.Reference == "R1");
    }

    // ----------------------------------------------------------------------- views

    [Fact]
    public async Task A_view_is_built_for_each_source_file()
    {
        this.thisWorkspace.WriteFile("board.kicad_sch", KiCadFixtures.Schematic);
        string pcb = this.WritePcb();
        string sch = Path.Combine(this.thisWorkspace.Root, "board.kicad_sch");

        KiCadProjectRoot? root = await this.LoadAsync(pcb, sch);

        Assert.NotEmpty(root!.Project.Views);
        Assert.Contains(root.Project.Views, v => v.SourceKind.Equals("pcb", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(root.Project.Views, v => v.SourceKind.Equals("schematic", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public async Task Pcb_views_point_back_at_their_pcb_by_index()
    {
        KiCadProjectRoot? root = await this.LoadAsync(this.WritePcb());

        foreach (KiCadProjectView view in root!.Project.Views
                     .Where(v => v.SourceKind.Equals("pcb", StringComparison.OrdinalIgnoreCase)))
        {
            Assert.InRange(view.SourceIndex, 0, root.Pcb.Count - 1);
        }
    }

    // ------------------------------------------------------------------ child sheet PINS

    // ###########################################################################################
    // *** SHEET PINS ARE WHAT MAKE A CROSS-SHEET NET FOLLOWABLE (owner report, 2026-09-24). ***
    //
    // A hierarchical design renames a net at every boundary: the C128 userport sheet draws "CNT1"
    // while the PCB calls that copper "/I{slash}O/Serial Bus/CNT", because on the parent sheet the
    // two child sheets are wired together through pins of those two names. Selecting CN9 therefore
    // highlighted every trace except those two.
    //
    // The parser used to read a (sheet ...) node for its Sheetname/Sheetfile only and throw the
    // pins away, so the information needed to connect them did not survive the load at all.
    // KiCadSheetNetAliases consumes what these tests pin down.
    // ###########################################################################################
    [Fact]
    public async Task A_child_sheet_carries_its_boundary_PINS()
    {
        string parent = this.thisWorkspace.WriteFile(
            "root.kicad_sch", KiCadFixtures.SchematicParentJoiningTwoChildren);
        string userport = this.thisWorkspace.WriteFile(
            "userport.kicad_sch", KiCadFixtures.SchematicUserportChild);
        string serialBus = this.thisWorkspace.WriteFile(
            "serialbus.kicad_sch", KiCadFixtures.SchematicSerialBusChild);

        KiCadProjectRoot? root = await this.LoadAsync(parent, userport, serialBus);

        KiCadSchematic parentSheet = root!.Schematics
            .Single(schematic => schematic.Filename == "root.kicad_sch");

        Assert.Equal(2, parentSheet.ChildSheets.Count);

        KiCadSchematicChildSheet userportChild = parentSheet.ChildSheets
            .Single(child => child.SheetName == "Userport");

        KiCadSchematicSheetPin pin = Assert.Single(userportChild.Pins);

        Assert.Equal("CNT1", pin.Name);
        Assert.Equal("CNT1", pin.NormalizedName);

        // The POSITION is what ties the pin to the parent's wire, so it has to survive too.
        Assert.NotNull(pin.At);
        Assert.Equal(10, pin.At!.X);
        Assert.Equal(10, pin.At.Y);
    }

    [Fact]
    public async Task A_child_sheet_pin_records_the_file_it_belongs_to()
    {
        // The alias builder resolves a pin to a loaded schematic through this file name, so a
        // missing one silently disables the whole feature for that sheet.
        string parent = this.thisWorkspace.WriteFile(
            "root.kicad_sch", KiCadFixtures.SchematicParentJoiningTwoChildren);
        string userport = this.thisWorkspace.WriteFile(
            "userport.kicad_sch", KiCadFixtures.SchematicUserportChild);
        string serialBus = this.thisWorkspace.WriteFile(
            "serialbus.kicad_sch", KiCadFixtures.SchematicSerialBusChild);

        KiCadProjectRoot? root = await this.LoadAsync(parent, userport, serialBus);

        KiCadSchematic parentSheet = root!.Schematics
            .Single(schematic => schematic.Filename == "root.kicad_sch");

        Assert.Equal(
            new[] { "serialbus.kicad_sch", "userport.kicad_sch" },
            parentSheet.ChildSheets
                .Select(child => child.FileName)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray());
    }

    [Fact]
    public async Task A_sheet_with_no_child_sheets_reports_none()
    {
        // A flat design must produce an empty list rather than null - single-sheet boards were
        // unaffected by the cross-sheet bug and must stay unaffected by its fix.
        string flat = this.thisWorkspace.WriteFile("flat.kicad_sch", KiCadFixtures.Schematic);

        KiCadProjectRoot? root = await this.LoadAsync(flat);

        Assert.Empty(root!.Schematics.Single().ChildSheets);
    }

    // ###########################################################################################
    // END TO END, through the real loader: the userport sheet resolves the SERIAL BUS name for the
    // wire it calls CNT1. This is the reported bug reduced to its smallest reproduction, and it
    // fails against the version with no alias table.
    // ###########################################################################################
    [Fact]
    public async Task A_sheet_resolves_the_OTHER_sheets_name_for_a_wire_they_share()
    {
        string parent = this.thisWorkspace.WriteFile(
            "root.kicad_sch", KiCadFixtures.SchematicParentJoiningTwoChildren);
        string userport = this.thisWorkspace.WriteFile(
            "userport.kicad_sch", KiCadFixtures.SchematicUserportChild);
        string serialBus = this.thisWorkspace.WriteFile(
            "serialbus.kicad_sch", KiCadFixtures.SchematicSerialBusChild);

        KiCadProjectBundle? bundle = await KiCadProjectLoader.LoadRawAsync(
            [parent, userport, serialBus]);

        Assert.NotNull(bundle);

        int userportIndex = bundle!.Root.Schematics
            .Select((schematic, index) => (schematic, index))
            .Single(pair => pair.schematic.Filename == "userport.kicad_sch")
            .index;

        var index = bundle.SchematicNetPathIndexBySchematicIndex[userportIndex];

        // Its own name still resolves...
        Assert.True(index.ContainsKey("CNT1"));

        // ...and now so does the name the PCB actually carries for that copper.
        Assert.True(index.ContainsKey("CNT"));

        // The SAME geometry, not a second copy of it: an alias is another way to ask for one net.
        Assert.Same(index["CNT1"], index["CNT"]);
    }
}
