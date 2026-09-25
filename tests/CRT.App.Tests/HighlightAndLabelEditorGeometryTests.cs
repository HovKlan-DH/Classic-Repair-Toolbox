using Avalonia;
using Avalonia.Input;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for the last two Tier B extractions.
//
// HighlightRectBuilder decides which component rectangles appear on a schematic for the active
// region (PAL / NTSC) - get it wrong and a user sees components that do not exist on their board.
// LabelEditorGeometry decides where the resize handles are and what the keyboard chords do.
public class HighlightRectBuilderTests
{
    private static BoardData Board(
        (string Label, string Region)[] components,
        (string Schematic, string Label, string X, string Y, string W, string H)[] highlights)
    {
        var data = new BoardData();

        foreach (var (label, region) in components)
        {
            data.Components.Add(new ComponentEntry { BoardLabel = label, Region = region });
        }

        foreach (var (schematic, label, x, y, w, h) in highlights)
        {
            data.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = schematic, BoardLabel = label, X = x, Y = y, Width = w, Height = h
            });
        }

        return data;
    }

    [Fact]
    public void A_highlight_is_grouped_by_schematic_then_board_label()
    {
        BoardData board = Board(
            new[] { ("U1", "") },
            new[] { ("Sheet 1", "U1", "10", "20", "30", "40") });

        var rects = HighlightRectBuilder.BuildHighlightRects(board, "PAL");

        Rect rect = Assert.Single(rects["Sheet 1"]["U1"]);
        Assert.Equal(new Rect(10, 20, 30, 40), rect);
    }

    [Fact]
    public void A_component_with_no_declared_region_is_visible_in_every_region()
    {
        BoardData board = Board(
            new[] { ("U1", "") },
            new[] { ("Sheet 1", "U1", "1", "1", "5", "5") });

        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]);
        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "NTSC")["Sheet 1"]);
    }

    [Fact]
    public void A_component_declared_for_one_region_is_hidden_in_the_other()
    {
        BoardData board = Board(
            new[] { ("U1", "PAL") },
            new[] { ("Sheet 1", "U1", "1", "1", "5", "5") });

        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]);
        Assert.Empty(HighlightRectBuilder.BuildHighlightRects(board, "NTSC"));
    }

    [Fact]
    public void A_component_listed_for_both_regions_is_visible_in_both()
    {
        BoardData board = Board(
            new[] { ("U1", "PAL"), ("U1", "NTSC") },
            new[] { ("Sheet 1", "U1", "1", "1", "5", "5") });

        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]);
        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "NTSC")["Sheet 1"]);
    }

    [Fact]
    public void Region_matching_ignores_case()
    {
        BoardData board = Board(
            new[] { ("U1", "pal") },
            new[] { ("Sheet 1", "U1", "1", "1", "5", "5") });

        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]);
    }

    [Fact]
    public void A_highlight_for_a_label_that_has_no_component_row_is_still_shown()
    {
        // The label editor can create a highlight before its Components row exists.
        BoardData board = Board(
            Array.Empty<(string, string)>(),
            new[] { ("Sheet 1", "ORPHAN", "1", "1", "5", "5") });

        Assert.Single(HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]);
    }

    [Fact]
    public void Several_rectangles_for_one_label_are_all_kept()
    {
        BoardData board = Board(
            new[] { ("U1", "") },
            new[]
            {
                ("Sheet 1", "U1", "1", "1", "5", "5"),
                ("Sheet 1", "U1", "50", "50", "5", "5")
            });

        Assert.Equal(2, HighlightRectBuilder.BuildHighlightRects(board, "PAL")["Sheet 1"]["U1"].Count);
    }

    [Theory]
    [InlineData("", "1", "5", "5")]        // blank X
    [InlineData("1", "1", "0", "5")]       // zero width
    [InlineData("1", "1", "5", "-1")]      // negative height
    [InlineData("1", "1", "abc", "5")]     // unparseable
    [InlineData("1,5", "1", "5", "5")]     // comma decimal separator
    public void A_malformed_or_degenerate_highlight_is_dropped(string x, string y, string w, string h)
    {
        BoardData board = Board(
            new[] { ("U1", "") },
            new[] { ("Sheet 1", "U1", x, y, w, h) });

        Assert.Empty(HighlightRectBuilder.BuildHighlightRects(board, "PAL"));
    }

    [Fact]
    public void A_highlight_with_no_schematic_or_label_is_dropped()
    {
        BoardData board = Board(
            new[] { ("U1", "") },
            new[]
            {
                ("", "U1", "1", "1", "5", "5"),
                ("Sheet 1", "  ", "1", "1", "5", "5")
            });

        Assert.Empty(HighlightRectBuilder.BuildHighlightRects(board, "PAL"));
    }

    [Fact]
    public void A_board_with_no_highlights_yields_an_empty_lookup()
    {
        Assert.Empty(HighlightRectBuilder.BuildHighlightRects(new BoardData(), "PAL"));
    }
}

public class LabelEditorGeometryTests
{
    // ------------------------------------------------- IsLabelEditorRectangleTooSmall

    [Theory]
    [InlineData(20, 20)]
    [InlineData(15, 15)]     // exactly at the minimum
    [InlineData(100, 15)]
    public void A_usable_rectangle_is_accepted(double width, double height)
    {
        Assert.False(LabelEditorGeometry.IsLabelEditorRectangleTooSmall(new Rect(0, 0, width, height)));
    }

    [Theory]
    [InlineData(14, 100)]    // too narrow, however tall
    [InlineData(100, 14)]    // too short, however wide
    [InlineData(0, 0)]
    [InlineData(5, 5)]
    public void A_rectangle_below_the_minimum_on_either_axis_is_rejected(double width, double height)
    {
        // An accidental click would otherwise create a highlight nobody can grab again.
        Assert.True(LabelEditorGeometry.IsLabelEditorRectangleTooSmall(new Rect(0, 0, width, height)));
    }

    // ------------------------------------- TryGetKeyboardLabelEditorResizeDragMode

    // LabelEditorDragMode is internal, so these cannot be [Theory] parameters on a public method.
    private static LabelEditorDragMode ResizeModeFor(Key key, KeyModifiers modifiers)
    {
        Assert.True(LabelEditorGeometry.TryGetKeyboardLabelEditorResizeDragMode(
            key, modifiers, out LabelEditorDragMode mode));
        return mode;
    }

    [Fact]
    public void Shift_plus_an_arrow_grows_the_edge_the_arrow_points_at()
    {
        Assert.Equal(LabelEditorDragMode.ResizeLeft, ResizeModeFor(Key.Left, KeyModifiers.Shift));
        Assert.Equal(LabelEditorDragMode.ResizeRight, ResizeModeFor(Key.Right, KeyModifiers.Shift));
        Assert.Equal(LabelEditorDragMode.ResizeTop, ResizeModeFor(Key.Up, KeyModifiers.Shift));
        Assert.Equal(LabelEditorDragMode.ResizeBottom, ResizeModeFor(Key.Down, KeyModifiers.Shift));
    }

    [Fact]
    public void Alt_plus_an_arrow_moves_the_opposite_edge()
    {
        Assert.Equal(LabelEditorDragMode.ResizeRight, ResizeModeFor(Key.Left, KeyModifiers.Alt));
        Assert.Equal(LabelEditorDragMode.ResizeLeft, ResizeModeFor(Key.Right, KeyModifiers.Alt));
        Assert.Equal(LabelEditorDragMode.ResizeBottom, ResizeModeFor(Key.Up, KeyModifiers.Alt));
        Assert.Equal(LabelEditorDragMode.ResizeTop, ResizeModeFor(Key.Down, KeyModifiers.Alt));
    }

    [Theory]
    [InlineData(KeyModifiers.None)]
    [InlineData(KeyModifiers.Control)]
    public void An_arrow_without_shift_or_alt_is_not_a_resize(KeyModifiers modifiers)
    {
        // Plain arrows nudge the whole highlight instead.
        Assert.False(LabelEditorGeometry.TryGetKeyboardLabelEditorResizeDragMode(
            Key.Left, modifiers, out _));
    }

    [Fact]
    public void Shift_and_alt_together_are_ambiguous_and_rejected()
    {
        Assert.False(LabelEditorGeometry.TryGetKeyboardLabelEditorResizeDragMode(
            Key.Left, KeyModifiers.Shift | KeyModifiers.Alt, out _));
    }

    [Fact]
    public void A_non_arrow_key_is_not_a_resize()
    {
        Assert.False(LabelEditorGeometry.TryGetKeyboardLabelEditorResizeDragMode(
            Key.A, KeyModifiers.Shift, out _));
    }

    // ------------------------------------------- BuildLabelEditorHandleHitRects

    [Fact]
    public void A_large_selection_gets_all_eight_handles()
    {
        var handles = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 200, 200), scale: 1.0);

        Assert.Equal(8, handles.Count);
        Assert.Equal(8, handles.Select(h => h.DragMode).Distinct().Count());
    }

    [Fact]
    public void The_four_corner_handles_always_exist()
    {
        // Corner drags resize two axes at once, so they take priority when space is tight.
        var handles = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 4, 4), scale: 1.0);

        var modes = handles.Select(h => h.DragMode).ToList();

        Assert.Contains(LabelEditorDragMode.ResizeTopLeft, modes);
        Assert.Contains(LabelEditorDragMode.ResizeTopRight, modes);
        Assert.Contains(LabelEditorDragMode.ResizeBottomRight, modes);
        Assert.Contains(LabelEditorDragMode.ResizeBottomLeft, modes);
    }

    [Fact]
    public void A_selection_too_narrow_for_side_handles_drops_them_rather_than_overlapping()
    {
        var handles = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 4, 200), scale: 1.0);

        var modes = handles.Select(h => h.DragMode).ToList();

        Assert.DoesNotContain(LabelEditorDragMode.ResizeTop, modes);
        Assert.DoesNotContain(LabelEditorDragMode.ResizeBottom, modes);
        Assert.Contains(LabelEditorDragMode.ResizeLeft, modes);    // vertical sides still fit
    }

    [Fact]
    public void Handles_grow_in_world_terms_as_the_view_zooms_out()
    {
        // Handles are sized in screen pixels, so a zoomed-out view needs bigger world rects.
        var zoomedIn = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 200, 200), scale: 4.0);
        var zoomedOut = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 200, 200), scale: 0.5);

        double inSize = zoomedIn.First(h => h.DragMode == LabelEditorDragMode.ResizeTopLeft).HitRect.Width;
        double outSize = zoomedOut.First(h => h.DragMode == LabelEditorDragMode.ResizeTopLeft).HitRect.Width;

        Assert.True(outSize > inSize);
    }

    [Fact]
    public void Handle_size_is_clamped_so_it_never_becomes_absurd()
    {
        var extreme = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(0, 0, 200, 200), scale: 0.001);

        double size = extreme.First(h => h.DragMode == LabelEditorDragMode.ResizeTopLeft).HitRect.Width;

        Assert.InRange(size, 4.0, 12.0);
    }

    [Fact]
    public void Corner_handles_are_centred_on_the_corners()
    {
        var handles = LabelEditorGeometry.BuildLabelEditorHandleHitRects(new Rect(100, 100, 200, 200), scale: 1.0);

        Rect topLeft = handles.First(h => h.DragMode == LabelEditorDragMode.ResizeTopLeft).HitRect;

        Assert.Equal(100, topLeft.Center.X, precision: 6);
        Assert.Equal(100, topLeft.Center.Y, precision: 6);
    }

    // ------------------------------------------------------------------------ ResizeRect

    // Each handle moves its own edge and leaves the opposite one alone - the property that makes a
    // resize feel like grabbing an edge rather than moving the whole shape.
    [Theory]
    [InlineData(LabelEditorDragMode.ResizeLeft, 10, 0, 110, 100, 90, 100)]
    [InlineData(LabelEditorDragMode.ResizeRight, 10, 0, 100, 100, 110, 100)]
    [InlineData(LabelEditorDragMode.ResizeTop, 0, 10, 100, 110, 100, 90)]
    [InlineData(LabelEditorDragMode.ResizeBottom, 0, 10, 100, 100, 100, 110)]
    internal void A_side_handle_moves_only_its_own_edge(
        LabelEditorDragMode mode, double dx, double dy,
        double expectedX, double expectedY, double expectedWidth, double expectedHeight)
    {
        var result = LabelEditorGeometry.ResizeRect(new Rect(100, 100, 100, 100), mode, dx, dy, 1.0);

        Assert.Equal(expectedX, result.X, 3);
        Assert.Equal(expectedY, result.Y, 3);
        Assert.Equal(expectedWidth, result.Width, 3);
        Assert.Equal(expectedHeight, result.Height, 3);
    }

    // A corner moves both of its edges.
    [Fact]
    public void A_corner_handle_moves_both_of_its_edges()
    {
        var result = LabelEditorGeometry.ResizeRect(
            new Rect(100, 100, 100, 100), LabelEditorDragMode.ResizeBottomRight, 20, 30, 1.0);

        Assert.Equal(new Rect(100, 100, 120, 130), result);
    }

    // Dragging an edge past its opposite must pin at the minimum, never invert. A negative width
    // is not "a small rectangle" - Rect mirrors it about the origin, so the area jumps to the wrong
    // side of the board and the user loses it.
    [Theory]
    [InlineData(LabelEditorDragMode.ResizeLeft, 500.0, 0.0)]
    [InlineData(LabelEditorDragMode.ResizeRight, -500.0, 0.0)]
    [InlineData(LabelEditorDragMode.ResizeTop, 0.0, 500.0)]
    [InlineData(LabelEditorDragMode.ResizeBottom, 0.0, -500.0)]
    [InlineData(LabelEditorDragMode.ResizeTopLeft, 500.0, 500.0)]
    [InlineData(LabelEditorDragMode.ResizeBottomRight, -500.0, -500.0)]
    internal void A_resize_can_never_invert_the_rectangle(LabelEditorDragMode mode, double dx, double dy)
    {
        var result = LabelEditorGeometry.ResizeRect(new Rect(100, 100, 100, 100), mode, dx, dy, 5.0);

        Assert.True(result.Width >= 5.0, $"width collapsed to {result.Width}");
        Assert.True(result.Height >= 5.0, $"height collapsed to {result.Height}");
    }

    // The edge that was NOT being dragged stays anchored even when the drag is clamped, so the
    // rectangle shrinks towards the fixed edge rather than sliding across the board.
    [Fact]
    public void A_clamped_resize_keeps_the_opposite_edge_anchored()
    {
        var result = LabelEditorGeometry.ResizeRect(
            new Rect(100, 100, 100, 100), LabelEditorDragMode.ResizeLeft, 500, 0, 5.0);

        // Dragging the left edge far right pins it 5px from the right edge, which has not moved.
        Assert.Equal(200.0, result.Right, 3);
        Assert.Equal(195.0, result.Left, 3);
    }

    // Move shifts the whole rectangle and changes neither dimension.
    [Fact]
    public void Move_translates_without_resizing()
    {
        var result = LabelEditorGeometry.ResizeRect(
            new Rect(100, 100, 60, 40), LabelEditorDragMode.Move, 25, -15, 1.0);

        Assert.Equal(new Rect(125, 85, 60, 40), result);
    }

    [Fact]
    public void An_unset_drag_mode_leaves_the_rectangle_untouched()
    {
        var original = new Rect(10, 20, 30, 40);

        Assert.Equal(original, LabelEditorGeometry.ResizeRect(original, LabelEditorDragMode.None, 99, 99, 1.0));
    }

    // ------------------------------------------------------------------ ClampRectToBounds

    // A worklog area dragged off the board would be saved at coordinates outside the bitmap, where
    // it cannot be seen or grabbed again - so it is pushed back inside instead.
    [Theory]
    [InlineData(-50.0, -50.0, 0.0, 0.0)]
    [InlineData(950.0, 550.0, 900.0, 500.0)]
    public void A_rectangle_outside_the_image_is_pushed_back_inside(
        double x, double y, double expectedX, double expectedY)
    {
        var result = LabelEditorGeometry.ClampRectToBounds(new Rect(x, y, 100, 100), new Size(1000, 600));

        Assert.Equal(expectedX, result.X, 3);
        Assert.Equal(expectedY, result.Y, 3);
        Assert.Equal(100.0, result.Width, 3);
        Assert.Equal(100.0, result.Height, 3);
    }

    // A rectangle already inside is returned untouched - clamping must not nudge a valid area.
    [Fact]
    public void A_rectangle_inside_the_image_is_left_alone()
    {
        var original = new Rect(100, 100, 200, 150);

        Assert.Equal(original, LabelEditorGeometry.ClampRectToBounds(original, new Size(1000, 600)));
    }

    // An area larger than the image is capped rather than left overhanging.
    [Fact]
    public void A_rectangle_larger_than_the_image_is_capped_to_it()
    {
        var result = LabelEditorGeometry.ClampRectToBounds(new Rect(-10, -10, 5000, 5000), new Size(1000, 600));

        Assert.Equal(new Rect(0, 0, 1000, 600), result);
    }

    // A zero-sized image cannot be clamped against, so the input is returned rather than collapsed.
    [Fact]
    public void An_unknown_image_size_leaves_the_rectangle_alone()
    {
        var original = new Rect(10, 20, 30, 40);

        Assert.Equal(original, LabelEditorGeometry.ClampRectToBounds(original, new Size(0, 0)));
    }

    // ------------------------------------------------------- ClampResizedRectToBounds

    // The whole reason this exists separately from ClampRectToBounds: a RESIZE clamped at the board
    // edge must trim the edge that strayed out and leave every other edge alone. The translate-clamp
    // preserved width instead, so dragging the LEFT edge out past x=0 slid the rectangle back and
    // pushed the RIGHT edge outward - the user dragged one handle and the opposite edge moved.
    [Fact]
    public void A_resize_clamped_at_the_left_edge_keeps_the_right_edge_anchored()
    {
        // Left edge dragged to -40; right edge is at 200 and must stay there.
        var result = LabelEditorGeometry.ClampResizedRectToBounds(
            new Rect(-40, 100, 240, 100), new Size(1000, 600), 8.0);

        Assert.Equal(0.0, result.Left, 3);
        Assert.Equal(200.0, result.Right, 3);
    }

    [Fact]
    public void A_resize_clamped_at_the_right_edge_keeps_the_left_edge_anchored()
    {
        // Right edge dragged to 1100 against a 1000-wide image; left stays at 900.
        var result = LabelEditorGeometry.ClampResizedRectToBounds(
            new Rect(900, 100, 200, 100), new Size(1000, 600), 8.0);

        Assert.Equal(900.0, result.Left, 3);
        Assert.Equal(1000.0, result.Right, 3);
    }

    [Fact]
    public void A_resize_clamped_at_the_top_keeps_the_bottom_anchored()
    {
        var result = LabelEditorGeometry.ClampResizedRectToBounds(
            new Rect(100, -30, 100, 230), new Size(1000, 600), 8.0);

        Assert.Equal(0.0, result.Top, 3);
        Assert.Equal(200.0, result.Bottom, 3);
    }

    // A rectangle already inside must come back untouched - clamping cannot nudge a valid area.
    [Fact]
    public void A_resized_rectangle_inside_the_image_is_left_alone()
    {
        var original = new Rect(100, 100, 200, 150);

        Assert.Equal(original, LabelEditorGeometry.ClampResizedRectToBounds(original, new Size(1000, 600), 8.0));
    }

    // The minimum is still honoured at the boundary: an edge dragged far outside collapses against
    // it rather than inverting.
    [Fact]
    public void A_resize_pinned_against_the_edge_still_respects_the_minimum()
    {
        var result = LabelEditorGeometry.ClampResizedRectToBounds(
            new Rect(-500, 100, 502, 100), new Size(1000, 600), 8.0);

        Assert.True(result.Width >= 8.0, $"width collapsed to {result.Width}");
        Assert.True(result.Left >= -0.001, $"left escaped the image at {result.Left}");
    }

    // A zero-sized image gives nothing to clamp against, so the input is returned rather than
    // collapsed to nothing.
    [Fact]
    public void An_unknown_image_size_leaves_a_resized_rectangle_alone()
    {
        var original = new Rect(10, 20, 30, 40);

        Assert.Equal(original, LabelEditorGeometry.ClampResizedRectToBounds(original, new Size(0, 0), 8.0));
    }
}
