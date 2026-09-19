using Avalonia.Input;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The KiCad trace calibration maths: handle remapping across flips, keyboard nudging, pointer
// drags, and the mirror-flag swap used when entering the mode.
//
// This logic sat inside TabSchematics at 0% coverage until it was extracted. It is worth real
// tests rather than a smoke pass for one reason: the flip remap is an eight-handle by
// four-flip-state table, and a wrong arm in it is INVISIBLE. Nothing throws, nothing looks
// broken - dragging a corner of a mirrored board just resizes the wrong edge, and only someone
// calibrating a flipped board by hand would ever notice. So the table is asserted entry by
// entry below rather than sampled.
//
// The one rule to hold on to while reading: mirroring is not a flag stored beside the edges,
// it IS the edge ordering. Left > Right means mirrored horizontally. That is why several tests
// build deliberately "backwards" boxes.
// ###########################################################################################
// A note on why the [Theory] cases pass drag modes as STRINGS rather than as the enum itself:
// LabelEditorDragMode is internal to the app assembly, xUnit requires test classes to be public
// (xUnit1000), and a public method cannot take an internal parameter type (CS0051). Naming the
// member and resolving it through Mode() below satisfies both, and reads at least as clearly in
// the test list as the enum would.
public class KiCadCalibrationGeometryTests
{
    // Resolves a drag-mode name used in [InlineData] to the internal enum value.
    private static LabelEditorDragMode Mode(string name) =>
        Enum.Parse<LabelEditorDragMode>(name);

    // An unmirrored box: edges ascending, so IsMirroredX/Y are both false.
    private static KiCadCalibrationBox Box() => new(10, 20, 110, 220);

    // Mirrored horizontally - Left and Right swapped, nothing else.
    private static KiCadCalibrationBox MirroredX() => new(110, 20, 10, 220);

    private static KiCadCalibrationBox MirroredY() => new(10, 220, 110, 20);

    private static KiCadCalibrationBox MirroredBoth() => new(110, 220, 10, 20);

    // -----------------------------------------------------------------------------------------
    // The box itself
    // -----------------------------------------------------------------------------------------

    // Mirroring is derived from the edge order, never stored - so the two cannot drift apart.
    [Fact]
    public void Mirroring_is_read_from_the_edge_order()
    {
        Assert.False(Box().IsMirroredX);
        Assert.False(Box().IsMirroredY);

        Assert.True(MirroredX().IsMirroredX);
        Assert.False(MirroredX().IsMirroredY);

        Assert.False(MirroredY().IsMirroredX);
        Assert.True(MirroredY().IsMirroredY);

        Assert.True(MirroredBoth().IsMirroredX);
        Assert.True(MirroredBoth().IsMirroredY);
    }

    // The normalised view is the same rectangle with its edges put back in ascending order,
    // which is the space the arithmetic works in. It must not depend on how the box is stored.
    [Fact]
    public void The_normalised_edges_are_the_same_rectangle_however_it_is_mirrored()
    {
        foreach (var box in new[] { Box(), MirroredX(), MirroredY(), MirroredBoth() })
        {
            Assert.Equal(10, box.NormalisedLeft);
            Assert.Equal(20, box.NormalisedTop);
            Assert.Equal(110, box.NormalisedRight);
            Assert.Equal(220, box.NormalisedBottom);
        }
    }

    // Coming back OUT of normalised space has to restore the mirroring, or a nudge silently
    // un-flips the board.
    [Fact]
    public void Rebuilding_from_normalised_edges_restores_the_mirroring()
    {
        var rebuilt = MirroredBoth().WithNormalisedEdges(1, 2, 3, 4);

        Assert.Equal(3, rebuilt.Left);
        Assert.Equal(4, rebuilt.Top);
        Assert.Equal(1, rebuilt.Right);
        Assert.Equal(2, rebuilt.Bottom);
        Assert.True(rebuilt.IsMirroredX);
        Assert.True(rebuilt.IsMirroredY);
    }

    // -----------------------------------------------------------------------------------------
    // Handle remapping across flips - the table this whole file exists for
    // -----------------------------------------------------------------------------------------

    // An unmirrored box changes nothing, for every handle. This is the common case and it must
    // be a pure pass-through.
    [Theory]
    [InlineData("ResizeTopLeft")]
    [InlineData("ResizeTop")]
    [InlineData("ResizeTopRight")]
    [InlineData("ResizeRight")]
    [InlineData("ResizeBottomRight")]
    [InlineData("ResizeBottom")]
    [InlineData("ResizeBottomLeft")]
    [InlineData("ResizeLeft")]
    [InlineData("Move")]
    [InlineData("None")]
    public void An_unmirrored_box_remaps_nothing(string dragModeName)
    {
        var dragMode = Mode(dragModeName);

        Assert.Equal(dragMode, KiCadCalibrationGeometry.RemapDragModeForFlip(Box(), dragMode));
    }

    // Mirrored horizontally: every handle that names a LEFT or RIGHT edge swaps to the other
    // side; the pure vertical handles (Top, Bottom) are untouched, because the flip is not on
    // their axis.
    [Theory]
    [InlineData("ResizeTopLeft", "ResizeTopRight")]
    [InlineData("ResizeTopRight", "ResizeTopLeft")]
    [InlineData("ResizeBottomLeft", "ResizeBottomRight")]
    [InlineData("ResizeBottomRight", "ResizeBottomLeft")]
    [InlineData("ResizeLeft", "ResizeRight")]
    [InlineData("ResizeRight", "ResizeLeft")]
    [InlineData("ResizeTop", "ResizeTop")]
    [InlineData("ResizeBottom", "ResizeBottom")]
    public void A_horizontally_mirrored_box_swaps_only_the_left_and_right_halves(
        string grabbed,
        string expected)
    {
        Assert.Equal(Mode(expected), KiCadCalibrationGeometry.RemapDragModeForFlip(MirroredX(), Mode(grabbed)));
    }

    // The mirror image of the test above: only the TOP/BOTTOM halves move.
    [Theory]
    [InlineData("ResizeTopLeft", "ResizeBottomLeft")]
    [InlineData("ResizeBottomLeft", "ResizeTopLeft")]
    [InlineData("ResizeTopRight", "ResizeBottomRight")]
    [InlineData("ResizeBottomRight", "ResizeTopRight")]
    [InlineData("ResizeTop", "ResizeBottom")]
    [InlineData("ResizeBottom", "ResizeTop")]
    [InlineData("ResizeLeft", "ResizeLeft")]
    [InlineData("ResizeRight", "ResizeRight")]
    public void A_vertically_mirrored_box_swaps_only_the_top_and_bottom_halves(
        string grabbed,
        string expected)
    {
        Assert.Equal(Mode(expected), KiCadCalibrationGeometry.RemapDragModeForFlip(MirroredY(), Mode(grabbed)));
    }

    // Flipped on both axes, every handle maps to its diagonal opposite - corners included.
    [Theory]
    [InlineData("ResizeTopLeft", "ResizeBottomRight")]
    [InlineData("ResizeBottomRight", "ResizeTopLeft")]
    [InlineData("ResizeTopRight", "ResizeBottomLeft")]
    [InlineData("ResizeBottomLeft", "ResizeTopRight")]
    [InlineData("ResizeTop", "ResizeBottom")]
    [InlineData("ResizeBottom", "ResizeTop")]
    [InlineData("ResizeLeft", "ResizeRight")]
    [InlineData("ResizeRight", "ResizeLeft")]
    public void A_doubly_mirrored_box_maps_every_handle_to_its_diagonal_opposite(
        string grabbed,
        string expected)
    {
        Assert.Equal(Mode(expected), KiCadCalibrationGeometry.RemapDragModeForFlip(MirroredBoth(), Mode(grabbed)));
    }

    // Remapping twice returns the original handle, on every flip state. An involution is a
    // property the table must have and a single wrong arm would break - it catches a mistake
    // the case-by-case tests above could only catch if I typed that exact pair correctly.
    [Fact]
    public void Remapping_a_handle_twice_returns_the_original_handle()
    {
        LabelEditorDragMode[] handles =
        {
            LabelEditorDragMode.ResizeTopLeft,
            LabelEditorDragMode.ResizeTop,
            LabelEditorDragMode.ResizeTopRight,
            LabelEditorDragMode.ResizeRight,
            LabelEditorDragMode.ResizeBottomRight,
            LabelEditorDragMode.ResizeBottom,
            LabelEditorDragMode.ResizeBottomLeft,
            LabelEditorDragMode.ResizeLeft,
        };

        foreach (var box in new[] { Box(), MirroredX(), MirroredY(), MirroredBoth() })
        {
            foreach (var handle in handles)
            {
                var once = KiCadCalibrationGeometry.RemapDragModeForFlip(box, handle);
                var twice = KiCadCalibrationGeometry.RemapDragModeForFlip(box, once);

                Assert.Equal(handle, twice);
            }
        }
    }

    // Move is not a resize, so no flip touches it. The tab relies on this - it passes Move
    // straight through without remapping, and the two must agree.
    [Fact]
    public void Move_is_never_remapped_whatever_the_flip()
    {
        foreach (var box in new[] { Box(), MirroredX(), MirroredY(), MirroredBoth() })
        {
            Assert.Equal(
                LabelEditorDragMode.Move,
                KiCadCalibrationGeometry.RemapDragModeForFlip(box, LabelEditorDragMode.Move));
        }
    }

    // -----------------------------------------------------------------------------------------
    // Keyboard nudging
    // -----------------------------------------------------------------------------------------

    // A bare arrow key moves the whole box one pixel, leaving its size alone.
    [Theory]
    [InlineData(Key.Left, -1, 0)]
    [InlineData(Key.Right, 1, 0)]
    [InlineData(Key.Up, 0, -1)]
    [InlineData(Key.Down, 0, 1)]
    public void A_bare_arrow_key_moves_the_box_without_resizing_it(Key key, double dx, double dy)
    {
        Assert.True(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.None, out var moved));

        Assert.Equal(10 + dx, moved.Left);
        Assert.Equal(110 + dx, moved.Right);
        Assert.Equal(20 + dy, moved.Top);
        Assert.Equal(220 + dy, moved.Bottom);
    }

    // Shift EXPANDS in the direction pressed: only the edge on that side moves outward.
    [Theory]
    [InlineData(Key.Left, 9, 20, 110, 220)]
    [InlineData(Key.Right, 10, 20, 111, 220)]
    [InlineData(Key.Up, 10, 19, 110, 220)]
    [InlineData(Key.Down, 10, 20, 110, 221)]
    public void Shift_expands_the_box_on_the_side_pressed(
        Key key,
        double left,
        double top,
        double right,
        double bottom)
    {
        Assert.True(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.Shift, out var grown));

        Assert.Equal(left, grown.Left);
        Assert.Equal(top, grown.Top);
        Assert.Equal(right, grown.Right);
        Assert.Equal(bottom, grown.Bottom);
    }

    // Alt SHRINKS from the OPPOSITE side to the one pressed - pressing Left pulls the right edge
    // in. That is deliberately not the mirror of Shift, and it matches the label editor.
    [Theory]
    [InlineData(Key.Left, 10, 20, 109, 220)]
    [InlineData(Key.Right, 11, 20, 110, 220)]
    [InlineData(Key.Up, 10, 20, 110, 219)]
    [InlineData(Key.Down, 10, 21, 110, 220)]
    public void Alt_shrinks_the_box_from_the_side_opposite_the_one_pressed(
        Key key,
        double left,
        double top,
        double right,
        double bottom)
    {
        Assert.True(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.Alt, out var shrunk));

        Assert.Equal(left, shrunk.Left);
        Assert.Equal(top, shrunk.Top);
        Assert.Equal(right, shrunk.Right);
        Assert.Equal(bottom, shrunk.Bottom);
    }

    // The guard that stops Alt collapsing the box to nothing. A box exactly one step wide cannot
    // be shrunk any further horizontally - if it could, it would reach zero width and there
    // would be nothing left on screen to grab.
    [Theory]
    [InlineData(Key.Left)]
    [InlineData(Key.Right)]
    public void Alt_refuses_to_shrink_a_box_that_is_already_one_step_wide(Key key)
    {
        var thin = new KiCadCalibrationBox(10, 20, 11, 220);

        Assert.False(KiCadCalibrationGeometry.TryApplyKeyboardStep(thin, key, KeyModifiers.Alt, out var unchanged));

        // The box must come back untouched, not half-applied.
        Assert.Equal(10, unchanged.Left);
        Assert.Equal(11, unchanged.Right);
    }

    [Theory]
    [InlineData(Key.Up)]
    [InlineData(Key.Down)]
    public void Alt_refuses_to_shrink_a_box_that_is_already_one_step_tall(Key key)
    {
        var flat = new KiCadCalibrationBox(10, 20, 110, 21);

        Assert.False(KiCadCalibrationGeometry.TryApplyKeyboardStep(flat, key, KeyModifiers.Alt, out var unchanged));

        Assert.Equal(20, unchanged.Top);
        Assert.Equal(21, unchanged.Bottom);
    }

    // Shift and Alt together would mean "expand and shrink at once". Rather than silently
    // picking one, the press is refused outright.
    [Theory]
    [InlineData(Key.Left)]
    [InlineData(Key.Right)]
    [InlineData(Key.Up)]
    [InlineData(Key.Down)]
    public void Shift_and_alt_together_are_refused(Key key)
    {
        Assert.False(
            KiCadCalibrationGeometry.TryApplyKeyboardStep(
                Box(),
                key,
                KeyModifiers.Shift | KeyModifiers.Alt,
                out _));
    }

    // Anything that is not an arrow key does nothing, in every modifier state. The tab uses the
    // false return to leave the key unhandled so it can reach whatever else wants it.
    [Theory]
    [InlineData(Key.Enter)]
    [InlineData(Key.Escape)]
    [InlineData(Key.A)]
    [InlineData(Key.Space)]
    public void A_non_arrow_key_changes_nothing(Key key)
    {
        Assert.False(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.None, out _));
        Assert.False(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.Shift, out _));
        Assert.False(KiCadCalibrationGeometry.TryApplyKeyboardStep(Box(), key, KeyModifiers.Alt, out _));
    }

    // Nudging a MIRRORED box has to move it the way the user sees it on screen, and leave it
    // still mirrored afterwards. This is the case the normalise/rebuild round-trip exists for:
    // work out the answer in ascending edges, then put the inversion back.
    [Fact]
    public void Nudging_a_mirrored_box_moves_it_on_screen_and_keeps_it_mirrored()
    {
        Assert.True(
            KiCadCalibrationGeometry.TryApplyKeyboardStep(
                MirroredBoth(),
                Key.Right,
                KeyModifiers.None,
                out var moved));

        // Still mirrored - the flip survived the round-trip.
        Assert.True(moved.IsMirroredX);
        Assert.True(moved.IsMirroredY);

        // And it genuinely moved right: both horizontal edges are one greater than before.
        Assert.Equal(11, moved.NormalisedLeft);
        Assert.Equal(111, moved.NormalisedRight);

        // Vertically untouched.
        Assert.Equal(20, moved.NormalisedTop);
        Assert.Equal(220, moved.NormalisedBottom);
    }

    // Shift-expanding a mirrored box expands the edge the user sees on that side, not the
    // stored field of the same name. Pressing Left must always grow the box leftward on screen.
    [Fact]
    public void Shift_expanding_a_mirrored_box_grows_the_side_the_user_sees()
    {
        Assert.True(
            KiCadCalibrationGeometry.TryApplyKeyboardStep(
                MirroredX(),
                Key.Left,
                KeyModifiers.Shift,
                out var grown));

        Assert.Equal(9, grown.NormalisedLeft);
        Assert.Equal(110, grown.NormalisedRight);
        Assert.True(grown.IsMirroredX);
    }

    // -----------------------------------------------------------------------------------------
    // Pointer drags
    // -----------------------------------------------------------------------------------------

    // Move shifts all four edges by the same offset, so the size is preserved exactly.
    [Fact]
    public void A_move_drag_shifts_every_edge_and_preserves_the_size()
    {
        var dragged = KiCadCalibrationGeometry.ApplyDrag(Box(), LabelEditorDragMode.Move, 5, -7);

        Assert.Equal(15, dragged.Left);
        Assert.Equal(115, dragged.Right);
        Assert.Equal(13, dragged.Top);
        Assert.Equal(213, dragged.Bottom);
    }

    // Each resize handle moves exactly the edges it names and leaves the others alone.
    [Theory]
    [InlineData("ResizeLeft", 15, 20, 110, 220)]
    [InlineData("ResizeRight", 10, 20, 115, 220)]
    [InlineData("ResizeTop", 10, 25, 110, 220)]
    [InlineData("ResizeBottom", 10, 20, 110, 225)]
    [InlineData("ResizeTopLeft", 15, 25, 110, 220)]
    [InlineData("ResizeTopRight", 10, 25, 115, 220)]
    [InlineData("ResizeBottomLeft", 15, 20, 110, 225)]
    [InlineData("ResizeBottomRight", 10, 20, 115, 225)]
    public void A_resize_drag_moves_only_the_edges_its_handle_names(
        string dragModeName,
        double left,
        double top,
        double right,
        double bottom)
    {
        var dragged = KiCadCalibrationGeometry.ApplyDrag(Box(), Mode(dragModeName), 5, 5);

        Assert.Equal(left, dragged.Left);
        Assert.Equal(top, dragged.Top);
        Assert.Equal(right, dragged.Right);
        Assert.Equal(bottom, dragged.Bottom);
    }

    // The behaviour that makes flipping possible at all: dragging one edge past its opposite is
    // ALLOWED, and the resulting inverted order is what "mirrored" means. Clamping here would
    // make a board impossible to flip by dragging.
    [Fact]
    public void Dragging_an_edge_past_its_opposite_flips_the_box_rather_than_clamping()
    {
        var flipped = KiCadCalibrationGeometry.ApplyDrag(Box(), LabelEditorDragMode.ResizeLeft, 200, 0);

        Assert.Equal(210, flipped.Left);
        Assert.Equal(110, flipped.Right);
        Assert.True(flipped.IsMirroredX);
    }

    // None is not a drag. The tab guards on it before calling, but the maths must agree.
    [Fact]
    public void A_drag_mode_of_none_changes_nothing()
    {
        var unchanged = KiCadCalibrationGeometry.ApplyDrag(Box(), LabelEditorDragMode.None, 40, 40);

        Assert.Equal(10, unchanged.Left);
        Assert.Equal(20, unchanged.Top);
        Assert.Equal(110, unchanged.Right);
        Assert.Equal(220, unchanged.Bottom);
    }

    // Drags are applied to the box as it stood when the drag STARTED, so re-applying a bigger
    // offset to that same start box gives the same answer as one direct drag. This is what stops
    // rounding drift accumulating over a long drag.
    [Fact]
    public void A_drag_is_derived_from_the_start_box_so_it_cannot_accumulate_drift()
    {
        var start = Box();

        var afterSmallStep = KiCadCalibrationGeometry.ApplyDrag(start, LabelEditorDragMode.Move, 3, 3);
        var afterFullDrag = KiCadCalibrationGeometry.ApplyDrag(start, LabelEditorDragMode.Move, 9, 9);

        // The intermediate position does not feed the next one; both come from "start".
        Assert.Equal(13, afterSmallStep.Left);
        Assert.Equal(19, afterFullDrag.Left);
    }

    // -----------------------------------------------------------------------------------------
    // Mirror flags
    // -----------------------------------------------------------------------------------------

    // Applying a saved mirror flag swaps the pair of edges it governs, which IS how mirroring is
    // represented - so an unmirrored box becomes mirrored and nothing else changes.
    [Fact]
    public void Applying_a_mirror_flag_swaps_that_axis_edges()
    {
        var mirroredX = KiCadCalibrationGeometry.ApplyMirrorFlags(Box(), mirrorX: true, mirrorY: false);

        Assert.Equal(110, mirroredX.Left);
        Assert.Equal(10, mirroredX.Right);
        Assert.True(mirroredX.IsMirroredX);
        Assert.False(mirroredX.IsMirroredY);

        // The untouched axis keeps its order.
        Assert.Equal(20, mirroredX.Top);
        Assert.Equal(220, mirroredX.Bottom);
    }

    [Fact]
    public void Applying_no_mirror_flags_leaves_the_box_alone()
    {
        var same = KiCadCalibrationGeometry.ApplyMirrorFlags(Box(), mirrorX: false, mirrorY: false);

        Assert.Equal(10, same.Left);
        Assert.Equal(20, same.Top);
        Assert.Equal(110, same.Right);
        Assert.Equal(220, same.Bottom);
    }

    // Applying the same flag twice returns the original box - the swap is its own inverse.
    [Fact]
    public void Applying_a_mirror_flag_twice_returns_the_original_box()
    {
        var once = KiCadCalibrationGeometry.ApplyMirrorFlags(Box(), mirrorX: true, mirrorY: true);
        var twice = KiCadCalibrationGeometry.ApplyMirrorFlags(once, mirrorX: true, mirrorY: true);

        Assert.Equal(10, twice.Left);
        Assert.Equal(20, twice.Top);
        Assert.Equal(110, twice.Right);
        Assert.Equal(220, twice.Bottom);
    }

    // -----------------------------------------------------------------------------------------
    // InsetForInitialVisibility - first-time calibration's corner handles landing on the image edge
    // -----------------------------------------------------------------------------------------
    //
    // Reported directly, with a screenshot: entering calibration mode for the first time on a
    // schematic (no saved calibration yet) seeds the box from the raw full-image bounds, so all
    // four corner handles sit exactly on the image's own edge - some clipped by the viewport frame
    // and effectively ungrabbable. These pin the fix: an edge sitting at (or within the margin of)
    // the image's own border gets pulled in; an edge that already has room is left alone.

    // The exact reported shape: a box seeded at the full image bounds gets every edge pulled in.
    [Fact]
    public void A_box_at_the_full_image_bounds_is_inset_on_every_edge()
    {
        var box = new KiCadCalibrationBox(0, 0, 1000, 800);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(box, imageWidth: 1000, imageHeight: 800);

        Assert.True(inset.Left > 0);
        Assert.True(inset.Top > 0);
        Assert.True(inset.Right < 1000);
        Assert.True(inset.Bottom < 800);
    }

    // A box that already has room on every side (e.g. seeded from a saved calibration covering
    // only part of the board) must be left exactly as computed - this only exists for the
    // full-image case, and must never nudge an already-reasonable box.
    [Fact]
    public void A_box_already_away_from_every_edge_is_left_untouched()
    {
        var box = new KiCadCalibrationBox(200, 150, 800, 600);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(box, imageWidth: 1000, imageHeight: 800);

        Assert.Equal(200, inset.Left);
        Assert.Equal(150, inset.Top);
        Assert.Equal(800, inset.Right);
        Assert.Equal(600, inset.Bottom);
    }

    // Only the edges actually AT the image border move - a box touching just the left/top must
    // not have its already-fine right/bottom edges disturbed.
    [Fact]
    public void Only_edges_touching_the_image_border_are_pulled_in()
    {
        var box = new KiCadCalibrationBox(0, 0, 500, 400);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(box, imageWidth: 1000, imageHeight: 800);

        Assert.True(inset.Left > 0);
        Assert.True(inset.Top > 0);
        Assert.Equal(500, inset.Right);
        Assert.Equal(400, inset.Bottom);
    }

    // The margin scales with the image, not a fixed pixel count - a large hi-res board scan gets a
    // proportionally larger inset than a small preview schematic.
    [Fact]
    public void The_margin_scales_with_the_image_size()
    {
        var small = KiCadCalibrationGeometry.InsetForInitialVisibility(
            new KiCadCalibrationBox(0, 0, 200, 100), imageWidth: 200, imageHeight: 100);

        var large = KiCadCalibrationGeometry.InsetForInitialVisibility(
            new KiCadCalibrationBox(0, 0, 4000, 2000), imageWidth: 4000, imageHeight: 2000);

        // Both insets are the same FRACTION of their own image, so the large image's absolute
        // inset is proportionally larger.
        Assert.Equal(small.Left / 200.0, large.Left / 4000.0, 6);
        Assert.Equal(small.Top / 100.0, large.Top / 2000.0, 6);
    }

    // Mirroring must survive the inset exactly as it survives every other edge operation in this
    // class - the inset works in normalised (ascending) edges and must come back out flipped.
    [Fact]
    public void Mirroring_survives_the_inset()
    {
        var mirroredBox = new KiCadCalibrationBox(1000, 0, 0, 800);
        Assert.True(mirroredBox.IsMirroredX);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(mirroredBox, imageWidth: 1000, imageHeight: 800);

        Assert.True(inset.IsMirroredX);
        Assert.False(inset.IsMirroredY);

        // The mirrored Left (stored as the box's Right, numerically) still moved in from 1000.
        Assert.True(inset.Left < 1000);
        Assert.True(inset.Right > 0);
    }

    // A degenerate box (already pinched down to a sliver against the image's own edge) must fall
    // back to the box exactly as given rather than producing an inverted or empty rectangle. The
    // margin is 3% of the IMAGE, so this needs a box whose own width/height is smaller than that
    // margin, with its far edge nowhere near the OPPOSITE margin (otherwise only the near edge
    // moves and nothing inverts) - a small image's own margin is too small to invert anything on
    // its own, which is what an earlier version of this test got wrong.
    [Fact]
    public void A_degenerate_box_leaves_it_untouched_rather_than_inverting_it()
    {
        // 3% of 1000 is 30: Left (0) is pulled in to 30, but Right (25) is nowhere near the
        // opposite margin (970) so it is left alone - and now sits BEHIND the pulled-in Left,
        // which is exactly the inversion the fallback exists to refuse.
        var box = new KiCadCalibrationBox(0, 0, 25, 800);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(box, imageWidth: 1000, imageHeight: 800);

        Assert.Equal(0, inset.Left);
        Assert.Equal(0, inset.Top);
        Assert.Equal(25, inset.Right);
        Assert.Equal(800, inset.Bottom);
    }

    // A zero/negative image size (no bitmap loaded yet) must not divide by zero or otherwise
    // misbehave - the box is simply returned as given.
    [Fact]
    public void A_zero_image_size_leaves_the_box_untouched()
    {
        var box = new KiCadCalibrationBox(0, 0, 100, 100);

        var inset = KiCadCalibrationGeometry.InsetForInitialVisibility(box, imageWidth: 0, imageHeight: 0);

        Assert.Equal(box.Left, inset.Left);
        Assert.Equal(box.Top, inset.Top);
        Assert.Equal(box.Right, inset.Right);
        Assert.Equal(box.Bottom, inset.Bottom);
    }
}
