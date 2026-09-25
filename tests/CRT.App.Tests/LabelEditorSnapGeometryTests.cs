using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The label editor's snapping maths, extracted from TabSchematics where nothing could reach it.
//
// All coordinates here are BITMAP PIXELS, the space the editor works in. The rules being pinned
// down, each of which a user would notice breaking:
//
//  - An edge snaps to a neighbour's edge within 2px (snapThreshold), and does NOT outside it.
//  - A snap is REJECTED when a third highlight sits between the moving edge and the neighbour -
//    you cannot snap through an intervening component.
//  - Only highlights inside the visible viewport participate, so a neighbour scrolled off screen
//    cannot silently drag an edge.
//  - snapOnMatch: false means "draw guides for exact alignment but do not move anything", which
//    is what keyboard nudging uses.
//  - A row that is part of the current selection never snaps to itself or its companions.
//
// No collection needed and no UiTest: this is pure arithmetic with no statics and no controls.
// ###########################################################################################
public class LabelEditorSnapGeometryTests
{
    private const string Schematic = "Board top";

    private static EditableComponentHighlight Row(
        double x,
        double y,
        double width,
        double height,
        string label = "U1",
        string schematic = Schematic)
    {
        return new EditableComponentHighlight
        {
            SchematicName = schematic,
            BoardLabel = label,
            X = x,
            Y = y,
            Width = width,
            Height = height,
        };
    }

    // Nothing selected and the whole bitmap visible - the ordinary case.
    private static LabelEditorSnapContext Context(
        IReadOnlyList<EditableComponentHighlight> rows,
        LabelEditorDragMode dragMode,
        Rect? visiblePixelRect = null,
        Func<EditableComponentHighlight, bool>? isSelected = null,
        string schematicName = Schematic)
    {
        return new LabelEditorSnapContext(
            rows,
            dragMode,
            schematicName,
            visiblePixelRect,
            isSelected ?? (_ => false));
    }

    // -----------------------------------------------------------------------------------------
    // Resize snapping
    // -----------------------------------------------------------------------------------------

    // The core behaviour: a top edge 1px away from a neighbour's top edge is pulled onto it.
    [Fact]
    public void A_resized_top_edge_snaps_to_a_neighbour_within_the_threshold()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(100, top);
        Assert.NotEmpty(guides);
    }

    // Just outside 2px, nothing moves. This is the boundary that stops a snap feeling "sticky"
    // across the whole board.
    [Fact]
    public void A_resized_edge_outside_the_threshold_does_not_snap()
    {
        var moving = Row(10, 105, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 105, right = 60, bottom = 155;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(105, top);
        Assert.Empty(guides);
    }

    // SHIFT suppresses snapping entirely - the escape hatch for placing something exactly where
    // the neighbour would otherwise capture it.
    [Fact]
    public void Suppressing_the_snap_leaves_the_rectangle_untouched()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: true);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // A resize snap is meaningless when nothing is being resized. Both modes return early.
    //
    // Written as two Facts calling a helper rather than as a [Theory]: LabelEditorDragMode is
    // internal, so it cannot appear as a parameter on a public test method (xUnit requires test
    // classes to be public). HighlightAndLabelEditorGeometryTests carries the same note.
    [Fact]
    public void A_drag_mode_of_none_is_ignored_by_the_resize_snap()
    {
        AssertResizeSnapIsIgnoredFor(LabelEditorDragMode.None);
    }

    [Fact]
    public void A_drag_mode_of_move_is_ignored_by_the_resize_snap()
    {
        AssertResizeSnapIsIgnoredFor(LabelEditorDragMode.Move);
    }

    private static void AssertResizeSnapIsIgnoredFor(LabelEditorDragMode mode)
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, mode),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // The blocking rule, and the pair of tests that pin it down.
    //
    // A candidate is refused when a THIRD highlight lies between the moving edge and the edge it
    // would snap to, horizontally overlapping the moving rect. Without that rule an edge jumps
    // straight through an intervening component to land on something behind it.
    //
    // These two are deliberately identical except for the blocker, because that is the only way
    // to prove the rule fires: asserting "it did not snap" alone passes just as well when the
    // candidate was never in range. The first test establishes that this geometry DOES snap; the
    // second adds the blocker and shows the same snap is refused.
    [Fact]
    public void A_snap_is_allowed_when_nothing_blocks_the_path()
    {
        var moving = Row(10, 101, 50, 50);

        // Overlaps the moving rect horizontally, top edge 1px above the moving top edge.
        var neighbour = Row(10, 100, 50, 10, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(100, top);
    }

    [Fact]
    public void A_snap_is_refused_when_another_highlight_blocks_the_path()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(10, 100, 50, 10, "U2");

        // Same span, but occupying the gap between the moving top edge (101) and the candidate
        // edge (100) - so the path to the candidate is obstructed and the snap must be refused.
        var blocker = Row(10, 100.2, 50, 0.6, "U3");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour, blocker }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        // Unchanged: the only candidate within threshold was blocked.
        Assert.Equal(101, top);
    }

    // Only what is on screen may participate. A neighbour outside the visible rect is ignored,
    // so scrolling changes what snaps - deliberately.
    [Fact]
    public void A_neighbour_outside_the_visible_viewport_does_not_snap()
    {
        var moving = Row(10, 101, 50, 50);
        var offScreen = Row(5000, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(
                new[] { moving, offScreen },
                LabelEditorDragMode.ResizeTop,
                visiblePixelRect: new Rect(0, 0, 500, 500)),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // A null visible rect means "no viewport information", and every row participates. This is
    // the state before the bitmap or view matrix is usable, and it must not disable snapping.
    [Fact]
    public void A_null_visible_rect_lets_every_neighbour_participate()
    {
        var moving = Row(10, 101, 50, 50);
        var faraway = Row(5000, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, faraway }, LabelEditorDragMode.ResizeTop, visiblePixelRect: null),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(100, top);
    }

    // Rows belonging to a different schematic are never candidates, however close they are - the
    // working set spans every schematic the editor has touched.
    [Fact]
    public void A_neighbour_on_another_schematic_is_never_a_candidate()
    {
        var moving = Row(10, 101, 50, 50);
        var otherSchematic = Row(200, 100, 50, 50, "U2", schematic: "Board bottom");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, otherSchematic }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // A selected row is skipped as a snap target: dragging a multi-selection must not have its
    // members snapping to each other.
    [Fact]
    public void A_selected_neighbour_is_skipped_as_a_snap_target()
    {
        var moving = Row(10, 101, 50, 50);
        var selectedNeighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(
                new[] { moving, selectedNeighbour },
                LabelEditorDragMode.ResizeTop,
                isSelected: row => ReferenceEquals(row, selectedNeighbour)),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // snapOnMatch: false is the keyboard-nudge mode - guides appear for an EXACT alignment, but
    // the rectangle is never moved, so the keypress stays in control of position.
    [Fact]
    public void With_snap_on_match_disabled_an_exact_alignment_guides_without_moving()
    {
        var moving = Row(10, 100, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 100, right = 60, bottom = 150;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false,
            snapOnMatch: false);

        Assert.Equal(100, top);
        Assert.NotEmpty(guides);
    }

    // The same mode with a NEAR (but inexact) alignment produces nothing at all: no move, and no
    // guide either, since guideMatchThreshold (0.5) is tighter than the snap threshold.
    [Fact]
    public void With_snap_on_match_disabled_a_near_alignment_produces_no_guide()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false,
            snapOnMatch: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // Re-pointing a context at a different edge is what ApplyNewRectangleSnap uses to drive all
    // four in turn. It is done by COPYING the context (WithDragMode) rather than by passing a
    // second override argument alongside it: with both, a caller could hand ApplyResizeSnap a
    // context saying one thing and an override saying another, and only one of them would be read.
    // The copy's mode must be the one that takes effect.
    [Fact]
    public void With_drag_mode_re_points_a_context_at_a_different_edge()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        // The original context says Move, which the resize snap ignores outright; the copy says
        // ResizeTop, so the top edge snaps to the neighbour's 100.
        var moveContext = Context(new[] { moving, neighbour }, LabelEditorDragMode.Move);

        LabelEditorSnapGeometry.ApplyResizeSnap(
            moveContext.WithDragMode(LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false,
            snapOnMatch: true);

        Assert.Equal(100, top);

        // The copy is a copy: the original is a readonly struct and still says Move.
        Assert.Equal(LabelEditorDragMode.Move, moveContext.DragMode);
    }

    // WithDragMode changes the mode and nothing else - the rows, schematic, visible rect and
    // selection predicate all carry over, or a snap driven through a copy would quietly see a
    // different world than the drag that built the context.
    [Fact]
    public void With_drag_mode_carries_every_other_context_value_over()
    {
        var rows = new[] { Row(10, 101, 50, 50), Row(200, 100, 50, 50, "U2") };
        var original = Context(rows, LabelEditorDragMode.Move);

        var copy = original.WithDragMode(LabelEditorDragMode.ResizeLeft);

        Assert.Equal(LabelEditorDragMode.ResizeLeft, copy.DragMode);
        Assert.Same(original.WorkingHighlights, copy.WorkingHighlights);
        Assert.Equal(original.SchematicName, copy.SchematicName);
        Assert.Equal(original.VisiblePixelRect, copy.VisiblePixelRect);
        Assert.Same(original.IsSelected, copy.IsSelected);
    }

    // A left-edge resize snaps horizontally, proving the X axis is wired the same way as Y.
    [Fact]
    public void A_resized_left_edge_snaps_horizontally()
    {
        var moving = Row(101, 300, 50, 50);
        var neighbour = Row(100, 10, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        double left = 101, top = 300, right = 151, bottom = 350;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.ResizeLeft),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(100, left);
    }

    // An empty working set cannot throw - the editor's very first drawn rectangle hits this.
    [Fact]
    public void An_empty_working_set_snaps_nothing_and_does_not_throw()
    {
        var moving = Row(10, 101, 50, 50);
        var guides = new List<(Point Start, Point End)>();

        double left = 10, top = 101, right = 60, bottom = 151;

        LabelEditorSnapGeometry.ApplyResizeSnap(
            Context(Array.Empty<EditableComponentHighlight>(), LabelEditorDragMode.ResizeTop),
            moving,
            ref left, ref top, ref right, ref bottom,
            guides,
            suppressSnap: false);

        Assert.Equal(101, top);
        Assert.Empty(guides);
    }

    // -----------------------------------------------------------------------------------------
    // Move snapping
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_moved_selection_snaps_its_bounds_to_a_neighbour()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var bounds = new Rect(10, 101, 50, 50);
        var sourceRects = new Dictionary<EditableComponentHighlight, Rect>
        {
            [moving] = new Rect(10, 101, 50, 50),
        };

        LabelEditorSnapGeometry.ApplyMoveSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.Move),
            new[] { moving },
            sourceRects,
            ref bounds,
            guides,
            suppressSnap: false);

        Assert.Equal(100, bounds.Top);
    }

    [Fact]
    public void A_moved_selection_outside_the_threshold_does_not_snap()
    {
        var moving = Row(10, 106, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var bounds = new Rect(10, 106, 50, 50);
        var sourceRects = new Dictionary<EditableComponentHighlight, Rect>
        {
            [moving] = new Rect(10, 106, 50, 50),
        };

        LabelEditorSnapGeometry.ApplyMoveSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.Move),
            new[] { moving },
            sourceRects,
            ref bounds,
            guides,
            suppressSnap: false);

        Assert.Equal(106, bounds.Top);
    }

    // An empty selection is a no-op rather than a crash - reachable by starting a drag and
    // releasing with nothing selected.
    [Fact]
    public void Moving_an_empty_selection_does_nothing()
    {
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var bounds = new Rect(10, 101, 50, 50);

        LabelEditorSnapGeometry.ApplyMoveSnap(
            Context(new[] { neighbour }, LabelEditorDragMode.Move),
            Array.Empty<EditableComponentHighlight>(),
            new Dictionary<EditableComponentHighlight, Rect>(),
            ref bounds,
            guides,
            suppressSnap: false);

        Assert.Equal(new Rect(10, 101, 50, 50), bounds);
        Assert.Empty(guides);
    }

    [Fact]
    public void A_suppressed_move_snap_leaves_the_bounds_untouched()
    {
        var moving = Row(10, 101, 50, 50);
        var neighbour = Row(200, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var bounds = new Rect(10, 101, 50, 50);
        var sourceRects = new Dictionary<EditableComponentHighlight, Rect>
        {
            [moving] = new Rect(10, 101, 50, 50),
        };

        LabelEditorSnapGeometry.ApplyMoveSnap(
            Context(new[] { moving, neighbour }, LabelEditorDragMode.Move),
            new[] { moving },
            sourceRects,
            ref bounds,
            guides,
            suppressSnap: true);

        Assert.Equal(new Rect(10, 101, 50, 50), bounds);
        Assert.Empty(guides);
    }

    // -----------------------------------------------------------------------------------------
    // New-rectangle snapping
    // -----------------------------------------------------------------------------------------

    // The newly drawn rectangle is snapped on all four edges in one pass. Previously this method
    // mutated the tab's drag-mode field around four calls and restored it in a finally; it now
    // passes the mode as an override, so this test also pins that rewrite down.
    [Fact]
    public void A_new_rectangle_snaps_on_all_four_edges()
    {
        var neighbour = Row(100, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        // Every edge sits 1px off the neighbour's corresponding edge.
        var rect = new Rect(101, 101, 48, 48);

        LabelEditorSnapGeometry.ApplyNewRectangleSnap(
            Context(new[] { neighbour }, LabelEditorDragMode.None),
            ref rect,
            guides,
            suppressSnap: false);

        Assert.Equal(100, rect.Left);
        Assert.Equal(100, rect.Top);
        Assert.Equal(150, rect.Right);
        Assert.Equal(150, rect.Bottom);
    }

    [Fact]
    public void A_suppressed_new_rectangle_snap_leaves_the_rectangle_untouched()
    {
        var neighbour = Row(100, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var rect = new Rect(101, 101, 48, 48);

        LabelEditorSnapGeometry.ApplyNewRectangleSnap(
            Context(new[] { neighbour }, LabelEditorDragMode.None),
            ref rect,
            guides,
            suppressSnap: true);

        Assert.Equal(new Rect(101, 101, 48, 48), rect);
        Assert.Empty(guides);
    }

    // A blank schematic name means there is nothing to snap against - the editor is not showing
    // a schematic, so the method returns before touching the rectangle.
    [Fact]
    public void A_new_rectangle_is_not_snapped_when_no_schematic_is_showing()
    {
        var neighbour = Row(100, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var rect = new Rect(101, 101, 48, 48);

        LabelEditorSnapGeometry.ApplyNewRectangleSnap(
            Context(new[] { neighbour }, LabelEditorDragMode.None, schematicName: "   "),
            ref rect,
            guides,
            suppressSnap: false);

        Assert.Equal(new Rect(101, 101, 48, 48), rect);
    }

    // A zero-area drag (a click without a drag) must not be snapped into existence.
    [Theory]
    [InlineData(0, 10)]
    [InlineData(10, 0)]
    [InlineData(0, 0)]
    public void A_new_rectangle_with_no_area_is_left_alone(double width, double height)
    {
        var neighbour = Row(100, 100, 50, 50, "U2");
        var guides = new List<(Point Start, Point End)>();

        var rect = new Rect(101, 101, width, height);

        LabelEditorSnapGeometry.ApplyNewRectangleSnap(
            Context(new[] { neighbour }, LabelEditorDragMode.None),
            ref rect,
            guides,
            suppressSnap: false);

        Assert.Equal(new Rect(101, 101, width, height), rect);
        Assert.Empty(guides);
    }
}
