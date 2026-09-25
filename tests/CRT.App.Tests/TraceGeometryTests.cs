using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The maths behind the user-drawn traces on a schematic, extracted from PolylineManagement.
//
// The rules being pinned down:
//  - A trace is STORED normalized (0..1 of the image content rect) and DRAWN in canvas pixels,
//    so it lands correctly at any zoom level and window size. The two conversions must round-trip.
//  - A degenerate content rect must not produce NaN - it is the ordinary pre-layout state.
//  - Clicking a trace projects onto the SEGMENT, never past its ends, so inserting a node cannot
//    kink the line.
//  - A dragged node snaps to its neighbours' axes independently on X and Y, which is what lets a
//    corner line up with the node before it horizontally and the one after it vertically.
//  - Traces saved in the old canvas-pixel format are detected by their own magnitude, since the
//    old format carried no version marker.
//
// Pure arithmetic: no collection, no UiTest, no statics.
// ###########################################################################################
public class TraceGeometryTests
{
    // A content rect deliberately offset from the origin, so a conversion that forgets to
    // subtract/add the origin fails rather than coincidentally passing.
    private static readonly Rect Content = new(100, 50, 400, 200);

    // -----------------------------------------------------------------------------------------
    // Coordinate conversion
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void The_content_rect_origin_maps_to_normalized_zero()
    {
        Assert.Equal(new Point(0, 0), TraceGeometry.CanvasToNormalized(new Point(100, 50), Content));
    }

    [Fact]
    public void The_content_rect_far_corner_maps_to_normalized_one()
    {
        Assert.Equal(new Point(1, 1), TraceGeometry.CanvasToNormalized(new Point(500, 250), Content));
    }

    [Fact]
    public void The_content_rect_centre_maps_to_normalized_half()
    {
        Assert.Equal(new Point(0.5, 0.5), TraceGeometry.CanvasToNormalized(new Point(300, 150), Content));
    }

    // The clamp is what stops a drag that leaves the image storing a node outside it, which would
    // then be drawn off the schematic on reload.
    [Fact]
    public void A_canvas_point_outside_the_content_rect_is_clamped_into_zero_to_one()
    {
        Assert.Equal(new Point(0, 0), TraceGeometry.CanvasToNormalized(new Point(-500, -500), Content));
        Assert.Equal(new Point(1, 1), TraceGeometry.CanvasToNormalized(new Point(9999, 9999), Content));
    }

    // The round-trip is the whole point of storing normalized: a node must come back where it went.
    [Fact]
    public void Converting_to_normalized_and_back_returns_the_original_canvas_point()
    {
        var original = new Point(233, 117);

        Point normalized = TraceGeometry.CanvasToNormalized(original, Content);
        Point roundTripped = TraceGeometry.NormalizedToCanvas(normalized, Content);

        Assert.Equal(original.X, roundTripped.X, 9);
        Assert.Equal(original.Y, roundTripped.Y, 9);
    }

    // The same normalized node must land in different canvas places at different zoom levels -
    // that is what makes the stored format zoom-independent.
    [Fact]
    public void The_same_normalized_node_maps_differently_as_the_content_rect_grows()
    {
        var node = new Point(0.5, 0.5);

        Point small = TraceGeometry.NormalizedToCanvas(node, new Rect(0, 0, 100, 100));
        Point large = TraceGeometry.NormalizedToCanvas(node, new Rect(0, 0, 1000, 1000));

        Assert.Equal(new Point(50, 50), small);
        Assert.Equal(new Point(500, 500), large);
    }

    // Pre-layout, the rect is empty. Dividing by its zero width would produce NaN and poison every
    // stored node, so this returns the origin instead.
    [Theory]
    [InlineData(0, 100)]
    [InlineData(100, 0)]
    [InlineData(0, 0)]
    public void A_degenerate_content_rect_normalizes_to_the_origin_rather_than_NaN(double width, double height)
    {
        Point result = TraceGeometry.CanvasToNormalized(new Point(42, 42), new Rect(0, 0, width, height));

        Assert.Equal(new Point(0, 0), result);
        Assert.False(double.IsNaN(result.X));
        Assert.False(double.IsNaN(result.Y));
    }

    // The inverse deliberately differs: it returns the input unchanged rather than collapsing every
    // node onto the origin, because the caller re-maps once real bounds arrive and collapsing would
    // be visible in the meantime.
    [Fact]
    public void A_degenerate_content_rect_leaves_a_normalized_point_unchanged()
    {
        var node = new Point(0.25, 0.75);

        Assert.Equal(node, TraceGeometry.NormalizedToCanvas(node, new Rect(0, 0, 0, 0)));
    }

    // -----------------------------------------------------------------------------------------
    // Point-to-segment distance
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_point_beside_a_segment_projects_perpendicularly_onto_it()
    {
        double distance = TraceGeometry.DistancePointToSegment(
            new Point(50, 30), new Point(0, 0), new Point(100, 0), out Point projection);

        Assert.Equal(30, distance, 9);
        Assert.Equal(new Point(50, 0), projection);
    }

    // The clamp to the segment: a point past the end projects onto the ENDPOINT, not onto the
    // infinite line beyond it. Without this, clicking off the end of a short trace would insert a
    // node in empty space.
    [Fact]
    public void A_point_beyond_the_segment_end_projects_onto_the_endpoint()
    {
        double distance = TraceGeometry.DistancePointToSegment(
            new Point(500, 0), new Point(0, 0), new Point(100, 0), out Point projection);

        Assert.Equal(400, distance, 9);
        Assert.Equal(new Point(100, 0), projection);
    }

    [Fact]
    public void A_point_before_the_segment_start_projects_onto_the_start()
    {
        double distance = TraceGeometry.DistancePointToSegment(
            new Point(-25, 0), new Point(0, 0), new Point(100, 0), out Point projection);

        Assert.Equal(25, distance, 9);
        Assert.Equal(new Point(0, 0), projection);
    }

    // A zero-length segment would divide by zero when normalising; it short-circuits instead.
    [Fact]
    public void A_zero_length_segment_measures_to_its_single_point()
    {
        double distance = TraceGeometry.DistancePointToSegment(
            new Point(3, 4), new Point(0, 0), new Point(0, 0), out Point projection);

        Assert.Equal(5, distance, 9);
        Assert.Equal(new Point(0, 0), projection);
    }

    [Fact]
    public void A_point_on_the_segment_measures_zero()
    {
        double distance = TraceGeometry.DistancePointToSegment(
            new Point(50, 50), new Point(0, 0), new Point(100, 100), out Point projection);

        Assert.Equal(0, distance, 9);
        Assert.Equal(50, projection.X, 9);
        Assert.Equal(50, projection.Y, 9);
    }

    // -----------------------------------------------------------------------------------------
    // Node snapping
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void A_node_within_tolerance_snaps_onto_its_neighbours_axis()
    {
        Point snapped = TraceGeometry.ApplyNodeSnapping(
            new Point(102, 200),
            new[] { new Point(100, 50) },
            tolerance: 5);

        Assert.Equal(100, snapped.X);
        Assert.Equal(200, snapped.Y);
    }

    [Fact]
    public void A_node_outside_tolerance_is_left_where_it_is()
    {
        Point snapped = TraceGeometry.ApplyNodeSnapping(
            new Point(140, 200),
            new[] { new Point(100, 50) },
            tolerance: 5);

        Assert.Equal(new Point(140, 200), snapped);
    }

    // The axes are decided independently, which is what lets a corner node align horizontally with
    // the node before it and vertically with the node after it in the same drag.
    [Fact]
    public void The_two_axes_snap_independently_to_different_neighbours()
    {
        Point snapped = TraceGeometry.ApplyNodeSnapping(
            new Point(101, 201),
            new[] { new Point(100, 999), new Point(999, 200) },
            tolerance: 5);

        Assert.Equal(100, snapped.X);
        Assert.Equal(200, snapped.Y);
    }

    [Fact]
    public void A_node_with_no_neighbours_is_never_snapped()
    {
        var point = new Point(123, 456);

        Assert.Equal(point, TraceGeometry.ApplyNodeSnapping(point, Array.Empty<Point>(), tolerance: 5));
    }

    // With two candidates on the same axis the closer one wins - snapping to the further neighbour
    // would jump the node past one the user can see.
    [Fact]
    public void The_closest_neighbour_wins_on_a_shared_axis()
    {
        Point snapped = TraceGeometry.ApplyNodeSnapping(
            new Point(100, 500),
            new[] { new Point(104, 0), new Point(101, 0) },
            tolerance: 10);

        Assert.Equal(101, snapped.X);
    }

    // -----------------------------------------------------------------------------------------
    // Legacy format detection
    // -----------------------------------------------------------------------------------------

    // Normalized coordinates are 0..1 by construction, so anything above 2.0 must be the old
    // canvas-pixel format. The margin above 1.0 is slack for rounding at an image's edge.
    [Fact]
    public void A_coordinate_above_two_is_treated_as_the_legacy_canvas_format()
    {
        Assert.True(TraceGeometry.IsLegacyCanvasCoordinate(350, 0.5));
        Assert.True(TraceGeometry.IsLegacyCanvasCoordinate(0.5, 350));
    }

    [Fact]
    public void A_normalized_coordinate_is_not_treated_as_legacy()
    {
        Assert.False(TraceGeometry.IsLegacyCanvasCoordinate(0.0, 0.0));
        Assert.False(TraceGeometry.IsLegacyCanvasCoordinate(1.0, 1.0));
    }

    // The slack band itself: a value just over 1.0 is rounding at the image edge, not legacy data.
    [Fact]
    public void A_coordinate_just_above_one_is_still_treated_as_normalized()
    {
        Assert.False(TraceGeometry.IsLegacyCanvasCoordinate(1.5, 1.5));
    }

    // A NEGATIVE coordinate is legacy too, and this is the case an upper-bound-only check misses.
    // Canvas pixels could go negative (a node drawn while the image was panned off the left or top
    // edge); a normalized one never can. Without this, a legacy trace lying wholly at negative
    // coordinates reads as already-normalized and gets multiplied by the content rect's size on
    // load, putting it thousands of pixels off-image.
    [Fact]
    public void A_negative_coordinate_is_treated_as_the_legacy_canvas_format()
    {
        Assert.True(TraceGeometry.IsLegacyCanvasCoordinate(-15, 40));
        Assert.True(TraceGeometry.IsLegacyCanvasCoordinate(0.5, -0.25));

        // The case that motivates it: BOTH coordinates negative and small in magnitude, so neither
        // trips the "> 2.0" half of the test.
        Assert.True(TraceGeometry.IsLegacyCanvasCoordinate(-1.5, -1.5));
    }
}
