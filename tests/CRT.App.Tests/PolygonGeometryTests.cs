using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for PolygonGeometry - the maths that decides which copper the KiCad overlay lights up.
//
// This logic used to live as private static methods inside TabSchematics, where nothing could
// reach it. It answers questions like "does this track run into that ground pour?", and a bug
// here does not crash: it silently highlights the wrong copper, which is worse.
//
// All coordinates are KiCad world millimetres.
public class PolygonGeometryTests
{
    // A 10x10 square with its corner at the origin.
    private static readonly IReadOnlyList<Point> Square = new[]
    {
        new Point(0, 0), new Point(10, 0), new Point(10, 10), new Point(0, 10)
    };

    private static IReadOnlyList<IReadOnlyList<Point>> Zone(params IReadOnlyList<Point>[] polygons) => polygons;

    // ------------------------------------------------------------- DistanceToSegment

    [Theory]
    [InlineData(5, 5, 5)]     // straight above the middle
    [InlineData(0, 3, 3)]     // above the start point
    [InlineData(10, 4, 4)]    // above the end point
    public void DistanceToSegment_measures_perpendicular_distance(double px, double py, double expected)
    {
        Assert.Equal(expected, PolygonGeometry.DistanceToSegment(new Point(px, py), 0, 0, 10, 0), precision: 9);
    }

    [Fact]
    public void DistanceToSegment_clamps_to_the_endpoints_rather_than_the_infinite_line()
    {
        // A point beyond the end of the segment measures to the endpoint, not to the projection.
        Assert.Equal(5, PolygonGeometry.DistanceToSegment(new Point(15, 0), 0, 0, 10, 0), precision: 9);
        Assert.Equal(5, PolygonGeometry.DistanceToSegment(new Point(-5, 0), 0, 0, 10, 0), precision: 9);
    }

    [Fact]
    public void DistanceToSegment_handles_a_zero_length_segment()
    {
        // Coincident start/end would divide by zero if it were not special-cased.
        Assert.Equal(5, PolygonGeometry.DistanceToSegment(new Point(3, 4), 0, 0, 0, 0), precision: 9);
    }

    [Fact]
    public void DistanceToSegment_is_zero_on_the_segment()
    {
        Assert.Equal(0, PolygonGeometry.DistanceToSegment(new Point(5, 0), 0, 0, 10, 0), precision: 9);
    }

    // -------------------------------------------------------------- IsPointInPolygon

    [Theory]
    [InlineData(5, 5)]
    [InlineData(0.001, 0.001)]
    [InlineData(9.999, 9.999)]
    public void IsPointInPolygon_finds_points_inside(double x, double y)
    {
        Assert.True(PolygonGeometry.IsPointInPolygon(Square, new Point(x, y)));
    }

    [Theory]
    [InlineData(-1, 5)]
    [InlineData(11, 5)]
    [InlineData(5, -1)]
    [InlineData(5, 11)]
    public void IsPointInPolygon_rejects_points_outside(double x, double y)
    {
        Assert.False(PolygonGeometry.IsPointInPolygon(Square, new Point(x, y)));
    }

    [Fact]
    public void IsPointInPolygon_needs_at_least_three_vertices()
    {
        Assert.False(PolygonGeometry.IsPointInPolygon(new[] { new Point(0, 0), new Point(10, 0) }, new Point(5, 0)));
        Assert.False(PolygonGeometry.IsPointInPolygon(Array.Empty<Point>(), new Point(0, 0)));
    }

    [Fact]
    public void IsPointInPolygon_handles_a_concave_shape()
    {
        // An L: the notch must read as outside even though it is inside the bounding box.
        var lShape = new[]
        {
            new Point(0, 0), new Point(10, 0), new Point(10, 4),
            new Point(4, 4), new Point(4, 10), new Point(0, 10)
        };

        Assert.True(PolygonGeometry.IsPointInPolygon(lShape, new Point(2, 2)));
        Assert.True(PolygonGeometry.IsPointInPolygon(lShape, new Point(8, 2)));
        Assert.False(PolygonGeometry.IsPointInPolygon(lShape, new Point(8, 8)));   // the notch
    }

    [Fact]
    public void IsPointInPolygon_gives_the_same_answer_for_either_winding_direction()
    {
        var clockwise = Square.Reverse().ToList();

        Assert.Equal(
            PolygonGeometry.IsPointInPolygon(Square, new Point(5, 5)),
            PolygonGeometry.IsPointInPolygon(clockwise, new Point(5, 5)));
    }

    // ------------------------------------------------- GetDistanceToPolygonBoundary

    [Fact]
    public void GetDistanceToPolygonBoundary_measures_to_the_nearest_edge_from_inside()
    {
        // (2,5) is 2 from the left edge and 8 from the right.
        Assert.Equal(2, PolygonGeometry.GetDistanceToPolygonBoundary(new Point(2, 5), Square), precision: 9);
    }

    [Fact]
    public void GetDistanceToPolygonBoundary_measures_to_the_nearest_edge_from_outside()
    {
        Assert.Equal(3, PolygonGeometry.GetDistanceToPolygonBoundary(new Point(13, 5), Square), precision: 9);
    }

    [Fact]
    public void GetDistanceToPolygonBoundary_returns_max_for_a_degenerate_polygon()
    {
        Assert.Equal(
            double.MaxValue,
            PolygonGeometry.GetDistanceToPolygonBoundary(new Point(0, 0), new[] { new Point(1, 1) }));
    }

    // ------------------------------------------------------------ IsPointInOrNearZone

    [Fact]
    public void A_point_inside_the_zone_reports_zero_distance()
    {
        Assert.True(PolygonGeometry.IsPointInOrNearZone(
            new Point(5, 5), Zone(Square), toleranceWorld: 0, out double distance));

        Assert.Equal(0, distance);
    }

    [Fact]
    public void A_point_just_outside_counts_when_inside_the_tolerance()
    {
        Assert.True(PolygonGeometry.IsPointInOrNearZone(
            new Point(11, 5), Zone(Square), toleranceWorld: 1.5, out double distance));

        Assert.Equal(1, distance, precision: 9);
    }

    [Fact]
    public void A_point_beyond_the_tolerance_does_not_count_but_still_reports_its_distance()
    {
        Assert.False(PolygonGeometry.IsPointInOrNearZone(
            new Point(15, 5), Zone(Square), toleranceWorld: 1.0, out double distance));

        Assert.Equal(5, distance, precision: 9);
    }

    [Fact]
    public void The_nearest_of_several_polygons_wins()
    {
        var far = new[] { new Point(100, 100), new Point(110, 100), new Point(110, 110), new Point(100, 110) };

        Assert.False(PolygonGeometry.IsPointInOrNearZone(
            new Point(-4, 5), Zone(far, Square), toleranceWorld: 1.0, out double distance));

        Assert.Equal(4, distance, precision: 9);
    }

    [Fact]
    public void An_empty_zone_never_matches()
    {
        Assert.False(PolygonGeometry.IsPointInOrNearZone(
            new Point(5, 5), Zone(), toleranceWorld: 1000, out _));
    }

    // -------------------------------------------------------------- touch predicates

    [Fact]
    public void A_circle_overlapping_the_zone_touches_it()
    {
        Assert.True(PolygonGeometry.DoesCircleTouchZone(new Point(11, 5), radiusWorld: 2, Zone(Square)));
    }

    [Fact]
    public void A_circle_clear_of_the_zone_does_not_touch_it()
    {
        Assert.False(PolygonGeometry.DoesCircleTouchZone(new Point(20, 5), radiusWorld: 2, Zone(Square)));
    }

    [Fact]
    public void A_segment_with_an_endpoint_in_the_zone_touches_it()
    {
        Assert.True(PolygonGeometry.DoesSegmentTouchZone(
            new Point(5, 5), new Point(50, 50), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void A_segment_crossing_the_zone_with_both_ends_outside_still_touches_it()
    {
        // Both endpoints are clear, so this only passes because the midpoints are sampled.
        Assert.True(PolygonGeometry.DoesSegmentTouchZone(
            new Point(-5, 5), new Point(15, 5), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void A_segment_clear_of_the_zone_does_not_touch_it()
    {
        Assert.False(PolygonGeometry.DoesSegmentTouchZone(
            new Point(-5, 50), new Point(15, 50), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void A_segment_never_touches_an_empty_zone()
    {
        Assert.False(PolygonGeometry.DoesSegmentTouchZone(
            new Point(0, 0), new Point(1, 1), radiusWorld: 1, Zone()));
    }

    [Fact]
    public void An_arc_with_a_control_point_in_the_zone_touches_it()
    {
        Assert.True(PolygonGeometry.DoesArcTouchZone(
            new Point(5, 5), new Point(50, 50), new Point(90, 90), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void An_arc_bowing_through_the_zone_touches_it_via_sampling()
    {
        // Start, mid and end are all outside the square; the curve still passes through it.
        Assert.True(PolygonGeometry.DoesArcTouchZone(
            new Point(-5, 5), new Point(5, -6), new Point(15, 5), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void An_arc_clear_of_the_zone_does_not_touch_it()
    {
        Assert.False(PolygonGeometry.DoesArcTouchZone(
            new Point(-5, 50), new Point(5, 60), new Point(15, 50), radiusWorld: 0.1, Zone(Square)));
    }

    [Fact]
    public void An_arc_never_touches_an_empty_zone()
    {
        Assert.False(PolygonGeometry.DoesArcTouchZone(
            new Point(0, 0), new Point(1, 1), new Point(2, 2), radiusWorld: 1, Zone()));
    }

    // ------------------------------------------------------------ GetPolygonSetBounds

    [Fact]
    public void Bounds_span_every_polygon_in_the_set()
    {
        var second = new[] { new Point(20, 20), new Point(30, 20), new Point(30, 30), new Point(20, 30) };

        Rect bounds = PolygonGeometry.GetPolygonSetBounds(Zone(Square, second));

        Assert.Equal(0, bounds.X);
        Assert.Equal(0, bounds.Y);
        Assert.Equal(30, bounds.Width);
        Assert.Equal(30, bounds.Height);
    }

    [Fact]
    public void Bounds_handle_negative_coordinates()
    {
        var negative = new[] { new Point(-10, -10), new Point(-5, -10), new Point(-5, -5), new Point(-10, -5) };

        Rect bounds = PolygonGeometry.GetPolygonSetBounds(Zone(negative));

        Assert.Equal(-10, bounds.X);
        Assert.Equal(-10, bounds.Y);
        Assert.Equal(5, bounds.Width);
    }

    [Fact]
    public void Bounds_of_an_empty_set_are_the_default_rect()
    {
        Assert.Equal(default, PolygonGeometry.GetPolygonSetBounds(Zone()));
    }

    [Fact]
    public void A_degenerate_polygon_still_gets_a_non_zero_size()
    {
        // A zero-width rect would collapse any downstream scaling, so the size is floored.
        var line = new[] { new Point(5, 5), new Point(5, 15) };

        Rect bounds = PolygonGeometry.GetPolygonSetBounds(Zone(line));

        Assert.True(bounds.Width > 0);
        Assert.Equal(10, bounds.Height);
    }
}
