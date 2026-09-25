using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for OverlayCullGeometry - the bounds maths that decides which KiCad overlay primitives are
// worth drawing.
//
// The asymmetry here is the point: drawing a primitive that turned out to be off screen costs a
// little time, while skipping one that was actually visible makes copper vanish from the overlay.
// Every bound is therefore deliberately generous, and these tests pin that down so a later
// "tightening" cannot quietly introduce dropouts.
public sealed class OverlayCullGeometryTests
{
    // ------------------------------------------------------------------ stroke margin

    [Fact]
    public void A_stroke_margin_grows_the_box_on_every_side()
    {
        // A stroke straddles the path it follows, so the geometry's own bounds are too small.
        var inflated = OverlayCullGeometry.InflateForStroke(new Rect(10, 10, 20, 20), 4.0);

        Assert.Equal(6, inflated.X, precision: 9);
        Assert.Equal(6, inflated.Y, precision: 9);
        Assert.Equal(28, inflated.Width, precision: 9);
        Assert.Equal(28, inflated.Height, precision: 9);
    }

    [Theory]
    [InlineData(0.0)]
    [InlineData(-3.0)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_missing_or_malformed_thickness_still_leaves_a_margin(double thickness)
    {
        // Never shrink the box on bad input - a zero margin would let hairline traces at the very
        // edge of the viewport be culled away.
        var inflated = OverlayCullGeometry.InflateForStroke(new Rect(10, 10, 20, 20), thickness);

        Assert.True(inflated.Width >= 22);
        Assert.True(inflated.Height >= 22);
    }

    // ------------------------------------------------------------------- point bounds

    [Fact]
    public void Point_bounds_span_every_point()
    {
        var bounds = OverlayCullGeometry.BoundsOfPoints(new[]
        {
            new Point(5, 10),
            new Point(-3, 40),
            new Point(20, 2),
        });

        Assert.Equal(-3, bounds.X, precision: 9);
        Assert.Equal(2, bounds.Y, precision: 9);
        Assert.Equal(23, bounds.Width, precision: 9);
        Assert.Equal(38, bounds.Height, precision: 9);
    }

    [Fact]
    public void Point_bounds_of_nothing_are_empty()
    {
        Assert.Equal(default, OverlayCullGeometry.BoundsOfPoints(Array.Empty<Point>()));
        Assert.Equal(default, OverlayCullGeometry.BoundsOfPoints(null!));
    }

    [Fact]
    public void A_two_point_run_still_has_bounds()
    {
        // Most copper runs are exactly two points, so this is the common case rather than an edge one.
        var bounds = OverlayCullGeometry.BoundsOfPoints(new[] { new Point(0, 0), new Point(10, 0) });

        Assert.Equal(10, bounds.Width, precision: 9);
        Assert.Equal(0, bounds.Height, precision: 9);
    }

    // ----------------------------------------------------------------- rotated bounds

    [Fact]
    public void An_unrotated_rect_keeps_its_own_bounds()
    {
        var rect = new Rect(10, 20, 6, 2);

        Assert.Equal(rect, OverlayCullGeometry.BoundsOfRotatedRect(rect, 0));
    }

    [Fact]
    public void A_half_turn_keeps_its_own_bounds()
    {
        // A rectangle turned 180 degrees about its centre occupies exactly the same box.
        var rect = new Rect(10, 20, 6, 2);

        Assert.Equal(rect, OverlayCullGeometry.BoundsOfRotatedRect(rect, 180));
    }

    [Fact]
    public void A_rotated_rect_gets_a_box_that_holds_it_at_any_angle()
    {
        // A 90-degree pad swaps width and height, so the box has to cover both orientations. Using
        // the circumscribed circle covers every angle in between as well.
        var rotated = OverlayCullGeometry.BoundsOfRotatedRect(new Rect(0, 0, 6, 2), 90);

        Assert.True(rotated.Width >= 6);
        Assert.True(rotated.Height >= 6);
        Assert.Equal(new Point(3, 1), rotated.Center);
    }

    // -------------------------------------------------------------------- visible rect

    [Fact]
    public void Zooming_in_shrinks_the_visible_area_in_overlay_space()
    {
        // At 2x zoom the viewport shows half as much of the overlay in each direction - which is
        // exactly where culling starts paying off.
        var visible = OverlayCullGeometry.GetVisibleLocalRect(
            new Rect(0, 0, 800, 600),
            Matrix.CreateScale(2.0, 2.0));

        Assert.Equal(400, visible.Width, precision: 6);
        Assert.Equal(300, visible.Height, precision: 6);
    }

    [Fact]
    public void Panning_moves_the_visible_area()
    {
        var visible = OverlayCullGeometry.GetVisibleLocalRect(
            new Rect(0, 0, 800, 600),
            Matrix.CreateTranslation(-100, -50));

        Assert.Equal(100, visible.X, precision: 6);
        Assert.Equal(50, visible.Y, precision: 6);
    }

    [Fact]
    public void A_degenerate_view_matrix_falls_back_to_drawing_everything()
    {
        // A zero-scale matrix cannot be inverted. Returning the viewport keeps the overlay drawing
        // rather than going blank, which is the safe direction to fail in.
        var viewport = new Rect(0, 0, 800, 600);

        Assert.Equal(
            viewport,
            OverlayCullGeometry.GetVisibleLocalRect(viewport, new Matrix(0, 0, 0, 0, 0, 0)));
    }

    // ------------------------------------------------------------------- visibility

    [Fact]
    public void A_primitive_inside_the_view_is_drawn()
    {
        Assert.True(OverlayCullGeometry.IsVisible(new Rect(10, 10, 5, 5), new Rect(0, 0, 100, 100)));
    }

    [Fact]
    public void A_primitive_outside_the_view_is_skipped()
    {
        Assert.False(OverlayCullGeometry.IsVisible(new Rect(500, 500, 5, 5), new Rect(0, 0, 100, 100)));
    }

    [Fact]
    public void A_primitive_straddling_the_edge_is_drawn()
    {
        // Half-visible copper must still be drawn, or traces would be clipped at the viewport edge.
        Assert.True(OverlayCullGeometry.IsVisible(new Rect(-5, 10, 20, 5), new Rect(0, 0, 100, 100)));
    }

    [Fact]
    public void A_primitive_with_no_extent_is_drawn_rather_than_skipped()
    {
        // Some primitives carry no bounds of their own. Guessing they are invisible would be a silent
        // visual regression, so they are always drawn.
        Assert.True(OverlayCullGeometry.IsVisible(default, new Rect(0, 0, 100, 100)));
    }
}
