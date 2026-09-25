using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for ViewportMath - wheel-zoom scaling and the selection set comparison, previously
// private statics inside TabSchematics.
//
// The zoom factor is the one users feel directly: it is why scrolling behaves the same on a
// Windows notched wheel and a high-resolution Linux or macOS trackpad.
public class ViewportMathTests
{
    private const double Base = 1.1;

    [Fact]
    public void One_windows_wheel_notch_gives_exactly_the_base_factor()
    {
        Assert.Equal(Base, ViewportMath.ComputeWheelZoomFactor(1.0, Base), precision: 12);
    }

    [Fact]
    public void Scrolling_the_other_way_gives_the_reciprocal()
    {
        Assert.Equal(1.0 / Base, ViewportMath.ComputeWheelZoomFactor(-1.0, Base), precision: 12);
    }

    [Fact]
    public void A_zoom_in_then_out_of_the_same_magnitude_returns_to_the_start()
    {
        double inFactor = ViewportMath.ComputeWheelZoomFactor(0.4, Base);
        double outFactor = ViewportMath.ComputeWheelZoomFactor(-0.4, Base);

        Assert.Equal(1.0, inFactor * outFactor, precision: 12);
    }

    [Fact]
    public void A_bigger_delta_zooms_further()
    {
        Assert.True(
            ViewportMath.ComputeWheelZoomFactor(2.0, Base) >
            ViewportMath.ComputeWheelZoomFactor(1.0, Base));
    }

    [Fact]
    public void Many_small_trackpad_deltas_zoom_less_than_one_full_notch()
    {
        // The whole point of the magnitude scaling: a high-resolution trackpad delivers lots of
        // tiny deltas, and each must move less than a full wheel notch.
        double smallStep = ViewportMath.ComputeWheelZoomFactor(0.1, Base);

        Assert.True(smallStep > 1.0);
        Assert.True(smallStep < ViewportMath.ComputeWheelZoomFactor(1.0, Base));
    }

    [Fact]
    public void A_huge_delta_is_clamped_so_one_event_cannot_zoom_wildly()
    {
        Assert.Equal(
            ViewportMath.ComputeWheelZoomFactor(3.0, Base),
            ViewportMath.ComputeWheelZoomFactor(500.0, Base),
            precision: 12);
    }

    [Fact]
    public void A_tiny_delta_is_floored_so_it_still_does_something()
    {
        Assert.Equal(
            ViewportMath.ComputeWheelZoomFactor(0.1, Base),
            ViewportMath.ComputeWheelZoomFactor(0.000001, Base),
            precision: 12);
    }

    [Fact]
    public void A_zero_delta_is_treated_as_a_zoom_out_at_the_floor_magnitude()
    {
        // Documents current behaviour: the sign test is "> 0", so exactly zero takes the
        // reciprocal branch rather than being a no-op.
        double factor = ViewportMath.ComputeWheelZoomFactor(0.0, Base);

        Assert.True(factor < 1.0);
        Assert.Equal(1.0 / ViewportMath.ComputeWheelZoomFactor(0.1, Base), factor, precision: 12);
    }

    // ------------------------------------------------- ComputeAxisTranslationRange

    // The numbers below are one real schematic viewport: the schematics container is one half
    // of a split pane, so it is 398 wide and 600 tall, and a landscape schematic fitted into it
    // by Stretch="Uniform" ends up 398 x 298.5 - it fills the width and leaves 301.5 points of
    // empty space below it. That empty space is the whole reason this function exists.
    private const double ViewportWidth = 398.0;
    private const double ViewportHeight = 600.0;
    private const double FittedContentWidth = 398.0;
    private const double FittedContentHeight = 298.5;

    [Fact]
    public void At_the_fitted_scale_only_the_fitted_position_is_allowed_on_the_filling_axis()
    {
        // Scale 1 on the axis the image fills: it is already exactly the viewport, so there is
        // nowhere to go.
        (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
            0.0, ViewportWidth, 0.0, FittedContentWidth, scale: 1.0);

        Assert.Equal(0.0, min, precision: 9);
        Assert.Equal(0.0, max, precision: 9);
    }

    [Fact]
    public void A_zoomed_image_may_be_panned_but_never_far_enough_to_show_an_empty_edge()
    {
        // Twice the fitted scale on the filling axis: 796 points of image in a 398 point
        // viewport, so the image may slide by 398 and no further in either direction.
        (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
            0.0, ViewportWidth, 0.0, FittedContentWidth, scale: 2.0);

        Assert.Equal(-398.0, min, precision: 9);
        Assert.Equal(0.0, max, precision: 9);
    }

    // ###########################################################################################
    // The bug this function was written for. On the letterboxed axis the old rule - keep the
    // image inside the viewport - forbade exactly the positions a cursor-anchored zoom needs,
    // so the point under the cursor slid away as you zoomed. Zooming 1.5x about a point 200
    // from the top has to land the translation on 200 * (1 - 1.5) = -100, which means letting
    // the top of the image travel above the top of the viewport.
    // ###########################################################################################
    [Fact]
    public void A_letterboxed_axis_allows_the_position_a_cursor_anchored_zoom_needs()
    {
        (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
            0.0, ViewportHeight, 0.0, FittedContentHeight, scale: 1.5);

        double anchoredAt200 = 200.0 * (1.0 - 1.5);

        Assert.True(
            min <= anchoredAt200 && anchoredAt200 <= max,
            $"Anchored zoom needs {anchoredAt200}, which is outside [{min}, {max}].");
    }

    [Fact]
    public void The_furthest_a_letterboxed_image_may_travel_is_what_anchoring_on_its_far_edge_needs()
    {
        // The cursor cannot be further down the image than its bottom edge, so anchoring there
        // is the extreme case - and it is exactly the limit. Anything beyond it would be the
        // image being pushed off the viewport rather than a point being held under the cursor.
        const double Scale = 1.5;

        (double min, _) = ViewportMath.ComputeAxisTranslationRange(
            0.0, ViewportHeight, 0.0, FittedContentHeight, Scale);

        Assert.Equal(FittedContentHeight * (1.0 - Scale), min, precision: 9);
    }

    [Fact]
    public void A_letterboxed_image_can_never_be_pushed_up_past_its_fitted_bottom_edge()
    {
        // The same limit stated the way a user sees it: however far in you zoom, the bottom of
        // the image can never rise above where the fitted first view put it, so the empty band
        // below the image can never grow bigger than it already is at the first view.
        foreach (double scale in new[] { 1.0, 1.5, 2.25, 5.0, 20.0 })
        {
            (double min, _) = ViewportMath.ComputeAxisTranslationRange(
                0.0, ViewportHeight, 0.0, FittedContentHeight, scale);

            double lowestBottomEdge = min + (scale * FittedContentHeight);

            Assert.Equal(FittedContentHeight, lowestBottomEdge, precision: 9);
        }
    }

    [Fact]
    public void The_range_never_shrinks_as_the_image_is_zoomed_in()
    {
        // A shrinking range means a zoom step could be legal and the next one not, which is felt
        // as the image snapping back mid-gesture.
        double previousSize = -1.0;

        foreach (double scale in new[] { 1.0, 1.25, 1.5, 2.0, 2.01, 3.0, 8.0, 20.0 })
        {
            (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
                0.0, ViewportHeight, 0.0, FittedContentHeight, scale);

            double size = max - min;

            Assert.True(size >= previousSize - 1e-9, $"The range shrank at scale {scale}.");
            previousSize = size;
        }
    }

    [Fact]
    public void An_edge_panel_that_hides_part_of_the_view_can_still_be_panned_out_from_behind()
    {
        // The net connections panel covers the right of the container, so the visible viewport
        // is narrower than the image. The last 98 points of image sit behind it and panning has
        // to be able to bring them out, even at the fitted scale.
        (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
            0.0, 300.0, 0.0, FittedContentWidth, scale: 1.0);

        Assert.Equal(-98.0, min, precision: 9);
        Assert.Equal(0.0, max, precision: 9);
    }

    [Fact]
    public void A_degenerate_content_rect_produces_a_usable_range()
    {
        // No image loaded yet: the content rect can be empty, and the caller still asks for a
        // range. It must come back ordered rather than inverted.
        (double min, double max) = ViewportMath.ComputeAxisTranslationRange(
            0.0, ViewportHeight, 0.0, 0.0, scale: 1.0);

        Assert.True(min <= max);
    }

    // ------------------------------------------------- SetEqualsOrdinalIgnoreCase

    [Fact]
    public void Two_sets_with_the_same_members_in_any_order_are_equal()
    {
        Assert.True(ViewportMath.SetEqualsOrdinalIgnoreCase(
            new[] { "GND", "VCC", "CLK" },
            new[] { "CLK", "GND", "VCC" }));
    }

    [Fact]
    public void Membership_ignores_case()
    {
        Assert.True(ViewportMath.SetEqualsOrdinalIgnoreCase(
            new[] { "gnd", "vcc" },
            new[] { "GND", "VCC" }));
    }

    [Fact]
    public void Different_sizes_are_not_equal()
    {
        Assert.False(ViewportMath.SetEqualsOrdinalIgnoreCase(
            new[] { "GND" },
            new[] { "GND", "VCC" }));
    }

    [Fact]
    public void Different_members_are_not_equal()
    {
        Assert.False(ViewportMath.SetEqualsOrdinalIgnoreCase(
            new[] { "GND", "VCC" },
            new[] { "GND", "CLK" }));
    }

    [Fact]
    public void Two_empty_sets_are_equal()
    {
        Assert.True(ViewportMath.SetEqualsOrdinalIgnoreCase(
            Array.Empty<string>(), Array.Empty<string>()));
    }

    [Fact]
    public void The_same_instance_is_equal_to_itself()
    {
        var set = new[] { "GND" };
        Assert.True(ViewportMath.SetEqualsOrdinalIgnoreCase(set, set));
    }

    [Fact]
    public void Duplicates_collapse_so_a_repeated_member_still_compares_equal_by_set()
    {
        // CURRENT BEHAVIOUR: the counts are compared first, so a duplicate on one side makes the
        // sizes differ and the result is false even though the SETS are equal.
        Assert.False(ViewportMath.SetEqualsOrdinalIgnoreCase(
            new[] { "GND", "GND" },
            new[] { "GND" }));
    }
}
