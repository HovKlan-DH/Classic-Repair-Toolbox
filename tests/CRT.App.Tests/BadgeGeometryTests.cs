using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// The worklog "#N" badges keep a constant size on screen while the board zooms, via a
// ScaleTransform centred on the badge (RenderTransformOrigin 0.5,0.5). Canvas.SetLeft/SetTop
// position the PRE-transform layout box, so pinning a badge to a point on the image needs an
// offset that depends on the scale.
//
// The bug these pin down: the offset used to be "minus half the SCALED size", which is correct at
// scale 0.5 and wrong at every other zoom. The badges slid away from the top-left corner of their
// marked areas as the user zoomed in - by nearly a badge-width at high zoom.
//
// The worklog pills are CENTRED on their area's corner - half the pill hanging outside, which
// reads as a label attached to the area. GetCenterScaledRenderedTopLeft is where the scale still
// matters: the viewport clamp needs the badge's true on-screen rectangle, not its layout box.
public class BadgeGeometryTests
{
    private static readonly Size BadgeSize = new(60, 20);
    // The centre offset is half the badge back, in each axis, with no scale term - a centred
    // ScaleTransform maps a box's centre to itself, so the compensation a CORNER anchor would need
    // cancels out completely. The independence is now structural (the function takes no scale), and
    // this pins the value itself.
    [Fact]
    public void The_centre_offset_is_half_the_badge_in_each_axis()
    {
        var offset = BadgeGeometry.GetCenterScaledCentreOffset(BadgeSize);

        Assert.Equal(-(BadgeSize.Width / 2.0), offset.X, 9);
        Assert.Equal(-(BadgeSize.Height / 2.0), offset.Y, 9);
    }

    // Width and height are compensated independently - a wide, short badge (which is what the "#N"
    // pills are) must not take its vertical offset from its width.
    [Fact]
    public void Width_and_height_are_compensated_independently()
    {
        var offset = BadgeGeometry.GetCenterScaledCentreOffset(new Size(100, 10));

        Assert.Equal(-50.0, offset.X, 6);
        Assert.Equal(-5.0, offset.Y, 6);
    }

    // Exactly half the badge sits either side of the anchor, in both axes, whatever the zoom -
    // which is what makes it read as attached to the corner rather than placed near it.
    [Theory]
    [InlineData(0.25)]
    [InlineData(1.0)]
    [InlineData(4.0)]
    public void Half_the_badge_hangs_outside_the_anchor_at_any_scale(double scale)
    {
        var anchor = new Point(100, 250);

        var offset = BadgeGeometry.GetCenterScaledCentreOffset(BadgeSize);
        var layoutTopLeft = new Point(anchor.X + offset.X, anchor.Y + offset.Y);

        var renderedTopLeft = BadgeGeometry.GetCenterScaledRenderedTopLeft(layoutTopLeft, BadgeSize, scale);

        double onScreenWidth = BadgeSize.Width * scale;
        double onScreenHeight = BadgeSize.Height * scale;

        Assert.Equal(onScreenWidth / 2.0, anchor.X - renderedTopLeft.X, 6);
        Assert.Equal(onScreenHeight / 2.0, anchor.Y - renderedTopLeft.Y, 6);
    }

    [Fact]
    public void An_unmeasured_badge_gets_no_centre_offset()
    {
        var offset = BadgeGeometry.GetCenterScaledCentreOffset(new Size(0, 0));

        Assert.Equal(0.0, offset.X, 9);
        Assert.Equal(0.0, offset.Y, 9);
    }

    // ------------------------------------------------------------- viewport clamping

    private static readonly Size Viewport = new(800, 600);

    // A badge fully inside the view must not be touched. Nudging one that already fits would make
    // every badge drift as the user pans, which is worse than the problem being solved.
    [Fact]
    public void A_badge_already_inside_the_viewport_is_not_moved()
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(100, 100, 60, 20), Viewport);

        Assert.Equal(0.0, nudge.X, 9);
        Assert.Equal(0.0, nudge.Y, 9);
    }

    // The reported bug: an area whose corner sits near the left edge puts half its badge
    // off-screen, where the "#N" cannot be read and the badge cannot be clicked. It is pushed right
    // by exactly the overhang.
    [Fact]
    public void A_badge_hanging_off_the_left_edge_is_pushed_right()
    {
        // 25px of a 60px-wide badge is off-screen.
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(-25, 100, 60, 20), Viewport);

        Assert.Equal(25.0, nudge.X, 6);
        Assert.Equal(0.0, nudge.Y, 9);
    }

    [Fact]
    public void A_badge_hanging_off_the_right_edge_is_pushed_left()
    {
        // Right edge at 810 against a 800-wide viewport.
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(750, 100, 60, 20), Viewport);

        Assert.Equal(-10.0, nudge.X, 6);
        Assert.Equal(0.0, nudge.Y, 9);
    }

    [Theory]
    [InlineData(-15.0, 15.0)]
    [InlineData(595.0, -15.0)]
    public void A_badge_hanging_off_the_top_or_bottom_is_pushed_back_vertically(double y, double expectedDy)
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(100, y, 60, 20), Viewport);

        Assert.Equal(0.0, nudge.X, 9);
        Assert.Equal(expectedDy, nudge.Y, 6);
    }

    // A corner overhang is corrected in both axes at once, not one at a time.
    [Fact]
    public void A_badge_off_a_corner_is_pushed_back_in_both_axes()
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(-30, -10, 60, 20), Viewport);

        Assert.Equal(30.0, nudge.X, 6);
        Assert.Equal(10.0, nudge.Y, 6);
    }

    // Applying the nudge must actually bring the badge inside - the property the UI depends on,
    // asserted directly rather than inferred from the deltas above.
    [Theory]
    [InlineData(-25.0, 100.0)]
    [InlineData(770.0, 100.0)]
    [InlineData(100.0, -18.0)]
    [InlineData(100.0, 592.0)]
    [InlineData(-40.0, -40.0)]
    public void A_nudged_badge_ends_up_fully_inside_the_viewport(double x, double y)
    {
        var badge = new Rect(x, y, 60, 20);

        var nudge = BadgeGeometry.GetViewportNudge(badge, Viewport);
        var moved = new Rect(badge.X + nudge.X, badge.Y + nudge.Y, badge.Width, badge.Height);

        Assert.True(moved.Left >= -0.001, $"still off the left: {moved}");
        Assert.True(moved.Top >= -0.001, $"still off the top: {moved}");
        Assert.True(moved.Right <= Viewport.Width + 0.001, $"still off the right: {moved}");
        Assert.True(moved.Bottom <= Viewport.Height + 0.001, $"still off the bottom: {moved}");
    }

    // The margin keeps a nudged badge just clear of the edge instead of flush against it.
    [Fact]
    public void The_margin_insets_a_nudged_badge_from_the_edge()
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(-25, 100, 60, 20), Viewport, margin: 4.0);

        Assert.Equal(29.0, nudge.X, 6);
    }

    // A badge too big for the viewport cannot satisfy both edges. Pinning the left keeps the "#N"
    // visible; the alternative is it oscillating between two edges it can never both meet.
    [Fact]
    public void A_badge_larger_than_the_viewport_is_pinned_to_the_top_left()
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(-50, -50, 900, 700), Viewport);

        Assert.Equal(50.0, nudge.X, 6);
        Assert.Equal(50.0, nudge.Y, 6);
    }

    // A viewport that has not been laid out yet gives no basis for clamping, so nothing moves.
    [Fact]
    public void An_unmeasured_viewport_produces_no_nudge()
    {
        var nudge = BadgeGeometry.GetViewportNudge(new Rect(-100, -100, 60, 20), new Size(0, 0));

        Assert.Equal(0.0, nudge.X, 9);
        Assert.Equal(0.0, nudge.Y, 9);
    }
}
