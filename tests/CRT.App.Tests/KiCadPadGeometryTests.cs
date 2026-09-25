using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadPadGeometry - pad shape classification and the KiCad-angle to screen-angle
// conversion the trace overlay uses.
//
// This logic exists because pad rotation used to be dropped entirely: a rectangular pad was always
// drawn axis-aligned from its width and height, so a pad rotated 90 degrees came out horizontal
// when it should be vertical. It is only visible on non-square pads, which is why it survived on
// every shipped board - each has only a handful of them among a thousand or more round pads.
public sealed class KiCadPadGeometryTests
{
    // ----------------------------------------------------------------- shape classification

    [Theory]
    [InlineData("rect")]
    [InlineData("roundrect")]
    [InlineData("trapezoid")]
    [InlineData("RECT")]
    [InlineData("  rect  ")]
    public void Rectangular_pad_shapes_are_drawn_as_rectangles(string shape)
    {
        Assert.True(KiCadPadGeometry.IsRectangularShape(shape));
    }

    [Theory]
    [InlineData("circle")]
    [InlineData("oval")]
    [InlineData("custom")]
    [InlineData("")]
    [InlineData(null)]
    public void Every_other_pad_shape_falls_back_to_an_ellipse(string? shape)
    {
        // An unknown or missing shape must still draw something, and a round pad is the safe
        // approximation - it never claims an orientation the file did not state.
        Assert.False(KiCadPadGeometry.IsRectangularShape(shape));
    }

    // ------------------------------------------------------------------- screen rotation

    [Fact]
    public void A_kicad_angle_is_negated_because_avalonia_rotates_the_other_way()
    {
        // KiCad measures a positive pad angle counter-clockwise as drawn; Avalonia's
        // Matrix.CreateRotation is clockwise in its Y-down device space.
        Assert.Equal(270, KiCadPadGeometry.ResolveScreenRotationDegrees(90, false, false), precision: 9);
    }

    [Fact]
    public void A_single_mirrored_axis_reverses_the_direction_of_rotation()
    {
        // Viewing a board through a mirrored calibration box reflects the geometry, and a
        // reflection swaps clockwise for counter-clockwise.
        Assert.Equal(90, KiCadPadGeometry.ResolveScreenRotationDegrees(90, true, false), precision: 9);
        Assert.Equal(90, KiCadPadGeometry.ResolveScreenRotationDegrees(90, false, true), precision: 9);
    }

    [Fact]
    public void Mirroring_both_axes_restores_the_direction_of_rotation()
    {
        // Two reflections compose into a 180-degree turn, which preserves handedness.
        Assert.Equal(270, KiCadPadGeometry.ResolveScreenRotationDegrees(90, true, true), precision: 9);
    }

    [Fact]
    public void A_negative_kicad_angle_wraps_into_the_positive_range()
    {
        // KiCad writes -90 as readily as 270 for the same placement, so both must resolve alike.
        Assert.Equal(
            KiCadPadGeometry.ResolveScreenRotationDegrees(270, false, false),
            KiCadPadGeometry.ResolveScreenRotationDegrees(-90, false, false),
            precision: 9);
    }

    [Fact]
    public void An_unrotated_pad_stays_unrotated()
    {
        Assert.Equal(0, KiCadPadGeometry.ResolveScreenRotationDegrees(0, false, false), precision: 9);
        Assert.Equal(0, KiCadPadGeometry.ResolveScreenRotationDegrees(360, false, false), precision: 9);
    }

    [Fact]
    public void A_malformed_angle_is_treated_as_no_rotation()
    {
        // Contributed board files are not always well formed; a NaN must not propagate into a
        // render transform, where it would erase the pad rather than misplace it.
        Assert.Equal(0, KiCadPadGeometry.ResolveScreenRotationDegrees(double.NaN, false, false), precision: 9);
        Assert.Equal(0, KiCadPadGeometry.ResolveScreenRotationDegrees(double.PositiveInfinity, true, false), precision: 9);
    }

    // ----------------------------------------------------------------- axis alignment

    [Theory]
    [InlineData(0)]
    [InlineData(180)]
    [InlineData(360)]
    [InlineData(-180)]
    public void Half_turns_count_as_axis_aligned_so_the_renderer_can_skip_them(double degrees)
    {
        // Both shapes a pad is drawn with - a rectangle and an ellipse - are symmetric about their
        // centre, so a half turn is invisible. 180-degree placements are common on a real board,
        // so skipping them is worth doing.
        Assert.True(KiCadPadGeometry.IsAxisAligned(degrees));
    }

    [Theory]
    [InlineData(90)]
    [InlineData(270)]
    [InlineData(45)]
    [InlineData(-90)]
    public void A_quarter_turn_is_not_axis_aligned(double degrees)
    {
        Assert.False(KiCadPadGeometry.IsAxisAligned(degrees));
    }

    // ----------------------------------------------------------------------- normalising

    [Theory]
    [InlineData(0, 0)]
    [InlineData(90, 90)]
    [InlineData(-90, 270)]
    [InlineData(450, 90)]
    [InlineData(-450, 270)]
    public void Angles_normalise_into_the_zero_to_three_sixty_range(double input, double expected)
    {
        Assert.Equal(expected, KiCadPadGeometry.NormalizeDegrees(input), precision: 9);
    }
}
