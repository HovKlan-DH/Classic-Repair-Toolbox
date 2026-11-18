using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PinBadgeGeometry - a pin's number on its pad while SHIFT is held (owner request, 2026-11-17: "a pin number badge for all
// component pins ... placed directly on top of the pin").
// ###########################################################################################
public sealed class PinBadgeGeometryTests
{
    // As tall as the pad is on screen, but never too small to read nor larger than a badge needs to be.
    [Theory]
    [InlineData(4, PinBadgeGeometry.LeastHeight)]
    [InlineData(15, 15)]
    [InlineData(80, PinBadgeGeometry.MostHeight)]
    [InlineData(double.NaN, PinBadgeGeometry.LeastHeight)]
    public void A_badge_is_as_tall_as_its_pad_within_its_least_and_most(double padOnScreen, double height)
    {
        Assert.Equal(height, PinBadgeGeometry.HeightFor(padOnScreen));
    }

    // Centered on the pad, as wide as its number with room each side, never narrower than it is tall.
    [Fact]
    public void A_badge_stands_centered_on_its_pad()
    {
        Rect one = PinBadgeGeometry.AreaOf(new Point(100, 50), 14, numberWidth: 5);
        Assert.Equal(new Point(100, 50), one.Center);
        Assert.Equal(14, one.Width);

        Rect wide = PinBadgeGeometry.AreaOf(new Point(100, 50), 14, numberWidth: 20);
        Assert.Equal(20 + (2 * PinBadgeGeometry.SideRoom), wide.Width);
        Assert.Equal(new Point(100, 50), wide.Center);
    }

    // Pin 1 where the footprint has one; else the first by KiCad's order - as pin 1's orange pad is chosen.
    [Fact]
    public void The_first_pin_is_pin_1_else_the_first_designator()
    {
        Assert.Equal("1", PinBadgeGeometry.FirstPinOf(["3", " 1 ", "2"]));
        Assert.Equal("A1", PinBadgeGeometry.FirstPinOf(["B2", "A2", "A1"]));
        Assert.Equal("2", PinBadgeGeometry.FirstPinOf(["10", "2"]));
        Assert.Equal(string.Empty, PinBadgeGeometry.FirstPinOf([" ", ""]));
    }

    // ###########################################################################################
    // Pin 1's orange pad framed (code review, 2026-11-18: its maths sat untested in the control): at any zoom the frame's
    // 1-wide line stands just outside the orange, which ends half the overlay's outline outside the pad's size at that zoom.
    // ###########################################################################################
    [Theory]
    [InlineData(1.0)]
    [InlineData(0.25)]
    [InlineData(6.0)]
    public void Pin_1s_frame_stands_just_outside_its_orange_pad(double scale)
    {
        var centre = new Point(100, 50);
        var pad = new Size(1.6, 2.4);

        Rect frame = PinBadgeGeometry.FirstPinFrame(centre, pad, scale);

        double orangeOut = KiCadPadGeometry.OutlineWidth / 2 * scale;
        Assert.Equal(centre.X, frame.Center.X, precision: 6);
        Assert.Equal(centre.Y, frame.Center.Y, precision: 6);
        Assert.Equal((pad.Width * scale) + (2 * orangeOut) + 1, frame.Width, precision: 6);
        Assert.Equal((pad.Height * scale) + (2 * orangeOut) + 1, frame.Height, precision: 6);
    }
}
