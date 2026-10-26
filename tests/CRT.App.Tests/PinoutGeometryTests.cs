using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PinoutGeometry - where the pinout drawing puts a case's pins (CRT 4.0 step 2, 2026-10-25: "CRT
// draws the top view, with the notch and pin 1"). A DIP seen from above has pin 1 top left, the
// numbers running DOWN the left side and UP the right side, so its last pin sits opposite pin 1 -
// as on every datasheet; a drawing that put pin 21 top right would send someone probing the wrong
// leg of a 6510.
// ###########################################################################################
public sealed class PinoutGeometryTests
{
    [Fact]
    public void A_DIP_runs_down_the_left_and_up_the_right_with_the_last_pin_opposite_pin_1()
    {
        PinoutLayout layout = PinoutGeometry.DualInLine(40, wide: false);

        Assert.Equal(40, layout.Pins.Count);
        Assert.True(layout.IsDualInLine);

        PinoutPin first = layout.Pins[0];
        PinoutPin twentieth = layout.Pins[19];
        PinoutPin twentyFirst = layout.Pins[20];
        PinoutPin last = layout.Pins[39];

        Assert.Equal(PinSide.Left, first.Side);
        Assert.Equal(PinSide.Left, twentieth.Side);
        Assert.Equal(PinSide.Right, twentyFirst.Side);
        Assert.Equal(PinSide.Right, last.Side);

        // Pin 1 at the top, pin 20 at the bottom of the left side; 21 opposite 20, 40 opposite 1.
        Assert.True(first.Lead.Center.Y < twentieth.Lead.Center.Y);
        Assert.Equal(twentieth.Lead.Center.Y, twentyFirst.Lead.Center.Y, 3);
        Assert.Equal(first.Lead.Center.Y, last.Lead.Center.Y, 3);
    }

    // The leads leave the body at its edges, and every pin sits inside the drawing's size.
    [Fact]
    public void The_leads_leave_the_body_and_everything_fits_the_drawing()
    {
        PinoutLayout layout = PinoutGeometry.DualInLine(8, wide: false);
        var drawing = new Rect(layout.Size);

        Assert.All(layout.Pins, pin => Assert.True(drawing.Contains(pin.Lead), $"pin {pin.Number}"));
        Assert.Equal(layout.Body.Left, layout.Pins[0].Lead.Right, 3);
        Assert.Equal(layout.Body.Right, layout.Pins[7].Lead.Left, 3);
        Assert.True(drawing.Contains(layout.Body));

        // The notch in the middle of the top edge, the dot by pin 1.
        Assert.Equal(layout.Body.Center.X, layout.Notch.X, 3);
        Assert.Equal(layout.Body.Top, layout.Notch.Y, 3);
        Assert.True(layout.FirstPinDot.Y < layout.Pins[1].Lead.Center.Y);
        Assert.True(layout.FirstPinDot.X < layout.Body.Center.X);

        // Clear of pin 1's number, written against the edge (found on the first rendering: the dot sat
        // on the "1") - two digits at the drawing's size are well under 12 wide, the dot's radius is 3.
        Assert.True(layout.FirstPinDot.X - 3 > layout.Pins[0].NumberAnchor.X + 12);
    }

    [Fact]
    public void A_wide_DIP_has_a_wider_body_and_the_same_height()
    {
        PinoutLayout narrow = PinoutGeometry.DualInLine(24, wide: false);
        PinoutLayout wide = PinoutGeometry.DualInLine(24, wide: true);

        Assert.True(wide.Body.Width > narrow.Body.Width);
        Assert.Equal(narrow.Body.Height, wide.Body.Height, 3);
    }

    // A single row leaves the case at the bottom, pin 1 on the left.
    [Fact]
    public void A_single_row_runs_left_to_right_along_the_bottom()
    {
        PinoutLayout layout = PinoutGeometry.SingleRow(3);

        Assert.False(layout.IsDualInLine);
        Assert.Equal([1, 2, 3], layout.Pins.Select(pin => pin.Number));
        Assert.All(layout.Pins, pin => Assert.Equal(PinSide.Bottom, pin.Side));
        Assert.True(layout.Pins[0].Lead.X < layout.Pins[2].Lead.X);
        Assert.All(layout.Pins, pin => Assert.Equal(layout.Body.Bottom, pin.Lead.Top, 3));
        Assert.True(new Rect(layout.Size).Contains(layout.Pins[2].Lead));
    }

    // ###########################################################################################
    // Which drawing a package gets: a DIP by its name, TO-92 and TO-220 three pins in a row, a SIP as
    // many as its highest pin number - and nothing for a case CRT cannot count.
    // ###########################################################################################
    [Fact]
    public void A_package_gets_the_drawing_its_name_says_or_none()
    {
        Assert.Equal(40, PinoutGeometry.For("DIP-40", [])!.Pins.Count);
        Assert.True(PinoutGeometry.For("DIP-28 wide", [])!.Body.Width > PinoutGeometry.For("DIP-28", [])!.Body.Width);
        Assert.False(PinoutGeometry.For("TO-92", [])!.IsDualInLine);
        Assert.Equal(9, PinoutGeometry.For("SIP", ["1", "9", "4", "A"])!.Pins.Count);

        Assert.Null(PinoutGeometry.For("SIP", ["1"]));
        Assert.Null(PinoutGeometry.For("SIP", []));
        Assert.Null(PinoutGeometry.For("Other", ["1", "2"]));
        Assert.Null(PinoutGeometry.For("PLCC-44", []));
        Assert.Null(PinoutGeometry.For(null, []));
    }

    [Fact]
    public void A_DIP_with_an_odd_count_or_no_pins_is_refused()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PinoutGeometry.DualInLine(7, wide: false));
        Assert.Throws<ArgumentOutOfRangeException>(() => PinoutGeometry.DualInLine(0, wide: false));
        Assert.Throws<ArgumentOutOfRangeException>(() => PinoutGeometry.SingleRow(0));
    }
}
