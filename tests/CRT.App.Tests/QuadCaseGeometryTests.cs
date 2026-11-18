using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// QuadCaseGeometry - a chip with pins on four sides, drawn square (owner request, 2026-11-18: "Drawing square packages ...
// there might come new hardware soon - e.g. Amiga 500 or Amiga 1000", and "It has a lot of SMD components"). The pins run
// counter-clockwise as every maker numbers them - a QFP's pin 1 at the top of the left side, a PLCC's in the middle of the
// top side - because a drawing numbering them any other way sends somebody probing the wrong leg of an Agnus.
// ###########################################################################################
public sealed class QuadCaseGeometryTests
{
    private static PinoutLayout Case(string package) => LogicChipGeometry.EmptyCaseFor(package)!.Case;

    // The pins of one side, in the order they stand: the left and right from the top down, the top and bottom from the left.
    private static int[] Side(PinoutLayout layout, PinSide side) =>
    [
        .. layout.Pins
            .Where(pin => pin.Side == side)
            .OrderBy(pin => side is PinSide.Left or PinSide.Right ? pin.Lead.Center.Y : pin.Lead.Center.X)
            .Select(pin => pin.Number),
    ];

    // ###########################################################################################
    // A QFP-44: pin 1 at the top of the left side, 1-11 down it, 12-22 along the bottom, 23-33 up the right, 34-44 back along
    // the top - the CX20185's numbering, and the 250477's layout's (11/12, 22/23, 33/34 at its corners).
    // ###########################################################################################
    [Fact]
    public void A_qfps_pins_run_counter_clockwise_from_the_top_of_the_left_side()
    {
        PinoutLayout layout = QuadCaseGeometryTests.Case("QFP-44");

        Assert.Equal(44, layout.Pins.Count);
        Assert.Equal(Enumerable.Range(1, 11), QuadCaseGeometryTests.Side(layout, PinSide.Left));
        Assert.Equal(Enumerable.Range(12, 11), QuadCaseGeometryTests.Side(layout, PinSide.Bottom));
        Assert.Equal(Enumerable.Range(23, 11).Reverse(), QuadCaseGeometryTests.Side(layout, PinSide.Right));
        Assert.Equal(Enumerable.Range(34, 11).Reverse(), QuadCaseGeometryTests.Side(layout, PinSide.Top));
        Assert.True(layout.IsQuad);
    }

    // ###########################################################################################
    // A PLCC-84 - the Amiga 500's Fat Agnus: pin 1 in the middle of the top side, 11 down to 1 and on from 84 to 75 along it,
    // 12-32 down the left, 33-53 along the bottom, 54-74 up the right. And a PLCC-44 the same way, 6 to 1 and 44 to 40.
    // ###########################################################################################
    [Fact]
    public void A_plccs_pins_run_counter_clockwise_from_the_middle_of_the_top_side()
    {
        PinoutLayout agnus = QuadCaseGeometryTests.Case("PLCC-84");

        Assert.Equal([.. Enumerable.Range(1, 11).Reverse(), .. Enumerable.Range(75, 10).Reverse()], QuadCaseGeometryTests.Side(agnus, PinSide.Top));
        Assert.Equal(Enumerable.Range(12, 21), QuadCaseGeometryTests.Side(agnus, PinSide.Left));
        Assert.Equal(Enumerable.Range(33, 21), QuadCaseGeometryTests.Side(agnus, PinSide.Bottom));
        Assert.Equal(Enumerable.Range(54, 21).Reverse(), QuadCaseGeometryTests.Side(agnus, PinSide.Right));

        PinoutLayout small = QuadCaseGeometryTests.Case("PLCC-44");

        Assert.Equal([6, 5, 4, 3, 2, 1, 44, 43, 42, 41, 40], QuadCaseGeometryTests.Side(small, PinSide.Top));
        Assert.Equal(Enumerable.Range(7, 11), QuadCaseGeometryTests.Side(small, PinSide.Left));

        // Pin 1 is the middle one of its side.
        PinoutPin first = small.Pins.Single(pin => pin.Number == 1);
        Assert.Equal(small.Body.Center.X, first.Lead.Center.X, 6);
    }

    // The pins are listed in number order, each once, its square outside the body on its side and its name's anchor beyond
    // the square - and no two squares touch.
    [Theory]
    [InlineData("QFP-32")]
    [InlineData("QFP-44")]
    [InlineData("PLCC-20")]
    [InlineData("PLCC-84")]
    [InlineData("LQFP-100")]
    [InlineData("QFP-208")]
    public void Every_pin_stands_once_outside_the_body_with_its_name_beyond_it(string package)
    {
        PinoutLayout layout = QuadCaseGeometryTests.Case(package);
        Rect body = layout.Body;

        Assert.Equal(Enumerable.Range(1, layout.Pins.Count), layout.Pins.Select(pin => pin.Number));
        Assert.Equal(layout.Pins.Count / 4, layout.Pins.Count(pin => pin.Side == PinSide.Top));

        foreach (PinoutPin pin in layout.Pins)
        {
            Assert.False(pin.Lead.Intersects(body), $"{package} pin {pin.Number} on the body");

            bool beyond = pin.Side switch
            {
                PinSide.Left => pin.Lead.Right < body.Left && pin.NameAnchor.X < pin.Lead.Left,
                PinSide.Right => pin.Lead.Left > body.Right && pin.NameAnchor.X > pin.Lead.Right,
                PinSide.Top => pin.Lead.Bottom < body.Top && pin.NameAnchor.Y < pin.Lead.Top,
                _ => pin.Lead.Top > body.Bottom && pin.NameAnchor.Y > pin.Lead.Bottom,
            };

            Assert.True(beyond, $"{package} pin {pin.Number}");
        }

        foreach (PinoutPin pin in layout.Pins)
        {
            Assert.DoesNotContain(layout.Pins, other => other != pin && other.Lead.Intersects(pin.Lead));
        }

        // The whole drawing holds them, the room for names on every side.
        Assert.True(layout.Pins.All(pin => new Rect(layout.Size).Contains(pin.Lead)));
        Assert.Equal(layout.Size.Width - body.Right, body.Left, 6);
        Assert.Equal(layout.Size.Height - body.Bottom, body.Top, 6);
    }

    // ###########################################################################################
    // A pin to point at and click, as on every case (PinoutGeometry.PinAt): the column over or under its square, the row
    // beside it - and nothing inside the body, nor two pins at one point.
    // ###########################################################################################
    [Fact]
    public void A_pin_is_found_on_its_column_or_row_and_never_in_the_body()
    {
        PinoutLayout layout = QuadCaseGeometryTests.Case("PLCC-84");

        foreach (PinoutPin pin in layout.Pins)
        {
            Assert.Same(pin, PinoutGeometry.PinAt(layout, pin.NumberAnchor));
            Assert.Same(pin, PinoutGeometry.PinAt(layout, pin.NameAnchor));
        }

        Assert.Null(PinoutGeometry.PinAt(layout, layout.Body.Center));

        // A corner of the drawing - beyond the ends of two sides - is no pin's.
        Assert.Null(PinoutGeometry.PinAt(layout, new Point(1, 1)));

        // No two pins' areas overlap.
        foreach (PinoutPin pin in layout.Pins)
        {
            Rect area = PinoutGeometry.AreaOf(layout, pin);

            Assert.DoesNotContain(layout.Pins, other => other != pin && PinoutGeometry.AreaOf(layout, other).Intersects(area.Deflate(0.5)));
        }
    }

    // The dot by pin 1 stands inside the body, beside pin 1; the cut corner is the top left's.
    [Fact]
    public void The_dot_by_pin_one_is_inside_the_body_beside_it_and_the_top_left_corner_is_cut()
    {
        foreach (string package in new[] { "QFP-44", "PLCC-44" })
        {
            PinoutLayout layout = QuadCaseGeometryTests.Case(package);
            PinoutPin first = layout.Pins.Single(pin => pin.Number == 1);
            Point dot = QuadCaseGeometry.FirstPinDotOf(layout);

            Assert.True(layout.Body.Deflate(QuadCaseGeometry.Bevel / 2).Contains(dot), package);
            Assert.True(first.Side == PinSide.Left ? dot.Y == first.Lead.Center.Y : dot.X == first.Lead.Center.X, package);
        }

        Rect body = QuadCaseGeometryTests.Case("QFP-44").Body;

        Assert.Equal(
            [new Point(body.Left + QuadCaseGeometry.Bevel, body.Top), body.TopRight, body.BottomRight, body.BottomLeft, new Point(body.Left, body.Top + QuadCaseGeometry.Bevel)],
            QuadCaseGeometry.OutlineOf(body));
    }

    // A number of three digits takes a wider square, and the pins stand further apart for it.
    [Fact]
    public void Three_digit_numbers_take_wider_squares()
    {
        PinoutLayout small = QuadCaseGeometryTests.Case("QFP-80");
        PinoutLayout large = QuadCaseGeometryTests.Case("QFP-100");

        Assert.Equal(LogicChipGeometry.Square, small.Pins[0].Lead.Width);
        Assert.Equal(LogicChipGeometry.Pitch, small.Pitch);
        Assert.True(large.Pins[0].Lead.Width > LogicChipGeometry.Square);
        Assert.True(large.Pitch > LogicChipGeometry.Pitch);
    }

    // What is no case of four sides is refused.
    [Fact]
    public void A_case_of_two_rows_is_no_case_of_four_sides()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => QuadCaseGeometry.Case(new ChipCase(ChipCaseKind.TwoRows, 40)));
        Assert.Throws<ArgumentOutOfRangeException>(() => QuadCaseGeometry.Case(new ChipCase(ChipCaseKind.Qfp, 42)));
    }
}
