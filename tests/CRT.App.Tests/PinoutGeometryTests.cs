using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PinoutGeometry - where the pinout drawing puts a case's pins (CRT 4.0 step 2, 2026-10-25: "CRT
// draws the top view, with the notch and pin 1"): a single row's, and for every case a pin's area
// under the pointer and how wide the names are. A DIP is LogicChipGeometry's case since 2026-11-17
// (LogicChipGeometryTests) - its pin 1 top left, the numbers running DOWN the left side and UP the
// right side, so its last pin sits opposite pin 1, as on every datasheet; a drawing that put pin 21
// top right would send someone probing the wrong leg of a 6510.
// ###########################################################################################
public sealed class PinoutGeometryTests
{
    // ###########################################################################################
    // A pin to point at and click (owner request, 2026-11-01): a whole row of the drawing, its name
    // and its number on its square - not its square alone - and the rows never overlap, so a point is
    // one pin. Nothing inside the case is a pin's (code review, 2026-11-17): a badge drawn there for
    // another pin was taken for the pin of its row.
    // ###########################################################################################
    [Fact]
    public void A_dips_pin_is_found_anywhere_on_its_row_from_its_name_to_its_number()
    {
        PinoutLayout layout = LogicChipGeometry.EmptyCaseFor("DIP-8")!.Case;
        PinoutPin first = layout.Pins[0];
        PinoutPin last = layout.Pins[7];

        Assert.Same(first, PinoutGeometry.PinAt(layout, new Point(1, first.Lead.Center.Y)));
        Assert.Same(first, PinoutGeometry.PinAt(layout, first.NumberAnchor));
        Assert.Same(last, PinoutGeometry.PinAt(layout, new Point(layout.Size.Width - 1, last.Lead.Center.Y)));
        Assert.Same(last, PinoutGeometry.PinAt(layout, last.NumberAnchor));

        // The middle of the case is no pin, nor a point just inside its edge.
        Assert.Null(PinoutGeometry.PinAt(layout, layout.Body.Center));
        Assert.Null(PinoutGeometry.PinAt(layout, new Point(layout.Body.Left + 4, first.Lead.Center.Y)));
        Assert.Null(PinoutGeometry.PinAt(layout, new Point(layout.Body.Right - 4, last.Lead.Center.Y)));

        // Off the line two rows share, which the first of them takes.
        for (int y = 0; y < (int)layout.Size.Height; y++)
        {
            Assert.True(layout.Pins.Count(pin => PinoutGeometry.AreaOf(layout, pin).Contains(new Point(10, y + 0.5))) <= 1);
        }
    }

    [Fact]
    public void A_single_rows_pin_is_found_on_its_lead_and_number_below_the_case()
    {
        PinoutLayout layout = PinoutGeometry.SingleRow(3);
        PinoutPin second = layout.Pins[1];

        Assert.Same(second, PinoutGeometry.PinAt(layout, second.Lead.Center));
        Assert.Same(second, PinoutGeometry.PinAt(layout, new Point(second.NumberAnchor.X, second.NumberAnchor.Y + 4)));
        Assert.Null(PinoutGeometry.PinAt(layout, layout.Body.Center));
    }

    // A DIP is as wide as its names need, up to where a longer one is cut - the room is the pin list's.
    [Theory]
    [InlineData(0, 8)]
    [InlineData(40, 48)]
    [InlineData(40.2, 49)]
    [InlineData(112, 120)]
    [InlineData(500, 120)]
    public void A_dips_name_columns_fit_its_longest_name(double longest, double expected)
    {
        Assert.Equal(expected, PinoutGeometry.NameWidthFor(longest));
    }

    [Fact]
    public void Narrower_name_columns_move_the_case_and_keep_its_pins_row_by_row()
    {
        PinoutLayout usual = LogicChipGeometry.EmptyCaseFor("DIP-8")!.Case;
        PinoutLayout narrow = LogicChipGeometry.EmptyCaseFor("DIP-8", nameWidth: 48)!.Case;

        Assert.Equal(usual.Size.Width - (2 * (PinoutGeometry.NameWidth - 48)), narrow.Size.Width);
        Assert.Equal(usual.Size.Height, narrow.Size.Height);
        Assert.Equal(48 + LogicChipGeometry.Square + LogicChipGeometry.LeadGap, narrow.Body.Left);
        Assert.Equal(usual.Pins.Select(pin => pin.Lead.Center.Y), narrow.Pins.Select(pin => pin.Lead.Center.Y));
    }

    // Beside the drawing when the whole list fits there, every text whole (owner request, 2026-11-09: each column as wide
    // as its longest text) - a 432 wide list beside a 340 wide drawing and its gap of 12; with no drawing, or before the
    // width is known, there is nothing to go under.
    [Theory]
    [InlineData(900, 340, 432, true)]
    [InlineData(784, 340, 432, true)]
    [InlineData(783, 340, 432, false)]
    [InlineData(552, 340, 200, true)]
    [InlineData(300, 340, 200, false)]
    [InlineData(300, 0, 432, true)]
    [InlineData(0, 340, 432, true)]
    public void The_pin_list_stands_beside_the_drawing_while_it_fits_whole(double width, double drawing, double list, bool beside)
    {
        Assert.Equal(beside, PinoutGeometry.ListFitsBeside(width, drawing, list));
    }

    [Fact]
    public void A_DIP_runs_down_the_left_and_up_the_right_with_the_last_pin_opposite_pin_1()
    {
        PinoutLayout layout = LogicChipGeometry.EmptyCaseFor("DIP-40")!.Case;

        Assert.Equal(40, layout.Pins.Count);

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

    // A single row leaves the case at the bottom, pin 1 on the left, its dot over it inside the body.
    [Fact]
    public void A_single_row_runs_left_to_right_along_the_bottom()
    {
        PinoutLayout layout = PinoutGeometry.SingleRow(3);

        Assert.Equal([1, 2, 3], layout.Pins.Select(pin => pin.Number));
        Assert.True(layout.Body.Contains(layout.FirstPinDot));
        Assert.Equal(layout.Pins[0].Lead.Center.X, layout.FirstPinDot.X, 3);
        Assert.All(layout.Pins, pin => Assert.Equal(PinSide.Bottom, pin.Side));
        Assert.True(layout.Pins[0].Lead.X < layout.Pins[2].Lead.X);
        Assert.All(layout.Pins, pin => Assert.Equal(layout.Body.Bottom, pin.Lead.Top, 3));
        Assert.True(new Rect(layout.Size).Contains(layout.Pins[2].Lead));
    }

    // ###########################################################################################
    // Which single row a package is: TO-92 and TO-220 three pins, a SIP as many as its highest pin number - and nothing
    // for a DIP, which is LogicChipGeometry's case, or a case CRT cannot count.
    // ###########################################################################################
    [Fact]
    public void A_package_gets_the_single_row_its_name_says_or_none()
    {
        Assert.Equal(3, PinoutGeometry.For("TO-92", [])!.Pins.Count);
        Assert.Equal(9, PinoutGeometry.For("SIP", ["1", "9", "4", "A"])!.Pins.Count);

        Assert.Null(PinoutGeometry.For("DIP-40", []));
        Assert.Null(PinoutGeometry.For("DIP-28 wide", []));
        Assert.Null(PinoutGeometry.For("SIP", ["1"]));
        Assert.Null(PinoutGeometry.For("SIP", []));
        Assert.Null(PinoutGeometry.For("Other", ["1", "2"]));
        Assert.Null(PinoutGeometry.For("PLCC-44", []));
        Assert.Null(PinoutGeometry.For(null, []));
    }

    [Fact]
    public void A_case_with_an_odd_count_or_no_pins_is_refused()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => LogicChipGeometry.Case(7, 72));
        Assert.Throws<ArgumentOutOfRangeException>(() => LogicChipGeometry.Case(0, 72));
        Assert.Throws<ArgumentOutOfRangeException>(() => PinoutGeometry.SingleRow(0));
    }
}
