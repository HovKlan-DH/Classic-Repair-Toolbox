using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// How wide the worklog bar's drop-downs are (BarPickersGeometry; owner request, 2026-11-06: "have the two
// drop-downs resize, so it will show as much as the text in the drop-downs as possible"): each its line with room;
// short of room, shared as water fills - so a short line stays whole beside a long one cut short.
// ###########################################################################################
public sealed class BarPickersGeometryTests
{
    [Fact]
    public void With_room_each_gets_what_it_wants()
    {
        Assert.Equal([150.0, 400.0], BarPickersGeometry.Widths(800, [150, 400]));
        Assert.Equal([150.0, 400.0], BarPickersGeometry.Widths(550, [150, 400]));
    }

    // The device's "C64 (3)" whole, the workbook's long title given the rest - never both cut.
    [Fact]
    public void Short_of_room_the_shorter_stays_whole_and_the_longer_takes_the_rest()
    {
        Assert.Equal([150.0, 300.0], BarPickersGeometry.Widths(450, [150, 400]));
        Assert.Equal([300.0, 150.0], BarPickersGeometry.Widths(450, [400, 150]));
    }

    // Both wanting more than half: half each.
    [Fact]
    public void Both_long_share_evenly()
    {
        Assert.Equal([200.0, 200.0], BarPickersGeometry.Widths(400, [500, 300 + 200]));
        Assert.Equal([150.0, 150.0], BarPickersGeometry.Widths(300, [500, 250]));
    }

    [Fact]
    public void No_room_is_nothing_for_either()
    {
        Assert.Equal([0.0, 0.0], BarPickersGeometry.Widths(0, [150, 400]));
        Assert.Equal([0.0, 0.0], BarPickersGeometry.Widths(-20, [150, 400]));
        Assert.Empty(BarPickersGeometry.Widths(300, []));
    }
}
