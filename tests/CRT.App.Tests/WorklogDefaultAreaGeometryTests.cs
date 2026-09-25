using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// The area invented for a worklog entry that was never drawn on a schematic - the one created from
// the oscilloscope capture flow, which is stored with no area and parks as a corner pill.
//
// The failure this prevents is subtle and has no exception: ticking "Show marked area" on such an
// entry used to leave it with a zero-sized rect, which draws as nothing or as a hairline and can
// never be grabbed and dragged into place. The entry then looked broken with no way to fix it from
// the UI. So the tests below care about the square being VISIBLE, GRABBABLE and INSIDE the image.
public sealed class WorklogDefaultAreaGeometryTests
{
    // A drawn area is anything with real width and height. Everything else is "never drawn" - the
    // check is a threshold rather than a comparison to zero, because a rect that has been through a
    // JSON round-trip or a hand edit can carry a sliver that is still not grabbable.
    [Theory]
    [InlineData(0, 0, true)]
    [InlineData(0.4, 0.4, true)]
    [InlineData(-5, -5, true)]
    [InlineData(100, 0, true)]
    [InlineData(0, 100, true)]
    [InlineData(100, 80, false)]
    public void An_area_counts_as_unset_when_it_could_not_be_seen_or_grabbed(
        double width, double height, bool expectedUnset)
    {
        Assert.Equal(expectedUnset, WorklogDefaultAreaGeometry.IsUnset(new Rect(0, 0, width, height)));
    }

    // The whole square must land inside the image, or it reads as clipped - and a corner of it
    // would be unreachable.
    [Theory]
    [InlineData(4220, 2941)]
    [InlineData(800, 600)]
    [InlineData(2941, 4220)]
    public void The_default_area_sits_wholly_inside_the_image(double width, double height)
    {
        var area = WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(width, height));

        Assert.True(area.X >= 0);
        Assert.True(area.Y >= 0);
        Assert.True(area.Right <= width);
        Assert.True(area.Bottom <= height);
    }

    // Bottom-right specifically: the parked pills live in the TOP-right, so placing a new area at
    // the opposite corner means it can never be mistaken for one of them. Asserted by the square
    // being in the bottom-right quadrant, not merely somewhere.
    [Fact]
    public void The_default_area_is_placed_in_the_bottom_right_corner()
    {
        var area = WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(4000, 3000));

        Assert.True(area.X > 4000 / 2.0);
        Assert.True(area.Y > 3000 / 2.0);
    }

    // Big enough to see and to grab on a huge board scan, small enough not to blanket a small one.
    // A 4220px schematic is the real size of a board scan in this app's own data.
    [Fact]
    public void The_default_area_is_large_enough_to_grab_but_does_not_blanket_the_board()
    {
        var large = WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(4220, 2941));

        Assert.True(large.Width >= 24);
        Assert.True(large.Width < 2941 / 4.0);

        // Square, so it reads as a placeholder to be moved rather than as a deliberate shape.
        Assert.Equal(large.Width, large.Height, 3);
    }

    // A tiny image must not receive a square larger than itself, which the pixel MINIMUM would
    // otherwise force - that would push it off the board entirely.
    [Fact]
    public void A_tiny_image_still_receives_an_area_that_fits_inside_it()
    {
        var area = WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(10, 8));

        Assert.True(area.Width <= 10);
        Assert.True(area.Height <= 8);
        Assert.True(area.Right <= 10);
        Assert.True(area.Bottom <= 8);
    }

    // No board means nothing to place anything on, and the caller leaves the entry parked. Reported
    // as an empty rect rather than throwing, since a bitmap that has not loaded is a normal state.
    [Theory]
    [InlineData(0, 0)]
    [InlineData(0, 500)]
    [InlineData(500, 0)]
    public void An_image_with_no_usable_size_yields_no_area(double width, double height)
    {
        var area = WorklogDefaultAreaGeometry.BuildDefaultArea(new Size(width, height));

        Assert.True(area.Width < 1.0 || area.Height < 1.0);
    }

    // The single decision point: an entry that already has a drawn rectangle keeps it, so ticking
    // and unticking the box any number of times can never move a user's own marked area.
    [Fact]
    public void An_entry_that_already_has_an_area_keeps_it_exactly()
    {
        var existing = new Rect(120, 340, 200, 150);

        var resolved = WorklogDefaultAreaGeometry.ResolveAreaForShowing(existing, new Size(4000, 3000));

        Assert.Equal(existing, resolved);
    }

    [Fact]
    public void An_entry_with_no_area_receives_the_default_one()
    {
        var imageSize = new Size(4000, 3000);

        var resolved = WorklogDefaultAreaGeometry.ResolveAreaForShowing(default, imageSize);

        Assert.Equal(WorklogDefaultAreaGeometry.BuildDefaultArea(imageSize), resolved);
    }
}
