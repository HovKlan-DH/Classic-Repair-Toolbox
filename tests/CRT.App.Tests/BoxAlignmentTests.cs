using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A box stretched to end where a box in the next column ends (BoxAlignment; owner request, 2026-11-09: "Purpose or note" ending
// with "Other technical details").
// ###########################################################################################
public sealed class BoxAlignmentTests
{
    // The other box ends 90 under this one's top: this one is 90 tall at least.
    [Fact]
    public void A_box_is_as_tall_as_it_takes_to_end_with_the_other()
    {
        Assert.Equal(90, BoxAlignment.MinHeightToMeet(90, floor: 64));
        Assert.Equal(91, BoxAlignment.MinHeightToMeet(90.6, floor: 64));
    }

    // The other box ending higher - or above its top, as on a narrow window - leaves the box its own height.
    [Fact]
    public void Never_under_its_own_height()
    {
        Assert.Equal(64, BoxAlignment.MinHeightToMeet(40, floor: 64));
        Assert.Equal(64, BoxAlignment.MinHeightToMeet(-20, floor: 64));
        Assert.Equal(64, BoxAlignment.MinHeightToMeet(double.NaN, floor: 64));
        Assert.Equal(64, BoxAlignment.MinHeightToMeet(double.PositiveInfinity, floor: 64));
    }

    // Half a pixel either way is the same height - so a layout pass that changes nothing asks for none after it.
    [Fact]
    public void A_height_within_half_a_pixel_is_already_set()
    {
        Assert.True(BoxAlignment.IsAlready(90, 90.4));
        Assert.False(BoxAlignment.IsAlready(90, 91));
    }
}
