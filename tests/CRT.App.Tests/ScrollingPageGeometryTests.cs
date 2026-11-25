using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ScrollingPageGeometry - a page that fills its scroll viewer while there is room, and scrolls once there is not (the
// "Feedback" tab; owner report, 2026-11-25: on a low window its text box was squeezed to nothing and "Send feedback"
// lay below the tab).
// ###########################################################################################
public sealed class ScrollingPageGeometryTests
{
    // With room, the page is the viewport's height: the stretching part takes what the rest leaves, and nothing scrolls.
    [Fact]
    public void With_room_the_page_fills_the_viewport()
    {
        Assert.Equal(900, ScrollingPageGeometry.PageHeight(900, 500, 140));
    }

    // Too low, the page is just high enough for the rest and the stretching part's least - taller than the viewport, so
    // it scrolls, and the stretching part keeps its least.
    [Fact]
    public void Too_low_the_page_keeps_the_stretching_part_its_least_and_scrolls()
    {
        Assert.Equal(640, ScrollingPageGeometry.PageHeight(400, 500, 140));
    }

    // Exactly enough is the viewport's height - no scrolling for nothing.
    [Fact]
    public void Exactly_enough_room_fills_the_viewport_without_scrolling()
    {
        Assert.Equal(640, ScrollingPageGeometry.PageHeight(640, 500, 140));
    }

    // Before the scroll viewer is laid out it has no height: the page asks no more than it needs, which the next layout
    // pass corrects - never a height of nothing, nor an endless one.
    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_viewport_not_laid_out_yet_asks_only_what_the_page_needs(double viewport)
    {
        Assert.Equal(640, ScrollingPageGeometry.PageHeight(viewport, 500, 140));
    }

    // Parts measured as nothing count as nothing, never less.
    [Fact]
    public void Negative_parts_count_as_nothing()
    {
        Assert.Equal(300, ScrollingPageGeometry.PageHeight(300, -20, -5));
        Assert.Equal(140, ScrollingPageGeometry.PageHeight(0, -20, 140));
    }
}
