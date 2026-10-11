using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// NoticeStackGeometry - how wide the notice cards over the tabs' bottom-right corner may be: at
// most half the tabs, so on the Schematics tab they stay over the thumbnail list (a full-width
// card in CRT's smallest window covered the "Global settings" and "Labels visible" panels).
// ###########################################################################################
public sealed class NoticeStackGeometryTests
{
    // In a big window the cap is Widest - half of 1700 would be a card wider than any sentence.
    [Theory]
    [InlineData(1700)]
    [InlineData(944)]
    public void Wide_tabs_give_the_widest_card(double tabsWidth)
    {
        Assert.Equal(NoticeStackGeometry.Widest, NoticeStackGeometry.MaxWidth(tabsWidth));
    }

    // In between, half the tabs less the edge margin - so the card's left edge never passes the
    // middle of the tabs.
    [Theory]
    [InlineData(614, 295)]
    [InlineData(800, 388)]
    public void Middling_tabs_give_half_their_width_less_the_margin(double tabsWidth, double expected)
    {
        double width = NoticeStackGeometry.MaxWidth(tabsWidth);

        Assert.Equal(expected, width);
        Assert.True(tabsWidth - NoticeStackGeometry.EdgeMargin - width >= tabsWidth / 2);
    }

    // Narrower than Narrowest a card is a column of single words, so the floor holds - and with
    // no width at all (before the first layout) the card still has one.
    [Theory]
    [InlineData(400)]
    [InlineData(0)]
    [InlineData(-5)]
    public void Narrow_tabs_never_squeeze_the_card_below_the_narrowest(double tabsWidth)
    {
        Assert.Equal(NoticeStackGeometry.Narrowest, NoticeStackGeometry.MaxWidth(tabsWidth));
    }

    // A width that is not a number would reach Avalonia as a MaxWidth it refuses.
    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_width_that_is_not_a_number_gives_the_widest_card(double tabsWidth)
    {
        Assert.Equal(NoticeStackGeometry.Widest, NoticeStackGeometry.MaxWidth(tabsWidth));
    }
}
