using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A thumbnail's highlight layer is drawn at most MaxWidth wide, its proportions kept (2026-11-16, performance
// walkthrough): drawn at the thumbnail's own 800 wide, every pause in typing a filter drew one per matching thumbnail.
// It is stretched over its thumbnail as the thumbnail is, so the proportions are what must hold.
// ###########################################################################################
public sealed class ThumbnailLayerGeometryTests
{
    [Fact]
    public void A_wide_thumbnail_has_its_layer_drawn_at_most_the_widest_keeping_its_proportions()
    {
        PixelSize layer = ThumbnailLayerGeometry.SizeFor(new PixelSize(800, 557));

        Assert.Equal(ThumbnailLayerGeometry.MaxWidth, layer.Width);
        Assert.Equal(Math.Round(557 * (ThumbnailLayerGeometry.MaxWidth / 800.0)), layer.Height);
    }

    [Fact]
    public void A_thumbnail_no_wider_than_the_widest_has_its_layer_at_its_own_size()
    {
        Assert.Equal(new PixelSize(300, 200), ThumbnailLayerGeometry.SizeFor(new PixelSize(300, 200)));
        Assert.Equal(new PixelSize(ThumbnailLayerGeometry.MaxWidth, 90), ThumbnailLayerGeometry.SizeFor(new PixelSize(ThumbnailLayerGeometry.MaxWidth, 90)));
    }

    [Fact]
    public void A_thumbnail_of_no_size_has_a_layer_of_one_pixel()
    {
        Assert.Equal(new PixelSize(1, 1), ThumbnailLayerGeometry.SizeFor(new PixelSize(0, 0)));
        Assert.Equal(new PixelSize(ThumbnailLayerGeometry.MaxWidth, 1), ThumbnailLayerGeometry.SizeFor(new PixelSize(5000, 1)));
    }
}
