using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ThumbnailDecodeGeometry - how wide a schematic image is decoded for its thumbnail, and the
// signature that lets a board shown again reuse its thumbnails (code review, 2026-10-13: every
// board switch decoded every image full size, twice).
public sealed class ThumbnailDecodeGeometryTests
{
    [Theory]
    [InlineData(4220, 800, 800)]
    [InlineData(800, 800, 800)]
    [InlineData(640, 800, 640)]   // never upscaled - the old render capped its scale at 1 too
    [InlineData(0, 800, 1)]
    public void An_image_is_decoded_at_the_thumbnail_width_but_never_wider_than_itself(int width, int max, int expected)
    {
        Assert.Equal(expected, ThumbnailDecodeGeometry.DecodeWidth(width, max));
    }

    [Theory]
    [InlineData(4220, 2941, 800, 558)]
    [InlineData(1000, 1000, 800, 800)]
    [InlineData(800, 1, 800, 1)]
    [InlineData(0, 500, 800, 1)]
    public void The_decode_height_keeps_the_aspect_ratio(int width, int height, int decodeWidth, int expected)
    {
        Assert.Equal(expected, ThumbnailDecodeGeometry.DecodeHeight(width, height, decodeWidth));
    }

    private static string Signature(params (string, string, long, long)[] images) =>
        ThumbnailDecodeGeometry.SignatureOf(images);

    [Fact]
    public void The_same_images_unchanged_give_the_same_signature()
    {
        Assert.Equal(
            Signature(("Top", "a/top.png", 100, 5), ("Bottom", "a/bottom.png", 200, 6)),
            Signature(("Top", "a/top.png", 100, 5), ("Bottom", "a/bottom.png", 200, 6)));
    }

    // Anything that would show a different picture, or the same picture under another name or in
    // another place, is another set.
    [Fact]
    public void A_changed_file_name_order_or_missing_image_gives_another_signature()
    {
        string original = Signature(("Top", "a/top.png", 100, 5), ("Bottom", "a/bottom.png", 200, 6));

        Assert.NotEqual(original, Signature(("Top", "a/top.png", 101, 5), ("Bottom", "a/bottom.png", 200, 6)));
        Assert.NotEqual(original, Signature(("Top", "a/top.png", 100, 7), ("Bottom", "a/bottom.png", 200, 6)));
        Assert.NotEqual(original, Signature(("PCB top", "a/top.png", 100, 5), ("Bottom", "a/bottom.png", 200, 6)));
        Assert.NotEqual(original, Signature(("Bottom", "a/bottom.png", 200, 6), ("Top", "a/top.png", 100, 5)));
        Assert.NotEqual(original, Signature(("Top", "a/top.png", -1, 0), ("Bottom", "a/bottom.png", 200, 6)));
        Assert.NotEqual(original, Signature(("Top", "a/top.png", 100, 5)));
    }
}
