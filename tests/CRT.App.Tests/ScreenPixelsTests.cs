using Avalonia;
using Avalonia.Media;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ScreenPixels - what a color dropper reads (owner request, 2026-11-03; Controls/ColorDropperButton.cs):
// the color at a point of a picture of CRT's window. Pure: a headless display draws nothing to read.
// ###########################################################################################
public sealed class ScreenPixelsTests
{
    // Two pixels side by side over two more: red, green / blue, half-covered white.
    private static byte[] Pixels(bool isBgra)
    {
        byte[] Pixel(byte red, byte green, byte blue, byte alpha) =>
            isBgra ? [blue, green, red, alpha] : [red, green, blue, alpha];

        return [.. Pixel(255, 0, 0, 255), .. Pixel(0, 255, 0, 255), .. Pixel(0, 0, 255, 255), .. Pixel(128, 128, 128, 128)];
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void The_color_at_a_point_is_read_whichever_way_the_picture_holds_its_bytes(bool isBgra)
    {
        var pixels = new ScreenPixels(Pixels(isBgra), new PixelSize(2, 2), isBgra, 1.0);

        Assert.Equal(Color.FromRgb(255, 0, 0), pixels.ColorAt(new Point(0, 0)));
        Assert.Equal(Color.FromRgb(0, 255, 0), pixels.ColorAt(new Point(1.9, 0)));
        Assert.Equal(Color.FromRgb(0, 0, 255), pixels.ColorAt(new Point(0, 1)));
    }

    // A window's units are not its pixels: on a 150 % screen 1 is pixel 1.5 - the second pixel.
    [Fact]
    public void A_point_is_read_at_the_render_scale()
    {
        var pixels = new ScreenPixels(Pixels(isBgra: true), new PixelSize(2, 2), isBgra: true, 1.5);

        Assert.Equal(Color.FromRgb(255, 0, 0), pixels.ColorAt(new Point(0.5, 0)));
        Assert.Equal(Color.FromRgb(0, 255, 0), pixels.ColorAt(new Point(1, 0)));
        Assert.Null(pixels.ColorAt(new Point(1.34, 0)));
    }

    [Theory]
    [InlineData(-1, 0)]
    [InlineData(0, -1)]
    [InlineData(2, 0)]
    [InlineData(0, 2)]
    public void Off_the_picture_there_is_no_color(double x, double y)
    {
        var pixels = new ScreenPixels(Pixels(isBgra: true), new PixelSize(2, 2), isBgra: true, 1.0);

        Assert.Null(pixels.ColorAt(new Point(x, y)));
    }

    // The color seen is opaque - a window's picture holds a see-through pixel premultiplied.
    [Fact]
    public void A_see_through_pixel_is_read_as_its_color_and_an_empty_one_as_none()
    {
        var pixels = new ScreenPixels(Pixels(isBgra: true), new PixelSize(2, 2), isBgra: true, 1.0);

        Color seen = pixels.ColorAt(new Point(1, 1))!.Value;

        Assert.Equal(255, seen.A);
        Assert.Equal((255, 255, 255), (seen.R, seen.G, seen.B));

        var empty = new ScreenPixels([0, 0, 0, 0], new PixelSize(1, 1), isBgra: true, 1.0);

        Assert.Null(empty.ColorAt(new Point(0, 0)));
    }

    [Fact]
    public void A_color_is_said_as_its_hex()
    {
        Assert.Equal("#CD5C5C", ScreenPixels.HexOf(Colors.IndianRed));
        Assert.Equal("#000000", ScreenPixels.HexOf(Color.FromRgb(0, 0, 0)));
    }

    // Beside the pointer, never under it - and on the other side where it would leave the window.
    [Fact]
    public void The_lens_stays_beside_the_pointer_and_inside_the_window()
    {
        var lens = new Size(200, 50);
        var area = new Size(1000, 600);

        Assert.Equal(new Point(118, 118), ScreenPixels.LensPlace(new Point(100, 100), lens, area));
        Assert.Equal(new Point(782, 118), ScreenPixels.LensPlace(new Point(1000, 100), lens, area));
        Assert.Equal(new Point(118, 532), ScreenPixels.LensPlace(new Point(100, 600), lens, area));

        // A window too narrow for it either way: at its edge, never off it.
        Assert.Equal(new Point(0, 32), ScreenPixels.LensPlace(new Point(100, 100), lens, new Size(150, 80)));
    }
}
