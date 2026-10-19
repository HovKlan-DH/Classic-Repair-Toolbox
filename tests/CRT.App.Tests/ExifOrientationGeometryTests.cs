using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ExifOrientationGeometry - a picture turned upright by its EXIF orientation (code review,
// 2026-10-16). The picture below is three pixels wide and two tall, each pixel a letter:
//
//     a b c
//     d e f
//
// and each orientation is checked against what EXIF says it shows - e.g. 6, "row 0 is the right
// side, column 0 the top", is the picture turned a quarter clockwise.
// ###########################################################################################
public sealed class ExifOrientationGeometryTests
{
    // Four bytes a pixel, the letter in every byte, so any byte order reads the same.
    private static byte[] Pixels(string letters) =>
        letters.SelectMany(letter => Enumerable.Repeat((byte)letter, 4)).ToArray();

    private static string Letters(byte[] pixels) =>
        new(pixels.Where((_, index) => index % 4 == 0).Select(value => (char)value).ToArray());

    [Theory]
    [InlineData(1, 3, 2, "abcdef")]
    [InlineData(2, 3, 2, "cbafed")]
    [InlineData(3, 3, 2, "fedcba")]
    [InlineData(4, 3, 2, "defabc")]
    [InlineData(5, 2, 3, "adbecf")]
    [InlineData(6, 2, 3, "daebfc")]
    [InlineData(7, 2, 3, "fcebda")]
    [InlineData(8, 2, 3, "cfbead")]
    public void Each_orientation_turns_the_picture_as_EXIF_says(int orientation, int width, int height, string rows)
    {
        (byte[] pixels, int newWidth, int newHeight) = ExifOrientationGeometry.Apply(Pixels("abcdef"), 3, 2, orientation);

        Assert.Equal((width, height), (newWidth, newHeight));
        Assert.Equal(rows, Letters(pixels));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(9)]
    public void An_orientation_EXIF_does_not_have_leaves_the_picture_as_it_is(int orientation)
    {
        (byte[] pixels, int width, int height) = ExifOrientationGeometry.Apply(Pixels("abcdef"), 3, 2, orientation);

        Assert.Equal((3, 2), (width, height));
        Assert.Equal("abcdef", Letters(pixels));
    }

    // Turned a quarter one way and a quarter back, a picture is what it was - every pixel, not
    // only the size.
    [Fact]
    public void A_quarter_clockwise_and_a_quarter_back_is_the_picture_again()
    {
        byte[] original = Pixels("abcdefghijkl");

        (byte[] turned, int width, int height) = ExifOrientationGeometry.Apply(original, 4, 3, 6);
        (byte[] back, int backWidth, int backHeight) = ExifOrientationGeometry.Apply(turned, width, height, 8);

        Assert.Equal((4, 3), (backWidth, backHeight));
        Assert.Equal(original, back);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(4, false)]
    [InlineData(5, true)]
    [InlineData(6, true)]
    [InlineData(8, true)]
    public void Orientations_5_to_8_swap_the_sides(int orientation, bool swaps)
    {
        Assert.Equal(swaps, ExifOrientationGeometry.SwapsSides(orientation));
        Assert.Equal(swaps ? (3024, 4032) : (4032, 3024), ExifOrientationGeometry.OrientedSize(4032, 3024, orientation));
    }

    [Fact]
    public void Pixels_that_are_not_the_size_given_are_refused()
    {
        Assert.Throws<ArgumentException>(() => ExifOrientationGeometry.Apply(new byte[20], 3, 2, 6));
    }
}
