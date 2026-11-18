using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ImageFrameGeometry - room round the "Schematics" tab's image so its badges cover none of it (owner request, 2026-11-13),
// and its board in the same place whether it is cropped or padded (owner request, 2026-11-17: "no matter if I have
// inserted empty spacing or not, then the actual PCB should be displayed same place"). The owner's two images of the
// Open128-2 board are the cases: "PCB Top", with room of its own all round, and "PCB Top CROPPED", cut to the board.
// ###########################################################################################
public sealed class ImageFrameGeometryTests
{
    private const uint Clear = 0x00FFFFFF;      // transparent - as both of the owner's images are outside the board
    private const uint Board = 0xFFDDF2EE;      // the board's light green
    private const uint Trace = 0xFF7FB8B0;

    // An image `width` by `height` with an empty border of `top`, `left`, `right`, `bottom` - and on the board's first
    // row and column one trace pixel, as a board always has something on each line across it.
    private static (Func<int, uint[]> Row, Func<int, uint[]> Column) Image(int width, int height, int top, int left, int right, int bottom, uint empty = Clear)
    {
        uint At(int x, int y)
        {
            if (y < top || y >= height - bottom || x < left || x >= width - right)
            {
                return empty;
            }

            return x == left + 1 || y == top + 1 ? Trace : Board;
        }

        return (y => Enumerable.Range(0, width).Select(x => At(x, y)).ToArray(),
                x => Enumerable.Range(0, height).Select(y => At(x, y)).ToArray());
    }

    [Fact]
    public void An_images_own_empty_border_is_counted_from_each_side()
    {
        (var row, var column) = Image(300, 200, top: 20, left: 10, right: 7, bottom: 15);

        Assert.Equal(new Thickness(10, 20, 7, 15), ImageFrameGeometry.BlankBorder(300, 200, row, column));
    }

    // Cut to the board: nothing of its own - and a board's edge row that is all one color but for one pixel is board.
    [Fact]
    public void An_image_cut_to_its_board_has_no_border_of_its_own()
    {
        (var row, var column) = Image(300, 200, top: 0, left: 0, right: 0, bottom: 0);

        Assert.Equal(default, ImageFrameGeometry.BlankBorder(300, 200, row, column));
    }

    // A plain white margin counts as a transparent one does, and a JPEG's few shades of noise in it still count as one
    // color; never more than a third of the image from a side.
    [Fact]
    public void A_plain_margin_with_a_little_noise_counts_and_never_past_a_third()
    {
        uint[] Noisy(int length) => Enumerable.Range(0, length).Select(i => i % 3 == 0 ? 0xFFF8F8F8u : 0xFFFFFFFFu).ToArray();

        Thickness noisy = ImageFrameGeometry.BlankBorder(90, 90, _ => Noisy(90), _ => Noisy(90));
        Assert.Equal(new Thickness(30, 30, 30, 30), noisy);

        (var row, var column) = Image(300, 200, top: 12, left: 0, right: 0, bottom: 0, empty: 0xFFFFFFFF);
        Assert.Equal(12, ImageFrameGeometry.BlankBorder(300, 200, row, column).Top);
    }

    [Fact]
    public void Two_transparent_pixels_are_alike_whatever_colors_they_carry()
    {
        Assert.True(ImageFrameGeometry.Alike(0x00FFFFFF, 0x00000000));
        Assert.True(ImageFrameGeometry.Alike(0xFFFFFFFF, 0xFFF4F4F4));
        Assert.False(ImageFrameGeometry.Alike(0xFFFFFFFF, 0xFFDDF2EE));
        Assert.False(ImageFrameGeometry.Alike(0x00FFFFFF, 0xFFFFFFFF));
    }

    // Where the image's CONTENT - the image less its own empty border - is drawn, fitted inside `frame` as the tab fits it.
    private static Rect BoardOnScreen(Size box, PixelSize image, Thickness blank)
    {
        Thickness frame = ImageFrameGeometry.Frame(box, image, blank, ImageFrameGeometry.Wanted);
        Rect drawn = RectGeometry.GetImageContentRect(box, image, frame);
        double scale = drawn.Width / image.Width;

        return new Rect(
            drawn.Left + (scale * blank.Left),
            drawn.Top + (scale * blank.Top),
            scale * (image.Width - blank.Left - blank.Right),
            scale * (image.Height - blank.Top - blank.Bottom));
    }

    // ###########################################################################################
    // *** CROPPED OR PADDED, THE BOARD STANDS IN THE SAME PLACE (owner request, 2026-11-17: "no matter if I have inserted
    // empty spacing or not, then the actual PCB should be displayed same place"). *** The owner's two images of the
    // Open128-2 as measured: "PCB Top" (5080 x 3604, empty 336 above, 211 below, 89 left, 73 right) and "PCB Top CROPPED"
    // (4903 x 3058, empty 0 above, 2 below, 1 left, 0 right). Until then the padded one kept its own room, and showed its
    // board lower and smaller.
    // ###########################################################################################
    [Theory]
    [InlineData(1400, 850)]
    [InlineData(990, 750)]
    public void The_owners_padded_and_cropped_images_show_their_board_in_the_same_place(double width, double height)
    {
        var box = new Size(width, height);
        Rect padded = BoardOnScreen(box, new PixelSize(5080, 3604), new Thickness(89, 336, 73, 211));
        Rect cropped = BoardOnScreen(box, new PixelSize(4903, 3058), new Thickness(1, 0, 0, 2));

        Assert.Equal(cropped.Left, padded.Left, precision: 3);
        Assert.Equal(cropped.Top, padded.Top, precision: 3);
        Assert.Equal(cropped.Width, padded.Width, tolerance: cropped.Width * 0.01);
        Assert.Equal(cropped.Height, padded.Height, tolerance: cropped.Height * 0.01);

        // Under both rows of badges, the room wanted at the sides, and wholly in the panel.
        Assert.Equal(ImageFrameGeometry.TopRoom, padded.Top, precision: 3);
        Assert.Equal(ImageFrameGeometry.SideRoom, padded.Left, precision: 3);
        Assert.True(padded.Right <= width - ImageFrameGeometry.SideRoom + 0.001);
        Assert.True(padded.Bottom <= height - ImageFrameGeometry.SideRoom + 0.001);
    }

    // ###########################################################################################
    // The content of an image as drawn - what the first view's panning is judged by (code review, 2026-11-18): the owner's
    // padded "PCB Top" reaches past the panel at the first view, while its board stands wholly inside it, so nothing is
    // there to pan to.
    // ###########################################################################################
    [Fact]
    public void A_padded_images_content_stands_inside_the_panel_though_the_image_reaches_past_it()
    {
        var box = new Size(990, 750);
        var image = new PixelSize(5080, 3604);
        var blank = new Thickness(89, 336, 73, 211);

        Thickness frame = ImageFrameGeometry.Frame(box, image, blank, ImageFrameGeometry.Wanted);
        Rect drawn = RectGeometry.GetImageContentRect(box, image, frame);
        Rect content = ImageFrameGeometry.ContentWithin(drawn, image, blank);

        Assert.True(drawn.Top < 0 || drawn.Height > box.Height || drawn.Width > box.Width - (2 * ImageFrameGeometry.SideRoom), $"{drawn} drawn");

        Rect expected = BoardOnScreen(box, image, blank);
        Assert.Equal(expected.Left, content.Left, precision: 3);
        Assert.Equal(expected.Top, content.Top, precision: 3);
        Assert.Equal(expected.Width, content.Width, precision: 3);
        Assert.Equal(expected.Height, content.Height, precision: 3);
        Assert.True(content.Width <= box.Width && content.Height <= box.Height, $"{content} in {box}");
    }

    // No border, or an image that is all border, is the image as drawn - as Frame fits it.
    [Fact]
    public void An_image_with_no_content_apart_from_its_border_is_its_whole_self()
    {
        var drawn = new Rect(10, 20, 300, 200);

        Assert.Equal(drawn, ImageFrameGeometry.ContentWithin(drawn, new PixelSize(300, 200), default));
        Assert.Equal(drawn, ImageFrameGeometry.ContentWithin(drawn, new PixelSize(100, 100), new Thickness(50, 50, 50, 50)));
        Assert.Equal(drawn, ImageFrameGeometry.ContentWithin(drawn, new PixelSize(0, 0), new Thickness(5)));
    }

    // The room above clears the second row of badges - a component or pin pointed at - 44 down and 26 high, with room to spare
    // (owner request, 2026-11-17: "displayed below both rows of badges + with a little margin, so it does not look cramped").
    [Fact]
    public void The_room_above_clears_both_rows_of_badges()
    {
        Assert.True(ImageFrameGeometry.TopRoom >= 44 + 26 + 8, $"{ImageFrameGeometry.TopRoom} above");
    }

    // An image's own border wider than the room wanted reaches past it - less than nothing added - and the image fitted inside
    // that room is exactly its board in the room wanted.
    [Fact]
    public void A_border_wider_than_the_room_wanted_reaches_past_it()
    {
        var box = new Size(1000, 800);
        var image = new PixelSize(1000, 700);
        var blank = new Thickness(100, 150, 100, 50);

        Thickness frame = ImageFrameGeometry.Frame(box, image, blank, ImageFrameGeometry.Wanted);

        Assert.True(frame.Left < 0 && frame.Top < 0, $"{frame} round it");

        Rect board = BoardOnScreen(box, image, blank);
        Assert.Equal(ImageFrameGeometry.SideRoom, board.Left, precision: 3);
        Assert.Equal(ImageFrameGeometry.TopRoom, board.Top, precision: 3);

        // Its board is fitted as large as the room allows: as wide as it, or as tall.
        double roomWidth = 1000 - (2 * ImageFrameGeometry.SideRoom);
        double roomHeight = 800 - ImageFrameGeometry.TopRoom - ImageFrameGeometry.SideRoom;
        Assert.True(Math.Abs(board.Width - roomWidth) < 0.01 || Math.Abs(board.Height - roomHeight) < 0.01, $"{board} in the room");
    }

    // An image with no border of its own gets the room wanted all round, as before.
    [Fact]
    public void An_image_cut_to_its_board_gets_the_room_wanted()
    {
        Thickness frame = ImageFrameGeometry.Frame(new Size(1400, 850), new PixelSize(4903, 3058), default, ImageFrameGeometry.Wanted);

        Assert.Equal(ImageFrameGeometry.Wanted, frame);
    }

    // An image that is all border, as far as can be told, is fitted whole.
    [Fact]
    public void An_image_that_is_all_border_is_fitted_whole()
    {
        Thickness frame = ImageFrameGeometry.Frame(new Size(1000, 800), new PixelSize(100, 100), new Thickness(50, 50, 50, 50), ImageFrameGeometry.Wanted);

        Assert.Equal(ImageFrameGeometry.Wanted, frame);
    }

    [Fact]
    public void No_image_or_no_room_left_is_no_frame()
    {
        Assert.Equal(default, ImageFrameGeometry.Frame(new Size(1400, 850), new PixelSize(0, 0), default, ImageFrameGeometry.Wanted));
        Assert.Equal(default, ImageFrameGeometry.Frame(new Size(20, 30), new PixelSize(100, 100), default, ImageFrameGeometry.Wanted));

        // A frame leaving no room is no frame to the fit either.
        Assert.Equal(
            RectGeometry.GetImageContentRect(new Size(20, 30), new PixelSize(100, 100)),
            RectGeometry.GetImageContentRect(new Size(20, 30), new PixelSize(100, 100), ImageFrameGeometry.Wanted));
    }

    // ###########################################################################################
    // Read a BAND of rows or columns at a time (code review, 2026-11-13: one read per row and column, thousands on a large
    // scan, before the image showed): the same border as line by line, in a handful of reads - bands not on a whole Band
    // of lines from the far side too.
    // ###########################################################################################
    [Fact]
    public void Read_in_bands_it_finds_the_same_border_in_a_handful_of_reads()
    {
        const int width = 450;
        const int height = 300;
        (var row, _) = Image(width, height, top: 70, left: 10, right: 90, bottom: 3);
        uint[][] rows = Enumerable.Range(0, height).Select(row).ToArray();
        int reads = 0;

        Thickness border = ImageFrameGeometry.BlankBorder(width, height, rect =>
        {
            reads++;
            return Enumerable.Range(rect.Y, rect.Height).SelectMany(y => rows[y].Skip(rect.X).Take(rect.Width)).ToArray();
        });

        Assert.Equal(new Thickness(10, 70, 90, 3), border);
        Assert.True(reads <= 8, $"{reads} reads");
    }

    // ###########################################################################################
    // A GRAYSCALE image (code review, 2026-11-13): it decodes as Gray8, whose pixels Avalonia 12.1.4 will not copy, so its
    // own border was never measured and the whole room was added over it (eleven shipped images, C128 250477's sheets
    // among them). Read from its file as 32 bits a pixel, it is measured as any other.
    // ###########################################################################################
    [Theory]
    [InlineData(SkiaSharp.SKColorType.Gray8)]
    [InlineData(SkiaSharp.SKColorType.Rgba8888)]
    public void A_grayscale_image_files_own_border_is_measured(SkiaSharp.SKColorType stored)
    {
        string file = Path.Combine(Path.GetTempPath(), $"crt-gray-{Guid.NewGuid():N}.png");

        try
        {
            using (var gray = new SkiaSharp.SKBitmap(new SkiaSharp.SKImageInfo(120, 80, stored, SkiaSharp.SKAlphaType.Opaque)))
            {
                gray.Erase(SkiaSharp.SKColors.White);

                for (int y = 12; y < 80 - 5; y++)
                {
                    for (int x = 7; x < 120 - 20; x++)
                    {
                        gray.SetPixel(x, y, x == 7 || y == 12 ? SkiaSharp.SKColors.Black : SkiaSharp.SKColors.Gray);
                    }
                }

                using SkiaSharp.SKData png = gray.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                File.WriteAllBytes(file, png.ToArray());
            }

            Assert.Equal(new Thickness(7, 12, 20, 5), ImageFrameGeometry.BlankBorderOfFile(file));
        }
        finally
        {
            File.Delete(file);
        }
    }

    // ###########################################################################################
    // A grayscale THUMBNAIL standing in for its picture (code review, 2026-11-18): it has no file of its own, so it is read
    // encoded, from memory - measured as no border, the stand-in stood elsewhere than the picture once it came.
    // ###########################################################################################
    [Theory]
    [InlineData(SkiaSharp.SKColorType.Gray8)]
    [InlineData(SkiaSharp.SKColorType.Rgba8888)]
    public void An_encoded_thumbnails_own_border_is_measured(SkiaSharp.SKColorType stored)
    {
        using var gray = new SkiaSharp.SKBitmap(new SkiaSharp.SKImageInfo(60, 40, stored, SkiaSharp.SKAlphaType.Opaque));
        gray.Erase(SkiaSharp.SKColors.White);

        for (int y = 6; y < 40 - 3; y++)
        {
            for (int x = 4; x < 60 - 10; x++)
            {
                gray.SetPixel(x, y, x == 4 || y == 6 ? SkiaSharp.SKColors.Black : SkiaSharp.SKColors.Gray);
            }
        }

        using SkiaSharp.SKData png = gray.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
        using var encoded = new MemoryStream(png.ToArray());

        Assert.Equal(new Thickness(4, 6, 10, 3), ImageFrameGeometry.BlankBorderOfEncoded(encoded));
        Assert.Equal(default, ImageFrameGeometry.BlankBorderOfEncoded(new MemoryStream([1, 2, 3])));
    }

    // ###########################################################################################
    // KEPT by the file's stamp once it has settled (code review, 2026-11-13: read again at every switch to the image): the
    // same place, length and write time give the border measured before - bytes swapped under the same stamp are not
    // read - while a file just written is read every time, as two writes in one clock tick share a stamp.
    // ###########################################################################################
    [Fact]
    public void An_image_files_border_is_kept_by_its_settled_stamp()
    {
        string file = Path.Combine(Path.GetTempPath(), $"crt-kept-{Guid.NewGuid():N}.png");

        try
        {
            using (var image = new SkiaSharp.SKBitmap(new SkiaSharp.SKImageInfo(60, 40, SkiaSharp.SKColorType.Gray8, SkiaSharp.SKAlphaType.Opaque)))
            {
                image.Erase(SkiaSharp.SKColors.White);

                for (int y = 4; y < 40; y++)
                {
                    for (int x = 3; x < 60; x++)
                    {
                        image.SetPixel(x, y, x == 3 || y == 4 ? SkiaSharp.SKColors.Black : SkiaSharp.SKColors.Gray);
                    }
                }

                using SkiaSharp.SKData png = image.Encode(SkiaSharp.SKEncodedImageFormat.Png, 100);
                File.WriteAllBytes(file, png.ToArray());
            }

            DateTime settled = DateTime.UtcNow.AddMinutes(-1);
            File.SetLastWriteTimeUtc(file, settled);

            Thickness measured = ImageFrameGeometry.BlankBorderOfFile(file);
            Assert.Equal(new Thickness(3, 4, 0, 0), measured);

            // The same length and write time, bytes no decoder reads: the border kept.
            File.WriteAllBytes(file, new byte[new FileInfo(file).Length]);
            File.SetLastWriteTimeUtc(file, settled);
            Assert.Equal(measured, ImageFrameGeometry.BlankBorderOfFile(file));

            // A stamp not settled yet is read again - and these bytes are no image.
            File.SetLastWriteTimeUtc(file, DateTime.UtcNow);
            Assert.Equal(default, ImageFrameGeometry.BlankBorderOfFile(file));
        }
        finally
        {
            File.Delete(file);
        }
    }
}
