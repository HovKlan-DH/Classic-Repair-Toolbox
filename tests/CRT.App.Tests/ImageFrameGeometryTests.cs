using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ImageFrameGeometry - room round the "Schematics" tab's image so its badges cover none of it (owner request, 2026-11-13:
// "Ideally the cropped one can be displayed nicely and with some padding to its borders ... and then ideally the
// "original" one is displayed as-is today"). The owner's two images of the Open128-2 board are the cases: "PCB Top", with
// room of its own all round, and "PCB Top CROPPED", cut to the board.
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

    // "PCB Top" as measured (5080 x 3604, empty 336 above, 211 below, 89 left, 73 right) in a panel of 1400 x 850: its own
    // room already gives more than is wanted on every side, so nothing is added - it shows as it always did.
    // At 990 x 750 too, where its 73 empty columns at the right draw as 14 - less than the 16 asked, but room of its own:
    // topped up by two it drew 0.2% smaller than it always did (the render of the real image, 2026-11-13).
    [Theory]
    [InlineData(1400, 850)]
    [InlineData(990, 750)]
    public void The_owners_padded_image_gets_no_room_added(double width, double height)
    {
        Thickness frame = ImageFrameGeometry.Frame(new Size(width, height), new PixelSize(5080, 3604), new Thickness(89, 336, 73, 211), ImageFrameGeometry.Wanted);

        Assert.Equal(default, frame);
    }

    // At the sides any room of its own is the image's choice, however little; above, the badges need their room whatever.
    [Fact]
    public void A_side_with_room_of_its_own_gets_none_added_but_the_top_is_topped_up()
    {
        Thickness frame = ImageFrameGeometry.Frame(new Size(1000, 800), new PixelSize(1000, 700), new Thickness(3, 3, 3, 3), ImageFrameGeometry.Wanted);

        Assert.Equal(0, frame.Left);
        Assert.Equal(0, frame.Right);
        Assert.Equal(0, frame.Bottom);
        Assert.True(frame.Top > 40, $"{frame.Top} above");
    }

    // "PCB Top CROPPED" as measured (4903 x 3058, empty 0 above, 2 below, 1 left, 0 right): the room wanted all round, less
    // the pixel or two it has, so the board stands under the badges' 44 with 16 at the other sides.
    [Fact]
    public void The_owners_cropped_image_gets_the_room_all_round()
    {
        Thickness frame = ImageFrameGeometry.Frame(new Size(1400, 850), new PixelSize(4903, 3058), new Thickness(1, 0, 0, 2), ImageFrameGeometry.Wanted);

        Assert.Equal(ImageFrameGeometry.TopRoom, frame.Top, precision: 3);
        Assert.Equal(ImageFrameGeometry.SideRoom, frame.Right, precision: 3);
        Assert.InRange(frame.Left, ImageFrameGeometry.SideRoom - 1, ImageFrameGeometry.SideRoom);
        Assert.InRange(frame.Bottom, ImageFrameGeometry.SideRoom - 1, ImageFrameGeometry.SideRoom);

        // And the image fitted inside it: wholly in the panel, its top below the badges.
        Rect content = RectGeometry.GetImageContentRect(new Size(1400, 850), new PixelSize(4903, 3058), frame);
        Assert.Equal(frame.Top, content.Top, precision: 3);
        Assert.Equal(frame.Left, content.Left, precision: 3);
        Assert.True(content.Bottom <= 850 - frame.Bottom + 0.001);
        Assert.True(content.Right <= 1400 - frame.Right + 0.001);
    }

    // Room of its own short of what is wanted is topped up, no more - at the scale the image is drawn at once the room is in.
    [Fact]
    public void Room_of_its_own_short_of_what_is_wanted_is_topped_up()
    {
        // 1000 x 500 in 1000 x 600: drawn about 0.97 - 20 empty rows above give ~19, so ~25 is added.
        Thickness frame = ImageFrameGeometry.Frame(new Size(1000, 600), new PixelSize(1000, 500), new Thickness(0, 20, 0, 0), ImageFrameGeometry.Wanted);
        Rect content = RectGeometry.GetImageContentRect(new Size(1000, 600), new PixelSize(1000, 500), frame);
        double scale = content.Width / 1000;

        Assert.Equal(ImageFrameGeometry.TopRoom, frame.Top + (scale * 20), precision: 2);
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
