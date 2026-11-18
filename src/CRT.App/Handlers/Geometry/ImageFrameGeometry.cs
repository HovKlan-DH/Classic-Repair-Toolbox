using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // ROOM ROUND THE "Schematics" TAB'S IMAGE, SO ITS BADGES COVER NONE OF IT (owner request, 2026-11-13: "The goal should
    // be, that the schematic image is initially (when shown at page first) fully visible and not so close to the top that its
    // info badges blocks something ... Ideally the cropped one can be displayed nicely and with some padding to its borders
    // ... and then ideally the "original" one is displayed as-is today"). The image's name and region stand over its top
    // left corner; an image cropped to the board had them over the board, while the owner's own images carry room of their
    // own for them.
    //
    //   - THE ROOM WANTED: TopRoom above the image - under BOTH rows of badges, the name and region's and the one a component
    //     or pin under the pointer adds, with room to spare (owner request, 2026-11-17: "make sure all schematic images, when
    //     displayed initially (not when zoomed), are displayed below both rows of badges + with a little margin, so it does
    //     not look cramped") - SideRoom on the other sides; in the first view, as the image is fitted, and so at the furthest
    //     it zooms out; zoomed in, the image moves over the whole panel as before.
    //   - THE IMAGE'S CONTENT IS WHAT IS FITTED, ITS OWN EMPTY BORDER LEFT OUT (BlankBorder): its rows and columns of one
    //     color from edge to edge - transparent or plain - counted from each side. A row across the board always holds
    //     something of it - its outline, a trace, a part - so only room round the board is counted, never the board. So
    //     WHETHER AN IMAGE IS CROPPED TO ITS BOARD OR PADDED, ITS BOARD STANDS IN THE SAME PLACE (owner request, 2026-11-17:
    //     "no matter if I have inserted empty spacing or not, then the actual PCB should be displayed same place" - an image's
    //     own room was kept as the owner's choice from 2026-11-13 until then, and the owner's padded "PCB Top" showed its
    //     board lower and smaller than "PCB Top CROPPED"): its empty border lies in the room round it, or past the panel's
    //     edge, where there is nothing to see (Frame).
    //
    // PURE: pixels and sizes in, room out - but for BlankBorderOfFile, which reads an image file for the pixels.
    // ###########################################################################################
    internal static class ImageFrameGeometry
    {
        // Above: the second row of badges stands 44 down (TabSchematics.axaml), 26 high, and 10 more. 44 - under the first row
        // alone - until 2026-11-17, so a component pointed at put its badges over the board's top edge.
        public const double TopRoom = 80;
        public const double SideRoom = 16;

        // How far two pixels may differ per channel and still be the one color - a picture saved as JPEG is never exact.
        public const int Tolerance = 16;

        public static Thickness Wanted => new(SideRoom, TopRoom, SideRoom, SideRoom);

        // ###########################################################################################
        // The image's own empty border, in its pixels: from each side, how many whole rows (top, bottom) or columns (left,
        // right) are of the color of that side's first pixel. `rowAt(y)` and `columnAt(x)` give a row's or a column's
        // pixels, 0xAARRGGBB as Bgra8888 reads in little-endian (Rgba8888 too: only the alpha's place matters). Two
        // pixels with no alpha are the same however their colors read. At most a third of the image from each side - an
        // image almost all one color is no border with a picture in it.
        // ###########################################################################################
        public static Thickness BlankBorder(int width, int height, Func<int, uint[]> rowAt, Func<int, uint[]> columnAt)
        {
            ArgumentNullException.ThrowIfNull(rowAt);
            ArgumentNullException.ThrowIfNull(columnAt);

            if (width <= 0 || height <= 0)
            {
                return default;
            }

            int Count(int lines, Func<int, uint[]> lineAt, bool fromEnd)
            {
                int most = lines / 3;
                uint[] first = lineAt(fromEnd ? lines - 1 : 0);

                if (first.Length == 0)
                {
                    return 0;
                }

                uint reference = first[0];
                int count = 0;

                while (count < most && ImageFrameGeometry.AllLike(lineAt(fromEnd ? lines - 1 - count : count), reference))
                {
                    count++;
                }

                return count;
            }

            return new Thickness(
                Count(width, columnAt, fromEnd: false),
                Count(height, rowAt, fromEnd: false),
                Count(width, columnAt, fromEnd: true),
                Count(height, rowAt, fromEnd: true));
        }

        // ###########################################################################################
        // The same, the pixels read a BAND of `Band` rows or columns at a time - `readRect` gives a rectangle's pixels, row
        // by row - so a large image is read in a few dozen reads, not one per row and column: thousands of them, up to a
        // third of a large scan's width and height each from both sides, before the image was shown (code review,
        // 2026-11-13).
        // ###########################################################################################
        public const int Band = 64;

        public static Thickness BlankBorder(int width, int height, Func<PixelRect, uint[]> readRect)
        {
            ArgumentNullException.ThrowIfNull(readRect);

            Func<int, uint[]> Lines(bool rows)
            {
                int lines = rows ? height : width;
                uint[] band = [];
                int bandStart = -1;
                int bandCount = 0;

                return index =>
                {
                    if (index < bandStart || index >= bandStart + bandCount)
                    {
                        bandStart = index / Band * Band;
                        bandCount = Math.Min(Band, lines - bandStart);
                        band = readRect(rows ? new PixelRect(0, bandStart, width, bandCount) : new PixelRect(bandStart, 0, bandCount, height));
                    }

                    int offset = index - bandStart;
                    var line = new uint[rows ? width : height];

                    if (rows)
                    {
                        Array.Copy(band, offset * width, line, 0, width);
                    }
                    else
                    {
                        for (int y = 0; y < height; y++)
                        {
                            line[y] = band[(y * bandCount) + offset];
                        }
                    }

                    return line;
                };
            }

            return ImageFrameGeometry.BlankBorder(width, height, Lines(rows: true), Lines(rows: false));
        }

        // ###########################################################################################
        // The same of an image FILE, decoded by SkiaSharp - the decoder Avalonia uses: for a bitmap whose pixels Avalonia will
        // not copy (a grayscale one, Gray8; TabSchematics.BlankBorderOf). A grayscale image is read as its gray bytes, each
        // as an opaque pixel of that gray; any other as 32 bits a pixel. None for a file SkiaSharp cannot read.
        //
        // KEPT per file, by where it lies, its length and write time, once that stamp has settled (StampSettling's rule): it
        // is read again at every switch to the image, and decoding a 4300 x 3000 sheet a second time - and a 32-bit copy of
        // it, 52 MB, until grayscale was read as it is - at each was the cost of a frame (code review, 2026-11-13). In
        // CRT.Data's StampedLruCache, the least used file going first - with a settle rule and a cache of its own here, the
        // rule was written twice and all of them went at once (code review, 2026-11-14).
        // ###########################################################################################
        public static Thickness BlankBorderOfFile(string file)
        {
            var info = new System.IO.FileInfo(file);
            string? stamp = info.Exists && Handlers.DataHandling.StampSettling.IsSettled(info.LastWriteTimeUtc)
                ? $"{info.Length}|{info.LastWriteTimeUtc.Ticks}"
                : null;

            if (stamp is not null && ImageFrameGeometry.BordersOfFiles.TryGet(info.FullName, stamp, out Thickness known))
            {
                return known;
            }

            Thickness border = ImageFrameGeometry.ReadBlankBorderOfFile(file);

            if (stamp is not null)
            {
                ImageFrameGeometry.BordersOfFiles.Store(info.FullName, stamp, border);
            }

            return border;
        }

        // A file's border is a few numbers, so many are kept: every schematic image of the boards looked at in a run.
        private static readonly Handlers.DataHandling.StampedLruCache<Thickness> BordersOfFiles = new(256, StringComparer.Ordinal);

        private static Thickness ReadBlankBorderOfFile(string file)
        {
            using SkiaSharp.SKBitmap? decoded = SkiaSharp.SKBitmap.Decode(file);

            return decoded is null ? default : ImageFrameGeometry.BlankBorderOf(decoded);
        }

        // ###########################################################################################
        // The same of an image ENCODED in `encoded` - a picture's thumbnail standing in for it, whose pixels Avalonia will not
        // copy when it is grayscale (TabSchematics.BlankBorderOf; code review, 2026-11-18: measured as no border, the stand-in
        // stood elsewhere than the picture once it came - the jump the stand-in is there to prevent). None for bytes SkiaSharp
        // cannot read.
        // ###########################################################################################
        public static Thickness BlankBorderOfEncoded(System.IO.Stream encoded)
        {
            ArgumentNullException.ThrowIfNull(encoded);

            using SkiaSharp.SKBitmap? decoded = SkiaSharp.SKBitmap.Decode(encoded);

            return decoded is null ? default : ImageFrameGeometry.BlankBorderOf(decoded);
        }

        private static Thickness BlankBorderOf(SkiaSharp.SKBitmap decoded)
        {
            if (decoded.ColorType == SkiaSharp.SKColorType.Gray8)
            {
                int rowBytes = decoded.RowBytes;

                return ImageFrameGeometry.BlankBorder(decoded.Width, decoded.Height, rect =>
                {
                    ReadOnlySpan<byte> all = decoded.GetPixelSpan();
                    var read = new uint[rect.Width * rect.Height];

                    for (int y = 0; y < rect.Height; y++)
                    {
                        ReadOnlySpan<byte> row = all.Slice(((rect.Y + y) * rowBytes) + rect.X, rect.Width);

                        for (int x = 0; x < rect.Width; x++)
                        {
                            uint gray = row[x];
                            read[(y * rect.Width) + x] = 0xFF000000u | (gray << 16) | (gray << 8) | gray;
                        }
                    }

                    return read;
                });
            }

            using SkiaSharp.SKBitmap? pixels = decoded.Copy(SkiaSharp.SKColorType.Bgra8888);

            if (pixels is null)
            {
                return default;
            }

            int stride = pixels.RowBytes / 4;

            // Each band read from the decoded pixels where they lie - no second copy of the whole image.
            return ImageFrameGeometry.BlankBorder(pixels.Width, pixels.Height, rect =>
            {
                ReadOnlySpan<uint> all = System.Runtime.InteropServices.MemoryMarshal.Cast<byte, uint>(pixels.GetPixelSpan());
                var read = new uint[rect.Width * rect.Height];

                for (int y = 0; y < rect.Height; y++)
                {
                    all.Slice(((rect.Y + y) * stride) + rect.X, rect.Width).CopyTo(read.AsSpan(y * rect.Width, rect.Width));
                }

                return read;
            });
        }

        private static bool AllLike(uint[] line, uint reference)
        {
            foreach (uint pixel in line)
            {
                if (!ImageFrameGeometry.Alike(pixel, reference))
                {
                    return false;
                }
            }

            return true;
        }

        internal static bool Alike(uint a, uint b)
        {
            if ((a >> 24) == 0 && (b >> 24) == 0)
            {
                return true;
            }

            for (int shift = 0; shift < 32; shift += 8)
            {
                if (Math.Abs((int)((a >> shift) & 0xFF) - (int)((b >> shift) & 0xFF)) > Tolerance)
                {
                    return false;
                }
            }

            return true;
        }

        // ###########################################################################################
        // The room round an image of `image` pixels in `box`, with `blank` of its own (BlankBorder), so its CONTENT - the image
        // less that border - is fitted in `box` less `wanted`, at its top left: `wanted` at each side less the image's own
        // border at the scale it is drawn at. Less than nothing where its border is wider than the room wanted - the image
        // then reaches past the room, or past the panel's edge, with nothing of its content there. RectGeometry's fit of the
        // whole image inside `box` less this room is then exactly that: a side's own border and the room taken from it add up
        // to `wanted` on that side. No image, no room; a box the room would leave nothing of, no room either; an image that
        // is all border is fitted whole.
        // ###########################################################################################
        public static Thickness Frame(Size box, PixelSize image, Thickness blank, Thickness wanted)
        {
            if (box.Width <= 0 || box.Height <= 0 || image.Width <= 0 || image.Height <= 0)
            {
                return default;
            }

            double areaWidth = box.Width - wanted.Left - wanted.Right;
            double areaHeight = box.Height - wanted.Top - wanted.Bottom;

            if (areaWidth <= 0 || areaHeight <= 0)
            {
                return default;
            }

            double contentWidth = image.Width - blank.Left - blank.Right;
            double contentHeight = image.Height - blank.Top - blank.Bottom;

            if (contentWidth <= 0 || contentHeight <= 0)
            {
                blank = default;
                contentWidth = image.Width;
                contentHeight = image.Height;
            }

            double scale = Math.Min(areaWidth / contentWidth, areaHeight / contentHeight);

            return new Thickness(
                wanted.Left - (scale * blank.Left),
                wanted.Top - (scale * blank.Top),
                wanted.Right - (scale * blank.Right),
                wanted.Bottom - (scale * blank.Bottom));
        }

        // ###########################################################################################
        // Where the image's CONTENT stands, drawn in `drawn`: the image of `image` pixels less its own empty border, at the
        // scale it is drawn at. Whether the first view can be panned asks THIS, never the whole image (code review,
        // 2026-11-18): a padded image reaches past the panel's edge at the first view (Frame), and judged by its whole size it
        // was pannable there, its border dragged into view and its board away from where a cropped image's stands. An
        // image that is all border is its whole self, as Frame fits it.
        // ###########################################################################################
        public static Rect ContentWithin(Rect drawn, PixelSize image, Thickness blank)
        {
            if (image.Width <= 0 || image.Height <= 0 ||
                image.Width - blank.Left - blank.Right <= 0 || image.Height - blank.Top - blank.Bottom <= 0)
            {
                return drawn;
            }

            double scaleX = drawn.Width / image.Width;
            double scaleY = drawn.Height / image.Height;

            return new Rect(
                drawn.X + (blank.Left * scaleX),
                drawn.Y + (blank.Top * scaleY),
                drawn.Width - ((blank.Left + blank.Right) * scaleX),
                drawn.Height - ((blank.Top + blank.Bottom) * scaleY));
        }
    }
}
