using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform;
using System;
using System.Runtime.InteropServices;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHAT A COLOR DROPPER READS (owner request, 2026-11-03: "I really miss the 'color picker' option,
    // as it then would be easy to select the same color as something existing"; ColorDropperButton): CRT's
    // window as a picture, 4 bytes a pixel, and the color at a point of it. Points are the window's own
    // units, a pixel is the render scale's - on a 150 % screen 100 is pixel 150.
    //
    // Pure, past Of: tested without a display, which draws nothing a test could read.
    // ###########################################################################################
    internal sealed class ScreenPixels
    {
        private readonly byte[] thisPixels;
        private readonly PixelSize thisSize;
        private readonly bool thisIsBgra;
        private readonly double thisScale;

        public ScreenPixels(byte[] pixels, PixelSize size, bool isBgra, double scale)
        {
            this.thisPixels = pixels;
            this.thisSize = size;
            this.thisIsBgra = isBgra;
            this.thisScale = scale;
        }

        // ###########################################################################################
        // The color at `point` - null off the picture, or where nothing was drawn. Always opaque: it is
        // the color seen, and a window's picture holds it premultiplied should anything be see-through.
        // ###########################################################################################
        public Color? ColorAt(Point point)
        {
            int x = (int)Math.Floor(point.X * this.thisScale);
            int y = (int)Math.Floor(point.Y * this.thisScale);

            if (x < 0 || y < 0 || x >= this.thisSize.Width || y >= this.thisSize.Height)
            {
                return null;
            }

            int at = ((y * this.thisSize.Width) + x) * 4;
            byte alpha = this.thisPixels[at + 3];

            if (alpha == 0)
            {
                return null;
            }

            byte first = ScreenPixels.Unmultiplied(this.thisPixels[at], alpha);
            byte green = ScreenPixels.Unmultiplied(this.thisPixels[at + 1], alpha);
            byte third = ScreenPixels.Unmultiplied(this.thisPixels[at + 2], alpha);

            return this.thisIsBgra ? Color.FromRgb(third, green, first) : Color.FromRgb(first, green, third);
        }

        private static byte Unmultiplied(byte value, byte alpha) =>
            alpha == 255 ? value : (byte)Math.Min(255, value * 255 / alpha);

        // A color as "#RRGGBB".
        public static string HexOf(Color color) => $"#{color.R:X2}{color.G:X2}{color.B:X2}";

        // ###########################################################################################
        // Where the lens beside the pointer goes: below it to the right, on the other side where it would
        // leave the window - and always inside it, as far as it fits.
        // ###########################################################################################
        public static Point LensPlace(Point pointer, Size lens, Size area)
        {
            const double gap = 18.0;

            double x = pointer.X + gap + lens.Width <= area.Width ? pointer.X + gap : pointer.X - gap - lens.Width;
            double y = pointer.Y + gap + lens.Height <= area.Height ? pointer.Y + gap : pointer.Y - gap - lens.Height;

            return new Point(Math.Max(0, x), Math.Max(0, y));
        }

        // A picture's pixels - null for a pixel format not of 4 bytes, or more than one array holds.
        public static ScreenPixels? Of(Bitmap bitmap, double scale)
        {
            if (bitmap.Format is not { } format || (format != PixelFormats.Bgra8888 && format != PixelFormats.Rgba8888))
            {
                return null;
            }

            PixelSize size = bitmap.PixelSize;
            long length = (long)size.Width * 4 * size.Height;

            if (size.Width <= 0 || size.Height <= 0 || length > int.MaxValue)
            {
                return null;
            }

            byte[] pixels = new byte[length];
            GCHandle pin = GCHandle.Alloc(pixels, GCHandleType.Pinned);

            try
            {
                bitmap.CopyPixels(new PixelRect(size), pin.AddrOfPinnedObject(), pixels.Length, size.Width * 4);
            }
            finally
            {
                pin.Free();
            }

            return new ScreenPixels(pixels, size, format == PixelFormats.Bgra8888, scale);
        }
    }
}
