using System;
using System.Runtime.InteropServices;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // A picture turned upright by its EXIF orientation (JpegExifOrientation reads it; code review,
    // 2026-10-16): its size once turned, and its pixels turned - 4 bytes a pixel, row after row,
    // as FeedbackComposer copies a bitmap's out. The eight values are EXIF's own:
    //
    //   1 as stored               2 mirrored left to right   3 turned half way round
    //   4 mirrored top to bottom  5 mirrored along the diagonal (top-left to bottom-right)
    //   6 turned a quarter clockwise   7 mirrored along the other diagonal
    //   8 turned a quarter anticlockwise
    //
    // 5 to 8 swap the sides: a 4032 x 3024 photo with orientation 6 is 3024 wide and 4032 tall.
    // Any other value is taken as 1.
    // ###########################################################################################
    internal static class ExifOrientationGeometry
    {
        public static bool SwapsSides(int orientation) => orientation is >= 5 and <= 8;

        public static (int Width, int Height) OrientedSize(int width, int height, int orientation) =>
            ExifOrientationGeometry.SwapsSides(orientation) ? (height, width) : (width, height);

        // ###########################################################################################
        // The pixels of a `width` x `height` picture, turned upright: a new array, with the size it
        // now has. Orientation 1 (or an unknown value) gives back a copy as it was.
        // ###########################################################################################
        public static (byte[] Pixels, int Width, int Height) Apply(byte[] pixels, int width, int height, int orientation)
        {
            ArgumentNullException.ThrowIfNull(pixels);
            ArgumentOutOfRangeException.ThrowIfNegative(width);
            ArgumentOutOfRangeException.ThrowIfNegative(height);

            if (pixels.Length != (long)width * height * 4)
                throw new ArgumentException($"{pixels.Length} bytes are not {width} x {height} pixels of 4 bytes.", nameof(pixels));

            (int newWidth, int newHeight) = ExifOrientationGeometry.OrientedSize(width, height, orientation);
            byte[] turned = new byte[pixels.Length];

            // A pixel at a time, whatever order its 4 bytes are in.
            ReadOnlySpan<uint> from = MemoryMarshal.Cast<byte, uint>(pixels);
            Span<uint> to = MemoryMarshal.Cast<byte, uint>(turned.AsSpan());

            for (int y = 0; y < height; y++)
            {
                for (int x = 0; x < width; x++)
                {
                    (int toX, int toY) = orientation switch
                    {
                        2 => (width - 1 - x, y),
                        3 => (width - 1 - x, height - 1 - y),
                        4 => (x, height - 1 - y),
                        5 => (y, x),
                        6 => (height - 1 - y, x),
                        7 => (height - 1 - y, width - 1 - x),
                        8 => (y, width - 1 - x),
                        _ => (x, y),
                    };

                    to[toY * newWidth + toX] = from[y * width + x];
                }
            }

            return (turned, newWidth, newHeight);
        }
    }
}
