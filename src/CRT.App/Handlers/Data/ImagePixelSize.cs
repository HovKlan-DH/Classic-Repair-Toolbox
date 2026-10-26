using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A PICTURE'S OWN PIXEL SIZE, READ FROM ITS HEADER - for every type CRT shows (ImageFileTypes:
    // PNG, JPEG, GIF, BMP, WebP), without decoding it. Highlights are kept in a picture's own pixels,
    // so a preview of a picture made small has to know its real size to place them (code review,
    // 2026-10-26: only PNG and JPEG were read - WorkbookPdfExporter.TryReadImageSize, used here for
    // those two - and a GIF, BMP or WebP had its marks drawn at the decoded size, six times too big).
    // Null when the header cannot be read: a preview then draws no marks rather than wrong ones.
    // ###########################################################################################
    internal static class ImagePixelSize
    {
        public static (int Width, int Height)? TryRead(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                return null;
            }

            if (WorkbookPdfExporter.TryReadImageSize(path) is { } pngOrJpeg)
            {
                return pngOrJpeg;
            }

            try
            {
                byte[] header = new byte[32];
                int read;

                using (FileStream stream = File.OpenRead(path))
                {
                    read = stream.ReadAtLeast(header, header.Length, throwOnEndOfStream: false);
                }

                return ImagePixelSize.FromHeader(header.AsSpan(0, read));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }

        // GIF, BMP and WebP from the file's first 32 bytes.
        internal static (int Width, int Height)? FromHeader(ReadOnlySpan<byte> header)
        {
            (int Width, int Height)? size =
                ImagePixelSize.Gif(header) ?? ImagePixelSize.Bmp(header) ?? ImagePixelSize.WebP(header);

            return size is { Width: > 0, Height: > 0 } ? size : null;
        }

        // "GIF87a"/"GIF89a", then the logical screen's width and height, little-endian 16 bits each.
        private static (int Width, int Height)? Gif(ReadOnlySpan<byte> header)
        {
            if (header.Length < 10 || header[0] != 'G' || header[1] != 'I' || header[2] != 'F' || header[3] != '8' ||
                (header[4] != '7' && header[4] != '9') || header[5] != 'a')
            {
                return null;
            }

            return (header[6] | (header[7] << 8), header[8] | (header[9] << 8));
        }

        // "BM", then the DIB header: the old 12-byte core header holds 16-bit sizes, every later one
        // 32-bit; a negative height is a picture stored top-down.
        private static (int Width, int Height)? Bmp(ReadOnlySpan<byte> header)
        {
            if (header.Length < 26 || header[0] != 'B' || header[1] != 'M')
            {
                return null;
            }

            int dibSize = BitConverter.ToInt32(header.Slice(14, 4));

            if (dibSize == 12)
            {
                return (BitConverter.ToUInt16(header.Slice(18, 2)), BitConverter.ToUInt16(header.Slice(20, 2)));
            }

            int width = BitConverter.ToInt32(header.Slice(18, 4));
            int height = BitConverter.ToInt32(header.Slice(22, 4));

            return height == int.MinValue ? null : (width, Math.Abs(height));
        }

        // "RIFF" .... "WEBP", then the first chunk: lossy "VP8 " (14-bit sizes after the key frame's
        // start code), lossless "VP8L" (14-bit sizes, less one, packed after the 0x2F signature), or
        // extended "VP8X" (the canvas's 24-bit sizes, less one).
        private static (int Width, int Height)? WebP(ReadOnlySpan<byte> header)
        {
            if (header.Length < 30 ||
                header[0] != 'R' || header[1] != 'I' || header[2] != 'F' || header[3] != 'F' ||
                header[8] != 'W' || header[9] != 'E' || header[10] != 'B' || header[11] != 'P' ||
                header[12] != 'V' || header[13] != 'P' || header[14] != '8')
            {
                return null;
            }

            switch ((char)header[15])
            {
                case ' ':
                    if (header[23] != 0x9D || header[24] != 0x01 || header[25] != 0x2A)
                        return null;

                    return (BitConverter.ToUInt16(header.Slice(26, 2)) & 0x3FFF, BitConverter.ToUInt16(header.Slice(28, 2)) & 0x3FFF);

                case 'L':
                    if (header[20] != 0x2F)
                        return null;

                    int bits = header[21] | (header[22] << 8) | (header[23] << 16) | (header[24] << 24);
                    return ((bits & 0x3FFF) + 1, ((bits >> 14) & 0x3FFF) + 1);

                case 'X':
                    return (1 + (header[24] | (header[25] << 8) | (header[26] << 16)),
                            1 + (header[27] | (header[28] << 8) | (header[29] << 16)));

                default:
                    return null;
            }
        }
    }
}
