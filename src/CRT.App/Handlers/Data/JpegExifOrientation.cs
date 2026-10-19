using System;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** WHICH WAY UP A JPEG IS MEANT TO BE SHOWN (code review, 2026-10-16). *** A phone stores a
    // photo as its sensor saw it and says in the EXIF "Orientation" tag how to turn it upright -
    // 6 for a portrait photo, stored on its side. A mail reader in a browser turns it; Avalonia's
    // decoder does not. So the Feedback tab's field showed such a photo sideways, at a size it would
    // not have in the mail (FeedbackComposer reads this and turns it).
    //
    // Read from the header only: the APP1 "Exif" segment before the image data, its TIFF header's
    // byte order, and tag 0x0112 in the first IFD. 1 (as stored) for anything else - not a JPEG, no
    // EXIF, a value outside 1..8, or a header cut short - so a malformed file is shown as stored,
    // never refused.
    // ###########################################################################################
    internal static class JpegExifOrientation
    {
        public const int AsStored = 1;

        private const int OrientationTag = 0x0112;
        private const int TypeShort = 3;

        public static int Read(string path)
        {
            try
            {
                using FileStream stream = File.OpenRead(path);
                return JpegExifOrientation.Read(stream);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return JpegExifOrientation.AsStored;
            }
        }

        public static int Read(Stream stream)
        {
            ArgumentNullException.ThrowIfNull(stream);

            if (stream.ReadByte() != 0xFF || stream.ReadByte() != 0xD8)
                return JpegExifOrientation.AsStored;

            while (true)
            {
                int prefix = stream.ReadByte();

                if (prefix < 0)
                    return JpegExifOrientation.AsStored;

                if (prefix != 0xFF)
                    continue;

                int marker = stream.ReadByte();

                // A run of 0xFF bytes is legal padding before a marker.
                while (marker == 0xFF)
                    marker = stream.ReadByte();

                if (marker < 0)
                    return JpegExifOrientation.AsStored;

                // Standalone markers carry no length.
                if (marker == 0x01 || (marker >= 0xD0 && marker <= 0xD8))
                    continue;

                // The image data, or its end: the EXIF always comes before.
                if (marker is 0xDA or 0xD9)
                    return JpegExifOrientation.AsStored;

                int high = stream.ReadByte();
                int low = stream.ReadByte();

                if (high < 0 || low < 0)
                    return JpegExifOrientation.AsStored;

                int length = (high << 8) | low;

                if (length < 2)
                    return JpegExifOrientation.AsStored;

                byte[] payload = new byte[length - 2];

                if (stream.ReadAtLeast(payload, payload.Length, throwOnEndOfStream: false) < payload.Length)
                    return JpegExifOrientation.AsStored;

                if (marker == 0xE1 && JpegExifOrientation.TryReadFromExif(payload, out int orientation))
                    return orientation;
            }
        }

        // An APP1 payload: "Exif\0\0", then a TIFF header and its first IFD.
        private static bool TryReadFromExif(byte[] payload, out int orientation)
        {
            orientation = JpegExifOrientation.AsStored;
            ReadOnlySpan<byte> exif = "Exif\0\0"u8;

            if (payload.Length < exif.Length + 8 || !payload.AsSpan(0, exif.Length).SequenceEqual(exif))
                return false;

            ReadOnlySpan<byte> tiff = payload.AsSpan(exif.Length);
            bool bigEndian;

            if (tiff[0] == 'M' && tiff[1] == 'M')
                bigEndian = true;
            else if (tiff[0] == 'I' && tiff[1] == 'I')
                bigEndian = false;
            else
                return false;

            if (JpegExifOrientation.Read16(tiff, 2, bigEndian) != 42)
                return false;

            long ifd = JpegExifOrientation.Read32(tiff, 4, bigEndian);

            if (ifd < 8 || ifd + 2 > tiff.Length)
                return false;

            int entries = JpegExifOrientation.Read16(tiff, (int)ifd, bigEndian);

            for (int i = 0; i < entries; i++)
            {
                long entry = ifd + 2 + 12L * i;

                if (entry + 12 > tiff.Length)
                    return false;

                if (JpegExifOrientation.Read16(tiff, (int)entry, bigEndian) != JpegExifOrientation.OrientationTag)
                    continue;

                if (JpegExifOrientation.Read16(tiff, (int)entry + 2, bigEndian) != JpegExifOrientation.TypeShort)
                    return false;

                int value = JpegExifOrientation.Read16(tiff, (int)entry + 8, bigEndian);

                if (value is < 1 or > 8)
                    return false;

                orientation = value;
                return true;
            }

            return false;
        }

        private static int Read16(ReadOnlySpan<byte> bytes, int at, bool bigEndian) =>
            bigEndian ? (bytes[at] << 8) | bytes[at + 1] : bytes[at] | (bytes[at + 1] << 8);

        private static long Read32(ReadOnlySpan<byte> bytes, int at, bool bigEndian) =>
            bigEndian
                ? ((long)bytes[at] << 24) | ((long)bytes[at + 1] << 16) | ((long)bytes[at + 2] << 8) | bytes[at + 3]
                : bytes[at] | ((long)bytes[at + 1] << 8) | ((long)bytes[at + 2] << 16) | ((long)bytes[at + 3] << 24);
    }
}
