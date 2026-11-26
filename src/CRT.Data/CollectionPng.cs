using System;
using System.Buffers.Binary;
using System.IO;
using System.IO.Compression;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE ONE KIND OF FILE A COLLECTION TAKES: A PLAIN PNG, CHECKED TO ITS LAST BYTE (owner request, 2026-11-25: uploads "only
    // ... from inside the CRT application, so it cannot be taken over and misused for other purpose"). CRT is open source, so
    // nothing it sends proves a file came from it - but a file that is nothing but a picture of a size an oscilloscope's screen
    // has is no use for keeping anything else. So, with no imaging library (the server decodes no picture):
    //
    //   - the PNG signature first, then chunks of only the kinds a picture needs (IHDR, PLTE, IDAT, IEND) and the few small
    //     ones that say how to show it - never a text chunk or one of a kind nobody knows;
    //   - every chunk's checksum right, each within its kind's size;
    //   - the picture's data unpacking to EXACTLY its rows, each with a known filter - nothing more inside the data;
    //   - IEND last, and nothing after it.
    //
    // Interlaced pictures are refused: CRT never writes one, and each pass would need its own sums. CRT checks its file by this
    // very rule before sending it (CollectionImages in CRT), and the server again as it takes it, so the two cannot disagree.
    // ###########################################################################################
    public static class CollectionPng
    {
        // The most a capture may be: the widest oscilloscope screens CRT reads are 1920 x 1080 (measured: CRT's shipped
        // baselines are all 800 x 480, 41 KB at the median and 148 KB at most) - room to spare, and nothing like a photo.
        public const int MaximumBytes = 2 * 1024 * 1024;

        public const int MaximumWidth = 2048;

        public const int MaximumHeight = 1536;

        public const string ContentType = "image/png";

        private static ReadOnlySpan<byte> Signature => [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        // ###########################################################################################
        // What the check found: the picture's size, or why it is no capture (a sentence a person reads).
        // ###########################################################################################
        public sealed record Result(bool IsCapture, int Width, int Height, string? Problem)
        {
            internal static Result No(string problem) => new(false, 0, 0, problem);
        }

        public const string NotAPicture = "This is not a PNG picture of an oscilloscope's screen.";

        public static Result Check(ReadOnlySpan<byte> bytes)
        {
            if (bytes.Length > CollectionPng.MaximumBytes)
                return Result.No($"A capture is at most {CollectionPng.MaximumBytes / (1024 * 1024)} MB.");

            if (bytes.Length < CollectionPng.Signature.Length + 12 || !bytes[..8].SequenceEqual(CollectionPng.Signature))
                return Result.No(CollectionPng.NotAPicture);

            int at = 8;
            int width = 0;
            int height = 0;
            byte bitDepth = 0;
            byte colorType = 0;
            bool seenHeader = false;
            bool seenPalette = false;
            bool seenData = false;
            bool dataEnded = false;
            var data = new MemoryStream();

            while (true)
            {
                if (bytes.Length - at < 12)
                    return Result.No(CollectionPng.NotAPicture);

                uint length = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(at, 4));

                if (length > (uint)(bytes.Length - at - 12))
                    return Result.No(CollectionPng.NotAPicture);

                ReadOnlySpan<byte> typeAndData = bytes.Slice(at + 4, 4 + (int)length);
                string type = System.Text.Encoding.ASCII.GetString(typeAndData[..4]);
                ReadOnlySpan<byte> chunk = typeAndData[4..];
                uint crc = BinaryPrimitives.ReadUInt32BigEndian(bytes.Slice(at + 8 + (int)length, 4));

                if (Crc32(typeAndData) != crc)
                    return Result.No(CollectionPng.NotAPicture);

                at += 12 + (int)length;

                if (!seenHeader && type != "IHDR")
                    return Result.No(CollectionPng.NotAPicture);

                // A data chunk ends the data only once another kind follows; nothing but IEND may come after it.
                if (seenData && type != "IDAT")
                    dataEnded = true;

                switch (type)
                {
                    case "IHDR":
                        if (seenHeader || length != 13)
                            return Result.No(CollectionPng.NotAPicture);

                        width = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(chunk[..4]), int.MaxValue);
                        height = (int)Math.Min(BinaryPrimitives.ReadUInt32BigEndian(chunk.Slice(4, 4)), int.MaxValue);
                        bitDepth = chunk[8];
                        colorType = chunk[9];

                        if (chunk[10] != 0 || chunk[11] != 0)
                            return Result.No(CollectionPng.NotAPicture);

                        if (chunk[12] != 0)
                            return Result.No("An interlaced picture cannot be taken - CRT never writes one.");

                        if (!CollectionPng.IsKnownFormat(colorType, bitDepth))
                            return Result.No(CollectionPng.NotAPicture);

                        if (width < 1 || height < 1)
                            return Result.No(CollectionPng.NotAPicture);

                        if (width > CollectionPng.MaximumWidth || height > CollectionPng.MaximumHeight)
                            return Result.No($"A capture is at most {CollectionPng.MaximumWidth} x {CollectionPng.MaximumHeight} pixels.");

                        seenHeader = true;
                        break;

                    case "PLTE":
                        if (seenPalette || seenData || length == 0 || length % 3 != 0 || length > 256 * 3 || colorType is 0 or 4)
                            return Result.No(CollectionPng.NotAPicture);

                        seenPalette = true;
                        break;

                    case "IDAT":
                        if (dataEnded || (colorType == 3 && !seenPalette))
                            return Result.No(CollectionPng.NotAPicture);

                        seenData = true;
                        data.Write(chunk);
                        break;

                    case "IEND":
                        if (length != 0 || !seenData || at != bytes.Length)
                            return Result.No(CollectionPng.NotAPicture);

                        data.Position = 0;

                        return CollectionPng.RowsAreExact(data, width, height, colorType, bitDepth)
                            ? new Result(true, width, height, null)
                            : Result.No(CollectionPng.NotAPicture);

                    default:
                        if (!CollectionPng.IsSmallAncillary(type, length) || seenData)
                            return Result.No(CollectionPng.NotAPicture);

                        break;
                }
            }
        }

        // The colour types and bit depths the PNG standard allows together.
        private static bool IsKnownFormat(byte colorType, byte bitDepth) => colorType switch
        {
            0 => bitDepth is 1 or 2 or 4 or 8 or 16,
            2 or 4 or 6 => bitDepth is 8 or 16,
            3 => bitDepth is 1 or 2 or 4 or 8,
            _ => false
        };

        // The chunks that only say how to show the picture, each no larger than its kind needs. The colour profile (iCCP) is
        // what a graphics library writes beside a picture in a colour space; it holds a compressed profile of a few KB.
        private static bool IsSmallAncillary(string type, uint length) => type switch
        {
            "tRNS" => length <= 256,
            "gAMA" => length == 4,
            "cHRM" => length == 32,
            "sRGB" => length == 1,
            "pHYs" => length == 9,
            "sBIT" => length <= 4,
            "bKGD" => length <= 6,
            "iCCP" => length <= 16 * 1024,
            _ => false
        };

        private static int ChannelsOf(byte colorType) => colorType switch
        {
            0 or 3 => 1,
            2 => 3,
            4 => 2,
            _ => 4
        };

        // The data unpacked row by row: each row a filter byte (0 to 4) and its pixels, and nothing after the last row.
        private static bool RowsAreExact(Stream compressed, int width, int height, byte colorType, byte bitDepth)
        {
            long bitsPerRow = (long)width * CollectionPng.ChannelsOf(colorType) * bitDepth;
            int bytesPerRow = (int)((bitsPerRow + 7) / 8);
            byte[] row = new byte[1 + bytesPerRow];

            try
            {
                using var unpacked = new ZLibStream(compressed, CompressionMode.Decompress);

                for (int y = 0; y < height; y++)
                {
                    if (!CollectionPng.ReadExactly(unpacked, row) || row[0] > 4)
                        return false;
                }

                // Nothing more inside the data.
                return unpacked.Read(new byte[1], 0, 1) == 0;
            }
            catch (InvalidDataException)
            {
                return false;
            }
        }

        private static bool ReadExactly(Stream stream, byte[] buffer)
        {
            int read = 0;

            while (read < buffer.Length)
            {
                int got = stream.Read(buffer, read, buffer.Length - read);

                if (got == 0)
                    return false;

                read += got;
            }

            return true;
        }

        // ###########################################################################################
        // The PNG chunk checksum - CRC-32 as ISO 3309 defines it, over a chunk's type and data.
        // ###########################################################################################
        private static readonly uint[] CrcTable = CollectionPng.MakeCrcTable();

        private static uint[] MakeCrcTable()
        {
            var table = new uint[256];

            for (uint n = 0; n < 256; n++)
            {
                uint c = n;

                for (int k = 0; k < 8; k++)
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;

                table[n] = c;
            }

            return table;
        }

        internal static uint Crc32(ReadOnlySpan<byte> bytes)
        {
            uint c = 0xFFFFFFFFu;

            foreach (byte b in bytes)
                c = CollectionPng.CrcTable[(c ^ b) & 0xFF] ^ (c >> 8);

            return c ^ 0xFFFFFFFFu;
        }
    }
}
