using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // PNG files made by hand for the collection tests (CollectionPng) - a good one, and each kind of wrong one, chunk by chunk
    // with right checksums unless a test breaks one on purpose. Here in CRT.App.Tests/Shared, as both App test projects use it
    // (CLAUDE.md), and compiled into CRT.Data.Tests and CRT.Server.Tests too (their csproj), so CRT's store, the rule and the
    // server's flows are tested on the very same files.
    // ###########################################################################################
    public static class TestPng
    {
        public static readonly byte[] Signature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        // A plain RGB picture, 8 bits a channel, every row filter 0 - its pixels made from `seed`, so two seeds are two pictures.
        public static byte[] Rgb(int width, int height, int seed = 1) =>
            TestPng.Build(TestPng.Header(width, height, colorType: 2, bitDepth: 8), TestPng.Data(TestPng.Rows(width, height, 3, seed)));

        public static byte[] Header(int width, int height, byte colorType = 2, byte bitDepth = 8, byte interlace = 0)
        {
            var header = new byte[13];

            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(0, 4), (uint)width);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4, 4), (uint)height);
            header[8] = bitDepth;
            header[9] = colorType;
            header[12] = interlace;

            return header;
        }

        // Rows of `bytesPerPixel` bytes a pixel, each led by its filter byte.
        public static byte[] Rows(int width, int height, int bytesPerPixel, int seed = 1, byte filter = 0)
        {
            int rowBytes = 1 + width * bytesPerPixel;
            var rows = new byte[rowBytes * height];

            for (int y = 0; y < height; y++)
            {
                rows[y * rowBytes] = filter;

                for (int x = 1; x < rowBytes; x++)
                    rows[y * rowBytes + x] = (byte)(x * 7 + y * 13 + seed * 31);
            }

            return rows;
        }

        public static byte[] Data(byte[] rows)
        {
            using var packed = new MemoryStream();

            using (var zlib = new ZLibStream(packed, CompressionLevel.Optimal, leaveOpen: true))
                zlib.Write(rows);

            return packed.ToArray();
        }

        // Signature, IHDR, the chunks given (between IHDR and IDAT), IDAT, IEND.
        public static byte[] Build(byte[] header, byte[] data, params (string Type, byte[] Data)[] extra)
        {
            using var png = new MemoryStream();

            png.Write(TestPng.Signature);
            png.Write(TestPng.Chunk("IHDR", header));

            foreach ((string type, byte[] bytes) in extra)
                png.Write(TestPng.Chunk(type, bytes));

            png.Write(TestPng.Chunk("IDAT", data));
            png.Write(TestPng.Chunk("IEND", []));

            return png.ToArray();
        }

        public static byte[] Chunk(string type, byte[] data)
        {
            var chunk = new byte[12 + data.Length];

            BinaryPrimitives.WriteUInt32BigEndian(chunk.AsSpan(0, 4), (uint)data.Length);
            Encoding.ASCII.GetBytes(type, chunk.AsSpan(4, 4));
            data.CopyTo(chunk, 8);
            BinaryPrimitives.WriteUInt32BigEndian(chunk.AsSpan(8 + data.Length, 4), TestPng.Crc(chunk.AsSpan(4, 4 + data.Length)));

            return chunk;
        }

        private static uint Crc(ReadOnlySpan<byte> bytes)
        {
            uint c = 0xFFFFFFFFu;

            foreach (byte b in bytes)
            {
                c ^= b;

                for (int k = 0; k < 8; k++)
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            }

            return c ^ 0xFFFFFFFFu;
        }
    }
}
