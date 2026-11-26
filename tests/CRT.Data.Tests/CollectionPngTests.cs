using System.Buffers.Binary;
using Handlers.DataHandling;

namespace CRT.Data.Tests
{
    // ###########################################################################################
    // A collection takes a plain PNG and nothing else (CollectionPng) - so the upload is no use for keeping any other file.
    // Every way a file can carry something more than its picture is refused: another kind of chunk, bytes after the end, more
    // data than the rows, a damaged checksum.
    // ###########################################################################################
    public class CollectionPngTests
    {
        [Fact]
        public void A_plain_picture_is_a_capture_and_its_size_is_read()
        {
            CollectionPng.Result result = CollectionPng.Check(TestPng.Rgb(800, 480));

            Assert.True(result.IsCapture, result.Problem);
            Assert.Equal(800, result.Width);
            Assert.Equal(480, result.Height);
        }

        [Theory]
        [InlineData(0, 8)]
        [InlineData(3, 8)]
        [InlineData(6, 8)]
        [InlineData(6, 16)]
        [InlineData(0, 1)]
        public void Every_format_the_standard_allows_is_taken(byte colorType, byte bitDepth)
        {
            int channels = colorType switch { 0 or 3 => 1, 2 => 3, 4 => 2, _ => 4 };
            int width = 16;
            int bytesPerRow = (width * channels * bitDepth + 7) / 8;
            var rows = new byte[(1 + bytesPerRow) * 4];

            var extra = colorType == 3 ? new[] { ("PLTE", new byte[] { 0, 0, 0, 255, 255, 255 }) } : [];

            byte[] png = TestPng.Build(TestPng.Header(width, 4, colorType, bitDepth), TestPng.Data(rows), extra);

            Assert.True(CollectionPng.Check(png).IsCapture);
        }

        // What a graphics library writes beside a picture to say how to show it - small, and taken.
        [Fact]
        public void The_chunks_that_say_how_to_show_it_are_taken()
        {
            byte[] png = TestPng.Build(
                TestPng.Header(8, 8),
                TestPng.Data(TestPng.Rows(8, 8, 3)),
                ("sRGB", [0]),
                ("gAMA", [0, 0, 0xB1, 0x8F]),
                ("pHYs", [0, 0, 0x0E, 0xC3, 0, 0, 0x0E, 0xC3, 1]),
                ("iCCP", new byte[3000]));

            Assert.True(CollectionPng.Check(png).IsCapture);
        }

        // *** A TEXT CHUNK COULD HOLD ANYTHING - and so could a chunk of a kind nobody knows. ***
        [Theory]
        [InlineData("tEXt")]
        [InlineData("zTXt")]
        [InlineData("iTXt")]
        [InlineData("eXIf")]
        [InlineData("abCd")]
        public void A_chunk_that_could_hold_anything_else_is_refused(string type)
        {
            byte[] png = TestPng.Build(TestPng.Header(8, 8), TestPng.Data(TestPng.Rows(8, 8, 3)), (type, new byte[100]));

            CollectionPng.Result result = CollectionPng.Check(png);

            Assert.False(result.IsCapture);
            Assert.Equal(CollectionPng.NotAPicture, result.Problem);
        }

        [Fact]
        public void A_colour_profile_larger_than_any_is_refused() =>
            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(8, 8), TestPng.Data(TestPng.Rows(8, 8, 3)), ("iCCP", new byte[20 * 1024]))).IsCapture);

        [Fact]
        public void Bytes_after_the_end_are_refused()
        {
            byte[] png = [.. TestPng.Rgb(16, 16), .. "anything at all"u8.ToArray()];

            Assert.False(CollectionPng.Check(png).IsCapture);
        }

        // The data unpacked must be exactly the rows - nothing hidden after the last.
        [Fact]
        public void More_data_than_the_rows_is_refused()
        {
            byte[] rows = [.. TestPng.Rows(16, 16, 3), .. new byte[500]];

            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(16, 16), TestPng.Data(rows))).IsCapture);
        }

        [Fact]
        public void Less_data_than_the_rows_is_refused()
        {
            byte[] rows = TestPng.Rows(16, 16, 3)[..^10];

            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(16, 16), TestPng.Data(rows))).IsCapture);
        }

        [Fact]
        public void A_row_with_no_known_filter_is_refused() =>
            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(16, 16), TestPng.Data(TestPng.Rows(16, 16, 3, filter: 9)))).IsCapture);

        [Fact]
        public void A_damaged_checksum_is_refused()
        {
            byte[] png = TestPng.Rgb(16, 16);

            // The last byte of IHDR's checksum.
            png[8 + 8 + 13 + 3] ^= 0xFF;

            Assert.False(CollectionPng.Check(png).IsCapture);
        }

        [Fact]
        public void Data_that_does_not_unpack_is_refused() =>
            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(16, 16), [1, 2, 3, 4, 5, 6, 7, 8])).IsCapture);

        [Fact]
        public void An_interlaced_picture_is_refused_saying_why()
        {
            CollectionPng.Result result = CollectionPng.Check(TestPng.Build(TestPng.Header(16, 16, interlace: 1), TestPng.Data(TestPng.Rows(16, 16, 3))));

            Assert.False(result.IsCapture);
            Assert.Contains("interlaced", result.Problem, StringComparison.Ordinal);
        }

        [Fact]
        public void A_picture_larger_than_a_screen_is_refused_saying_how_large_one_may_be()
        {
            CollectionPng.Result result = CollectionPng.Check(TestPng.Build(TestPng.Header(4000, 3000), TestPng.Data([])));

            Assert.False(result.IsCapture);
            Assert.Contains($"{CollectionPng.MaximumWidth} x {CollectionPng.MaximumHeight}", result.Problem, StringComparison.Ordinal);
        }

        [Fact]
        public void A_file_over_the_size_limit_is_refused_before_it_is_read()
        {
            byte[] large = new byte[CollectionPng.MaximumBytes + 1];
            TestPng.Signature.CopyTo(large, 0);

            CollectionPng.Result result = CollectionPng.Check(large);

            Assert.False(result.IsCapture);
            Assert.Contains("MB", result.Problem, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(new byte[0])]
        [InlineData(new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0, 0x10, 0x4A, 0x46, 0x49, 0x46, 0, 1, 1, 0, 0, 1, 0, 1, 0, 0 })]
        public void Something_that_is_no_png_is_refused(byte[] bytes) =>
            Assert.False(CollectionPng.Check(bytes).IsCapture);

        [Fact]
        public void A_palette_picture_without_its_palette_is_refused()
        {
            var rows = new byte[(1 + 8) * 8];

            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(8, 8, colorType: 3), TestPng.Data(rows))).IsCapture);
        }

        [Fact]
        public void A_format_the_standard_does_not_allow_is_refused() =>
            Assert.False(CollectionPng.Check(TestPng.Build(TestPng.Header(8, 8, colorType: 2, bitDepth: 4), TestPng.Data(new byte[100]))).IsCapture);

        [Fact]
        public void A_chunk_after_the_data_other_than_its_end_is_refused()
        {
            byte[] good = TestPng.Rgb(8, 8);
            byte[] iend = TestPng.Chunk("IEND", []);

            // The picture less its IEND, then a gAMA, then IEND again.
            byte[] png = [.. good[..^iend.Length], .. TestPng.Chunk("gAMA", [0, 0, 0xB1, 0x8F]), .. iend];

            Assert.False(CollectionPng.Check(png).IsCapture);
        }

        [Fact]
        public void A_chunk_claiming_more_bytes_than_the_file_has_is_refused()
        {
            byte[] png = TestPng.Rgb(8, 8);

            BinaryPrimitives.WriteUInt32BigEndian(png.AsSpan(8, 4), 0x7FFFFFFF);

            Assert.False(CollectionPng.Check(png).IsCapture);
        }
    }
}
