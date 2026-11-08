using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// JpegExifOrientation - which way up a phone photo is meant to be shown (code review, 2026-10-16:
// a portrait photo was shown on its side in the Feedback tab, at a size it would not have in the
// mail). Read from the header only; anything it cannot read is "as stored", never an error.
// ###########################################################################################
public sealed class JpegExifOrientationTests
{
    private static int Read(byte[] bytes) => JpegExifOrientation.Read(new MemoryStream(bytes));

    [Theory]
    [InlineData(1, false)] [InlineData(2, false)] [InlineData(3, false)] [InlineData(4, false)]
    [InlineData(5, false)] [InlineData(6, false)] [InlineData(7, false)] [InlineData(8, false)]
    [InlineData(1, true)] [InlineData(2, true)] [InlineData(3, true)] [InlineData(4, true)]
    [InlineData(5, true)] [InlineData(6, true)] [InlineData(7, true)] [InlineData(8, true)]
    public void Each_orientation_is_read_in_either_byte_order(int orientation, bool bigEndian)
    {
        Assert.Equal(orientation, Read(TestJpegs.Jpeg(640, 480, orientation, bigEndian)));
    }

    [Fact]
    public void A_JPEG_without_EXIF_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read(TestJpegs.Jpeg(640, 480, orientation: 0, bigEndian: false)));
    }

    [Fact]
    public void A_PNG_or_nothing_at_all_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read([0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]));
        Assert.Equal(JpegExifOrientation.AsStored, Read([]));
    }

    [Theory]
    [InlineData(9)]
    [InlineData(255)]
    public void An_orientation_EXIF_does_not_have_is_as_stored(int orientation)
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read(TestJpegs.Jpeg(640, 480, orientation, bigEndian: false)));
    }

    // The tag must be a SHORT; a LONG in its place is a file nobody wrote right, shown as stored.
    [Fact]
    public void An_orientation_that_is_not_a_SHORT_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read(TestJpegs.Jpeg(640, 480, 6, bigEndian: false, orientationType: 4)));
    }

    // Cut short anywhere, the header is read as far as it goes and never throws.
    [Fact]
    public void A_header_cut_short_anywhere_is_read_without_an_error()
    {
        byte[] whole = TestJpegs.Jpeg(640, 480, 6, bigEndian: true);

        for (int length = 0; length < whole.Length; length++)
        {
            int read = Read(whole[..length]);
            Assert.True(read is JpegExifOrientation.AsStored or 6, $"{length}: {read}");
        }
    }

    // The image size the composer reads comes from the same header, past the EXIF segment.
    [Fact]
    public void The_size_reader_finds_the_frame_header_behind_the_EXIF()
    {
        using var workspace = new TempWorkspace();
        string path = Path.Combine(workspace.Root, "photo.jpg");
        File.WriteAllBytes(path, TestJpegs.Jpeg(4032, 3024, 6, bigEndian: false));

        Assert.Equal((4032, 3024), WorkbookPdfExporter.TryReadImageSize(path));
        Assert.Equal(6, JpegExifOrientation.Read(path));
    }

    [Fact]
    public void A_file_that_is_not_there_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, JpegExifOrientation.Read(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.jpg")));
    }
}
