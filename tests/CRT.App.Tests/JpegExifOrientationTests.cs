using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// JpegExifOrientation - which way up a phone photo is meant to be shown (code review, 2026-10-16:
// a portrait photo was shown on its side in the Feedback tab, at a size it would not have in the
// mail). Read from the header only; anything it cannot read is "as stored", never an error.
// ###########################################################################################
public sealed class JpegExifOrientationTests
{
    // ###########################################################################################
    // A JPEG's header as a camera writes it: SOI, an APP1 "Exif" segment whose first IFD holds
    // another tag and then Orientation (0 leaves the segment out), the SOF0 frame header with
    // the size, and EOI - no image data, which neither this reader nor the size reader looks at.
    // ###########################################################################################
    internal static byte[] Jpeg(int width, int height, int orientation, bool bigEndian, ushort orientationType = 3)
    {
        var bytes = new List<byte> { 0xFF, 0xD8 };

        if (orientation != 0)
        {
            var tiff = new List<byte>();
            tiff.AddRange(bigEndian ? "MM"u8.ToArray() : "II"u8.ToArray());
            Add16(tiff, 42, bigEndian);
            Add32(tiff, 8, bigEndian);
            Add16(tiff, 2, bigEndian);

            // Make (ASCII, empty), then Orientation (SHORT, its value in the first two bytes).
            Add16(tiff, 0x010F, bigEndian); Add16(tiff, 2, bigEndian); Add32(tiff, 1, bigEndian); Add32(tiff, 0, bigEndian);
            Add16(tiff, 0x0112, bigEndian); Add16(tiff, orientationType, bigEndian); Add32(tiff, 1, bigEndian); Add16(tiff, orientation, bigEndian); Add16(tiff, 0, bigEndian);
            Add32(tiff, 0, bigEndian);

            byte[] payload = [.. "Exif\0\0"u8.ToArray(), .. tiff];
            int length = payload.Length + 2;
            bytes.AddRange([0xFF, 0xE1, (byte)(length >> 8), (byte)length]);
            bytes.AddRange(payload);
        }

        bytes.AddRange([0xFF, 0xC0, 0x00, 0x11, 0x08, (byte)(height >> 8), (byte)height, (byte)(width >> 8), (byte)width, 0x03, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1]);
        bytes.AddRange([0xFF, 0xD9]);

        return bytes.ToArray();
    }

    private static void Add16(List<byte> bytes, int value, bool bigEndian)
    {
        byte high = (byte)(value >> 8), low = (byte)value;
        bytes.AddRange(bigEndian ? [high, low] : [low, high]);
    }

    private static void Add32(List<byte> bytes, long value, bool bigEndian)
    {
        byte[] four = [(byte)(value >> 24), (byte)(value >> 16), (byte)(value >> 8), (byte)value];
        bytes.AddRange(bigEndian ? four : four.Reverse());
    }

    private static int Read(byte[] bytes) => JpegExifOrientation.Read(new MemoryStream(bytes));

    [Theory]
    [InlineData(1, false)] [InlineData(2, false)] [InlineData(3, false)] [InlineData(4, false)]
    [InlineData(5, false)] [InlineData(6, false)] [InlineData(7, false)] [InlineData(8, false)]
    [InlineData(1, true)] [InlineData(2, true)] [InlineData(3, true)] [InlineData(4, true)]
    [InlineData(5, true)] [InlineData(6, true)] [InlineData(7, true)] [InlineData(8, true)]
    public void Each_orientation_is_read_in_either_byte_order(int orientation, bool bigEndian)
    {
        Assert.Equal(orientation, Read(Jpeg(640, 480, orientation, bigEndian)));
    }

    [Fact]
    public void A_JPEG_without_EXIF_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read(Jpeg(640, 480, orientation: 0, bigEndian: false)));
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
        Assert.Equal(JpegExifOrientation.AsStored, Read(Jpeg(640, 480, orientation, bigEndian: false)));
    }

    // The tag must be a SHORT; a LONG in its place is a file nobody wrote right, shown as stored.
    [Fact]
    public void An_orientation_that_is_not_a_SHORT_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, Read(Jpeg(640, 480, 6, bigEndian: false, orientationType: 4)));
    }

    // Cut short anywhere, the header is read as far as it goes and never throws.
    [Fact]
    public void A_header_cut_short_anywhere_is_read_without_an_error()
    {
        byte[] whole = Jpeg(640, 480, 6, bigEndian: true);

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
        File.WriteAllBytes(path, Jpeg(4032, 3024, 6, bigEndian: false));

        Assert.Equal((4032, 3024), WorkbookPdfExporter.TryReadImageSize(path));
        Assert.Equal(6, JpegExifOrientation.Read(path));
    }

    [Fact]
    public void A_file_that_is_not_there_is_as_stored()
    {
        Assert.Equal(JpegExifOrientation.AsStored, JpegExifOrientation.Read(Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.jpg")));
    }
}
