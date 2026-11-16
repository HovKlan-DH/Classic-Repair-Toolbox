using Avalonia;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The "Workbooks" tab's board pane reads its schematic images through SchematicPreviewPictures (2026-11-16, performance
// walkthrough): decoded at most Width wide, once while the file stays as it is, with the picture's OWN size beside it -
// every worklog area is kept in the picture's own pixels. HeadlessUi: a picture is decoded as an Avalonia bitmap, and
// the reads it keeps are shared by every test of the process.
//
// The decoded bitmap's own size is the platform's - the suite's headless drawing decodes nothing, every bitmap 1x1 - so
// what these hold is the size carried beside it, read off the picture's header.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SchematicPreviewPicturesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // A minimal valid 1x1 PNG (the same bytes SchematicThumbnailSetTests uses).
    private static readonly byte[] OnePixelPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

    // The 1x1 PNG with its header saying `width` x `height`, the header's checksum made again.
    private string WritePng(string name, int width, int height)
    {
        byte[] bytes = (byte[])OnePixelPng.Clone();

        for (int at = 0; at < 4; at++)
        {
            bytes[16 + at] = (byte)(width >> (24 - (8 * at)));
            bytes[20 + at] = (byte)(height >> (24 - (8 * at)));
        }

        // The IHDR chunk's CRC covers its type and data: bytes 12 to 28, written at 29.
        uint crc = Crc32(bytes.AsSpan(12, 17));

        for (int at = 0; at < 4; at++)
        {
            bytes[29 + at] = (byte)(crc >> (24 - (8 * at)));
        }

        string path = this.thisWorkspace.Path_(name);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    private static uint Crc32(ReadOnlySpan<byte> data)
    {
        uint crc = 0xFFFFFFFF;

        foreach (byte b in data)
        {
            crc ^= b;

            for (int bit = 0; bit < 8; bit++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }

        return ~crc;
    }

    private static SchematicPreviewPicture ReadDone(string path)
    {
        Task<SchematicPreviewPicture?> reading = SchematicPreviewPictures.Read(path);
        Assert.True(reading.Wait(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken), "The picture was not read in time");
        return Assert.IsType<SchematicPreviewPicture>(reading.Result);
    }

    // A 4220x2941 schematic was decoded whole for a pane showing it at most 420 high - some 47 MB each. Decoded smaller,
    // it keeps the size every worklog area is mapped by.
    [Fact]
    public void A_picture_carries_its_own_size_however_small_it_was_decoded()
    {
        UiTest.Run(() =>
        {
            Assert.Equal(new PixelSize(4220, 2941), ReadDone(this.WritePng("wide.png", 4220, 2941)).OriginalPixelSize);
            Assert.Equal(new PixelSize(400, 300), ReadDone(this.WritePng("narrow.png", 400, 300)).OriginalPixelSize);
        });
    }

    [Fact]
    public void The_pane_decodes_a_picture_at_most_1600_wide() =>
        Assert.Equal(1600, SchematicPreviewPictures.Width);

    // Every board switch and every visit to the tab decoded the pane's pictures again.
    [Fact]
    public void A_picture_is_read_once_while_its_file_stays_as_it_is()
    {
        UiTest.Run(() =>
        {
            string path = this.WritePng("once.png", 300, 200);
            ReadDone(path);
            int readsBefore = SchematicPreviewPictures.ReadsMadeForTests;

            ReadDone(path);
            ReadDone(path);

            Assert.Equal(readsBefore, SchematicPreviewPictures.ReadsMadeForTests);

            // Replaced: read again.
            this.WritePng("once.png", 320, 200);
            File.SetLastWriteTimeUtc(path, File.GetLastWriteTimeUtc(path).AddMinutes(1));

            Assert.Equal(new PixelSize(320, 200), ReadDone(path).OriginalPixelSize);
            Assert.Equal(readsBefore + 1, SchematicPreviewPictures.ReadsMadeForTests);
        });
    }

    // ###########################################################################################
    // *** A read that came to nothing is not kept (code review, 2026-11-16). *** A picture held a moment by a virus scanner as
    // the pane first read it stayed "could not be found" until the file changed or 24 others were read. Here the file is held
    // open by another as it is first read; let go of - the very same file, length and write time - the next ask reads it.
    // ###########################################################################################
    [Fact]
    public void A_picture_that_could_not_be_read_is_read_again_at_the_next_ask()
    {
        UiTest.Run(() =>
        {
            string path = this.WritePng("held.png", 300, 200);
            Task<SchematicPreviewPicture?> first;

            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                first = SchematicPreviewPictures.Read(path);
                Assert.True(first.Wait(TimeSpan.FromSeconds(10), TestContext.Current.CancellationToken));
            }

            Assert.Null(first.Result);
            Assert.True(SpinWait.SpinUntil(() => !ReferenceEquals(SchematicPreviewPictures.Read(path), first), TimeSpan.FromSeconds(10)));
            Assert.Equal(new PixelSize(300, 200), ReadDone(path).OriginalPixelSize);
        });
    }

    [Fact]
    public void A_picture_not_there_reads_as_nothing_at_once()
    {
        UiTest.Run(() =>
        {
            Task<SchematicPreviewPicture?> reading = SchematicPreviewPictures.Read(this.thisWorkspace.Path_("gone.png"));

            Assert.True(reading.IsCompletedSuccessfully);
            Assert.Null(reading.Result);
        });
    }
}
