using Avalonia;
using Tabs.TabSchematics;

namespace ClassicRepairToolbox.Tests.Ui;

// SchematicThumbnailSet - a board's thumbnails decoded straight to thumbnail size, stopping when
// the board load is cancelled (code review, 2026-10-13). Bitmaps need Avalonia's platform, so these
// run through UiTest.Run.
[Collection("HeadlessUi")]
public sealed class SchematicThumbnailSetTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // A minimal valid 1x1 PNG (the same bytes WorkbooksBoardPreviewTests uses).
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

    private string WritePng(string name)
    {
        string path = this.thisWorkspace.Path_(name);
        File.WriteAllBytes(path, OnePixelPng);
        return path;
    }

    [Fact]
    public void Every_image_gets_an_entry_in_order_and_a_missing_one_has_no_thumbnail()
    {
        string top = this.WritePng("top.png");
        string missing = this.thisWorkspace.Path_("not there.png");

        UiTest.Run(() =>
        {
            using var set = SchematicThumbnailSet.Decode(
                [("Top", top), ("Gone", missing)],
                "signature",
                800,
                CancellationToken.None);

            Assert.Equal("signature", set.Signature);
            Assert.Equal(["Top", "Gone"], set.Images.Select(image => image.Name));

            Assert.NotNull(set.Images[0].Thumbnail);
            Assert.Equal(new PixelSize(1, 1), set.Images[0].OriginalPixelSize);

            Assert.Null(set.Images[1].Thumbnail);
            Assert.Equal(default, set.Images[1].OriginalPixelSize);
        });
    }

    // A file that is not an image still gets its entry, and the decode goes on: a failed board
    // load is not an option. (Whether it gets a picture is the platform's say - the headless test
    // platform makes a stub of anything; Skia throws, which DecodeOne turns into no thumbnail.)
    [Fact]
    public void A_file_that_is_not_an_image_does_not_stop_the_others()
    {
        string broken = this.thisWorkspace.WriteFile("broken.png", "not a png");
        string top = this.WritePng("top.png");

        UiTest.Run(() =>
        {
            using var set = SchematicThumbnailSet.Decode([("Broken", broken), ("Top", top)], "s", 800, CancellationToken.None);

            Assert.Equal(["Broken", "Top"], set.Images.Select(image => image.Name));
            Assert.NotNull(set.Images[1].Thumbnail);
        });
    }

    // Another board chosen: the decode stops instead of running to its end.
    [Fact]
    public void A_cancelled_decode_stops_and_throws()
    {
        string top = this.WritePng("top.png");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        UiTest.Run(() =>
        {
            Assert.Throws<OperationCanceledException>(() =>
                SchematicThumbnailSet.Decode([("Top", top)], "s", 800, cancelled.Token));
        });
    }

    // The signature moves when an image file does - so a changed image is decoded again.
    [Fact]
    public void The_signature_follows_the_files_on_disk()
    {
        string top = this.WritePng("top.png");
        var images = new List<(string, string)> { ("Top", top) };

        string before = SchematicThumbnailSet.SignatureOf(images);
        Assert.Equal(before, SchematicThumbnailSet.SignatureOf(images));

        File.WriteAllBytes(top, OnePixelPng.Concat(new byte[] { 0 }).ToArray());

        Assert.NotEqual(before, SchematicThumbnailSet.SignatureOf(images));
    }
}
