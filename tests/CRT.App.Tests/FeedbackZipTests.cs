using System.IO.Compression;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// FeedbackZip - how the Feedback tab's zip is made (taken out of TabFeedback, code review
// 2026-10-16): each entry's name, how hard it is compressed, and the pictures of the text at the
// zip's root under the names their markers say.
// ###########################################################################################
public sealed class FeedbackZipTests
{
    // The zip written by `write`, read back as entry name -> (bytes, compressed length).
    private static Dictionary<string, (byte[] Bytes, long Compressed)> Zip(Action<ZipArchive> write)
    {
        using var stream = new MemoryStream();

        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            write(archive);

        stream.Position = 0;
        using var read = new ZipArchive(stream, ZipArchiveMode.Read);

        return read.Entries.ToDictionary(entry => entry.FullName, entry =>
        {
            using var content = new MemoryStream();
            using Stream source = entry.Open();
            source.CopyTo(content);
            return (content.ToArray(), entry.CompressedLength);
        });
    }

    // A picture in the text (2026-10-15) is a marker line in the mail and a file in the zip: the
    // entry carries the name the marker says, at the zip's root, so the two meet.
    [Fact]
    public void A_pasted_picture_goes_into_the_zip_under_the_name_its_marker_says()
    {
        string name = FeedbackComposition.PastedName(2);
        byte[] bytes = [137, 80, 78, 71, 1, 2, 3];

        var zip = Zip(archive => FeedbackZip.AddPicture(archive, FeedbackZipPicture.FromBytes(name, bytes)));

        (string entryName, (byte[] read, _)) = Assert.Single(zip);
        Assert.Equal("pasted-image-2.png", entryName);
        Assert.Equal("[pasted-image-2.png]", FeedbackComposition.Marker(entryName));
        Assert.Equal(bytes, read);
    }

    // A dropped picture is streamed from its file - the bytes as they are on the disk, counted
    // for the progress as they go in.
    [Fact]
    public void A_dropped_picture_goes_in_from_its_file()
    {
        using var workspace = new TempWorkspace();
        string path = Path.Combine(workspace.Root, "photo.jpg");
        byte[] bytes = Enumerable.Range(0, 200_000).Select(i => (byte)i).ToArray();
        File.WriteAllBytes(path, bytes);
        long counted = 0;

        var zip = Zip(archive => FeedbackZip.AddPicture(archive, FeedbackZipPicture.FromFile("photo.jpg", path), added => counted += added));

        Assert.Equal(bytes, zip["photo.jpg"].Bytes);
        Assert.Equal(bytes.Length, counted);
        Assert.Equal(bytes.Length, FeedbackZip.SizeOf(FeedbackZipPicture.FromFile("photo.jpg", path)));
    }

    // A dropped picture whose file has gone throws - the send stops on it - and leaves no entry.
    [Fact]
    public void A_dropped_picture_whose_file_has_gone_throws_and_adds_nothing()
    {
        using var workspace = new TempWorkspace();
        string gone = Path.Combine(workspace.Root, "gone.png");

        var zip = Zip(archive => Assert.ThrowsAny<IOException>(() => FeedbackZip.AddPicture(archive, FeedbackZipPicture.FromFile("gone.png", gone))));

        Assert.Empty(zip);
        Assert.Equal(0, FeedbackZip.SizeOf(FeedbackZipPicture.FromFile("gone.png", gone)));
    }

    // A picture is compressed already: deflating it again only costs time under the overlay.
    [Theory]
    [InlineData("pasted-image-1.png", CompressionLevel.NoCompression)]
    [InlineData("photo.JPG", CompressionLevel.NoCompression)]
    [InlineData("anim.gif", CompressionLevel.NoCompression)]
    [InlineData("shot.webp", CompressionLevel.NoCompression)]
    [InlineData("scan.bmp", CompressionLevel.Optimal)]
    [InlineData("CRT.log", CompressionLevel.Optimal)]
    [InlineData("board/data.xlsx", CompressionLevel.Optimal)]
    public void A_picture_is_stored_and_everything_else_deflated(string entryName, CompressionLevel expected)
    {
        Assert.Equal(expected, FeedbackZip.CompressionFor(entryName));
    }

    // Stored means stored: a PNG's entry is no smaller than the PNG.
    [Fact]
    public void A_pictures_entry_is_stored_as_it_is()
    {
        byte[] compressible = new byte[50_000];

        var zip = Zip(archive => FeedbackZip.AddPicture(archive, FeedbackZipPicture.FromBytes("pasted-image-1.png", compressible)));

        Assert.Equal(compressible.Length, zip["pasted-image-1.png"].Compressed);
    }

    // Two zip entries of one name: the server took the first for the picture at the marker, or
    // saved one file and could not save the other. A file goes in under a name of its own.
    [Fact]
    public void No_two_zip_entries_share_a_name()
    {
        List<(string Source, string ZipEntryName)> named = FeedbackZip.WithUniqueEntryNames(
            ["shot.png", "pasted-image-1.png"],
            [
                ("CRT.log", "CRT.log"),
                ("one/shot.png", "shot.png"),
                ("two/SHOT.png", "SHOT.png"),
                ("logs/a.txt", "logs/a.txt"),
                ("other/logs/a.txt", "logs/a.txt"),
                ("notes.txt", "notes.txt"),
            ]);

        Assert.Equal(["CRT.log", "shot (2).png", "SHOT (3).png", "logs/a.txt", "logs/a (2).txt", "notes.txt"], named.Select(file => file.ZipEntryName));
        Assert.Equal(["CRT.log", "one/shot.png", "two/SHOT.png", "logs/a.txt", "other/logs/a.txt", "notes.txt"], named.Select(file => file.Source));
    }

    // Two folders of one name, every file clashing: each still gets a name of its own.
    [Fact]
    public void Two_folders_of_one_name_go_in_side_by_side()
    {
        var files = Enumerable.Range(0, 2000).Select(i => ($"a/{i}.txt", $"logs/{i}.txt"))
            .Concat(Enumerable.Range(0, 2000).Select(i => ($"b/{i}.txt", $"logs/{i}.txt")))
            .ToList();

        List<(string Source, string ZipEntryName)> named = FeedbackZip.WithUniqueEntryNames([], files);

        Assert.Equal(4000, named.Select(file => file.ZipEntryName).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal("logs/7 (2).txt", named.Single(file => file.Source == "b/7.txt").ZipEntryName);
    }

    [Fact]
    public void A_zip_picture_is_bytes_or_a_file()
    {
        Assert.Throws<ArgumentNullException>(() => FeedbackZipPicture.FromBytes("a.png", null!));
        Assert.Throws<ArgumentException>(() => FeedbackZipPicture.FromFile("a.png", " "));
        Assert.Throws<ArgumentException>(() => FeedbackZipPicture.FromBytes(" ", [1]));
    }
}
