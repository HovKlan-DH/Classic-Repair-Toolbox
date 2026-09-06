using System.IO.Compression;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The ZIP half of the workbook export - the PDF plus the workbook's original photos and files.
//
// WHY THIS EXISTS AT ALL, given that WorkbookPdfExporter is documented as deliberately untested
// (asserting on PDF bytes tests QuestPDF rather than this app): the ARCHIVE is not PDF bytes. It is
// this app deciding what goes in, under what name, and how a collision is resolved - and the first
// shipped version of it crashed the whole application on any workbook that had a single attachment.
//
// `ZipArchive` in `Create` mode is write-forward only: `GetEntry` and `Entries` both throw
// `NotSupportedException("Cannot access entries in Create mode")`. The collision check called
// `GetEntry`, so the crash needed an attachment to reach it - and the one manual probe run before
// shipping used a workbook with none. That is exactly the shape of bug a test is for, so the whole
// file is built around workbooks that HAVE attachments.
//
// Needs QuestPDF's licence set, since WriteZip generates the PDF it packs - see the constructor.
public class WorkbookZipExportTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public WorkbookZipExportTests()
    {
        // Set at startup in the real app (App.OnFrameworkInitializationCompleted), which no test
        // runs. Without it QuestPDF throws on the first generation.
        QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;
    }

    public void Dispose() => this.thisWorkspace.Dispose();

    private static WorkbookRecord Workbook() => new()
    {
        Id = 3,
        BoardKey = "C64|250469",
        Title = "Repair job",
        Status = "Open",
        StartDate = new DateTime(2026, 9, 3)
    };

    private static WorklogEntryRecord Entry(int id, string title) => new()
    {
        Id = id,
        SchematicName = "Sheet 1",
        Title = title,
        Category = "Issue",
        State = "Open"
    };

    // An entry's attachment folder, with the named files really written to it - ResolveAttachments
    // drops anything not on disk, so a fixture that only added records would silently test nothing.
    private string WriteAttachments(int entryId, params string[] fileNames)
    {
        string folder = Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}");
        Directory.CreateDirectory(folder);

        foreach (string name in fileNames)
            File.WriteAllText(Path.Combine(folder, name), $"contents of {name}");

        return folder;
    }

    private WorkbookExportModel.Document BuildDocument(params WorklogEntryRecord[] entries) =>
        WorkbookExportModel.Build(
            Workbook(),
            entries,
            null,
            entryId => Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}"),
            new DateTime(2026, 9, 4),
            "DKK");

    private string ZipPath() => Path.Combine(this.thisWorkspace.Root, "export.zip");

    // ###########################################################################################
    // THE REGRESSION TEST for the reported crash: exporting a workbook with an attachment threw
    // NotSupportedException out of the background write and took the application down.
    //
    // Fails against the version that called ZipArchive.GetEntry to test for a name collision.
    // ###########################################################################################
    [Fact]
    public void A_workbook_with_attachments_exports_without_throwing()
    {
        var entry = Entry(1, "Bad cap");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "board.png" });
        this.WriteAttachments(1, "board.png");

        string path = this.ZipPath();

        // The assertion is that this does not throw; the archive contents are checked below.
        WorkbookPdfExporter.WriteZip(this.BuildDocument(entry), path);

        Assert.True(File.Exists(path));
    }

    [Fact]
    public void The_archive_holds_the_pdf_and_every_attachment()
    {
        var entry = Entry(1, "Bad cap");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "board.png" });
        entry.Files.Add(new WorklogAttachmentRecord { Id = 1, FileName = "datasheet.txt" });
        this.WriteAttachments(1, "board.png", "datasheet.txt");

        string path = this.ZipPath();
        WorkbookPdfExporter.WriteZip(this.BuildDocument(entry), path);

        using var archive = ZipFile.OpenRead(path);
        var names = archive.Entries.Select(e => e.FullName).ToList();

        // "Workbook_{id}_{Hardware}_{Board}_{YYYYMMDD}.pdf" - it is the file that gets forwarded
        // once someone unpacks the archive, so it has to identify itself on its own.
        Assert.Contains("Workbook_3_C64_250469_20260904.pdf", names);

        // Attachments sit under "worklog_{id}" - the SAME folder name the entry's attachments have
        // in the local Workbooks folder, so what the recipient unpacks matches what the repairer
        // sees on their own disk. The worklog title is deliberately not in the folder name.
        Assert.Contains("worklog_1/board.png", names);
        Assert.Contains("worklog_1/datasheet.txt", names);
    }

    // ###########################################################################################
    // Two entries holding a file of the SAME name both survive, the second suffixed.
    //
    // This is the case the crashing collision check existed for in the first place: "front.jpg"
    // photographed for two different faults is entirely ordinary. They land in different folders
    // here, so they do not actually collide - which is the point: the check must handle the common
    // case without throwing, not only the rare one.
    // ###########################################################################################
    [Fact]
    public void Two_entries_with_identically_named_files_both_survive()
    {
        var first = Entry(1, "Bad cap");
        first.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "front.jpg" });
        this.WriteAttachments(1, "front.jpg");

        var second = Entry(2, "Cracked trace");
        second.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "front.jpg" });
        this.WriteAttachments(2, "front.jpg");

        string path = this.ZipPath();
        WorkbookPdfExporter.WriteZip(this.BuildDocument(first, second), path);

        using var archive = ZipFile.OpenRead(path);
        var names = archive.Entries.Select(e => e.FullName).ToList();

        Assert.Contains("worklog_1/front.jpg", names);
        Assert.Contains("worklog_2/front.jpg", names);
    }

    // ###########################################################################################
    // A genuine collision - the same file name twice WITHIN one entry, which entries.json can hold
    // (a photo row and a file row can name the same file) - keeps both, the second suffixed rather
    // than overwriting the first or producing a duplicate path most unzip tools refuse.
    // ###########################################################################################
    [Fact]
    public void A_name_used_twice_in_one_entry_is_suffixed_rather_than_overwritten()
    {
        var entry = Entry(1, "Bad cap");
        entry.Photos.Add(new WorklogAttachmentRecord { Id = 1, FileName = "shot.png" });
        entry.Files.Add(new WorklogAttachmentRecord { Id = 1, FileName = "shot.png" });
        this.WriteAttachments(1, "shot.png");

        string path = this.ZipPath();
        WorkbookPdfExporter.WriteZip(this.BuildDocument(entry), path);

        using var archive = ZipFile.OpenRead(path);
        var names = archive.Entries.Select(e => e.FullName).ToList();

        Assert.Contains("worklog_1/shot.png", names);
        Assert.Contains("worklog_1/shot-2.png", names);

        // Both are real files, not one empty placeholder.
        Assert.All(archive.Entries.Where(e => e.FullName.Contains("shot")), e => Assert.True(e.Length > 0));
    }

    // A workbook with no attachments at all still produces an archive holding the PDF - the case
    // that DID work before the fix, kept so it cannot break in the other direction.
    [Fact]
    public void A_workbook_with_no_attachments_still_produces_an_archive_with_the_pdf()
    {
        string path = this.ZipPath();
        WorkbookPdfExporter.WriteZip(this.BuildDocument(Entry(1, "Just a note")), path);

        using var archive = ZipFile.OpenRead(path);
        Assert.Single(archive.Entries);
        Assert.EndsWith(".pdf", archive.Entries[0].FullName);
    }

    // Exporting twice to the same path overwrites rather than failing or appending - the save
    // dialog already asked the user about overwriting by the time this runs.
    [Fact]
    public void Exporting_twice_to_the_same_path_overwrites_the_earlier_archive()
    {
        string path = this.ZipPath();
        var document = this.BuildDocument(Entry(1, "Just a note"));

        WorkbookPdfExporter.WriteZip(document, path);
        WorkbookPdfExporter.WriteZip(document, path);

        using var archive = ZipFile.OpenRead(path);
        Assert.Single(archive.Entries);
    }

    // ###########################################################################################
    // EnsureIconFontLoaded is safe to call when Avalonia is NOT available, and leaves the export
    // working without icons rather than throwing.
    //
    // This is the guard on a bug that shipped silently: the icon font is an Avalonia resource, and
    // AssetLoader cannot resolve its service on an arbitrary background thread - which is exactly
    // where the export's write runs. The load threw INSIDE a try, so every exported document simply
    // came out with no icons at all and nothing said so. The fix moved the byte-reading to the UI
    // thread; this test pins the other half of the contract - that the failure path is still a
    // working export.
    //
    // No UI thread here (this class is not in the HeadlessUi collection), so this exercises exactly
    // the "Avalonia unavailable" case.
    // ###########################################################################################
    [Fact]
    public void The_icon_font_loader_is_safe_without_avalonia_and_the_export_still_works()
    {
        var thrown = Record.Exception(WorkbookPdfExporter.EnsureIconFontLoaded);
        Assert.Null(thrown);

        string path = this.ZipPath();
        WorkbookPdfExporter.WriteZip(this.BuildDocument(Entry(1, "Just a note")), path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // ###########################################################################################
    // An entry whose marked area is SHOWN exports without throwing.
    //
    // THE REGRESSION THIS HOLDS: drawing the area and its "#N" badge over the schematic means
    // building a layer stack whose bands are sized proportionally. Two separate versions of that
    // produced a container with zero width or zero height, and QuestPDF answers a zero-sized
    // container that must hold text by throwing DocumentLayoutException and abandoning the WHOLE
    // document - not by clipping one pill. So a single bad band takes down the entire export.
    //
    // Every other test in this file uses entries with no drawn area at all, which take the parked
    // path and never touch that code - which is exactly how the first version of it shipped.
    [Fact]
    public void An_entry_with_a_shown_marked_area_exports_without_throwing()
    {
        var entry = Entry(1, "Bad cap");
        entry.ShowMarkedArea = true;
        entry.AreaX = 2060;
        entry.AreaY = 285;
        entry.AreaWidth = 373;
        entry.AreaHeight = 527;

        string schematic = this.WriteSchematicImage(3552, 2477);

        string path = this.ZipPath();

        WorkbookPdfExporter.WriteZip(
            WorkbookExportModel.Build(
                Workbook(),
                new[] { entry },
                new Dictionary<string, string> { [entry.SchematicName] = schematic },
                entryId => Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}"),
                new DateTime(2026, 9, 4),
                "DKK"),
            path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // An area flush against the image's right and bottom edges - the corner case where the
    // "what is left over" bands come out at zero and a naive layout asks QuestPDF for a
    // zero-weight item, which it rejects.
    [Fact]
    public void An_area_touching_the_far_edges_exports_without_throwing()
    {
        var entry = Entry(1, "Corner");
        entry.ShowMarkedArea = true;
        entry.AreaX = 0;
        entry.AreaY = 0;
        entry.AreaWidth = 1000;
        entry.AreaHeight = 800;

        string schematic = this.WriteSchematicImage(1000, 800);

        string path = this.ZipPath();

        WorkbookPdfExporter.WriteZip(
            WorkbookExportModel.Build(
                Workbook(),
                new[] { entry },
                new Dictionary<string, string> { [entry.SchematicName] = schematic },
                entryId => Path.Combine(this.thisWorkspace.Root, $"worklog_{entryId}"),
                new DateTime(2026, 9, 4),
                "DKK"),
            path);

        Assert.True(new FileInfo(path).Length > 0);
    }

    // A minimal real PNG, written by hand so the exporter's header reader sees genuine dimensions
    // (it parses the IHDR rather than decoding, so the pixel data can be a single transparent dot).
    private string WriteSchematicImage(int width, int height)
    {
        string path = Path.Combine(this.thisWorkspace.Root, $"schematic-{width}x{height}.png");
        File.WriteAllBytes(path, BuildPng(width, height));
        return path;
    }

    private static byte[] BuildPng(int width, int height)
    {
        static byte[] Chunk(string type, byte[] data)
        {
            var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
            var payload = typeBytes.Concat(data).ToArray();

            return BitConverter.GetBytes(data.Length).Reverse()
                .Concat(payload)
                .Concat(BitConverter.GetBytes(Crc32(payload)).Reverse())
                .ToArray();
        }

        var ihdr = BitConverter.GetBytes(width).Reverse()
            .Concat(BitConverter.GetBytes(height).Reverse())
            .Concat(new byte[] { 8, 6, 0, 0, 0 })
            .ToArray();

        // One fully transparent row per line, deflate-compressed - the smallest valid image data.
        var raw = new byte[height * (1 + width * 4)];
        byte[] compressed;

        using (var buffer = new MemoryStream())
        {
            using (var deflate = new System.IO.Compression.ZLibStream(
                buffer, System.IO.Compression.CompressionLevel.Fastest, leaveOpen: true))
            {
                deflate.Write(raw, 0, raw.Length);
            }

            compressed = buffer.ToArray();
        }

        return new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }
            .Concat(Chunk("IHDR", ihdr))
            .Concat(Chunk("IDAT", compressed))
            .Concat(Chunk("IEND", Array.Empty<byte>()))
            .ToArray();
    }

    private static uint Crc32(byte[] data)
    {
        uint crc = 0xFFFFFFFF;

        foreach (byte b in data)
        {
            crc ^= b;

            for (int i = 0; i < 8; i++)
                crc = (crc >> 1) ^ (0xEDB88320 & (uint)(-(crc & 1)));
        }

        return crc ^ 0xFFFFFFFF;
    }

    // ###########################################################################################
    // The click target for a worklog's LINK row.
    //
    // Exercised by reflection because it is private and has no seam of its own - the same approach
    // and the same reasoning as ExternalTargetLauncherTests and OnlineServicesTests. What it
    // decides is worth pinning: a PDF hyperlink with no scheme is silently ignored by every
    // reader, so a stored "example.com" that is not given one produces a link that looks right and
    // does nothing.
    //
    // These rows are DECLARED destinations, unlike a URL spotted inside prose - the add-link
    // dialog stores whatever the user typed without normalising it, so the scheme-less shape
    // genuinely occurs and TextLinkFinder (correctly) refuses to treat it as a link.
    // ###########################################################################################
    private static string? BuildLinkTarget(string url)
    {
        var method = typeof(WorkbookPdfExporter).GetMethod(
            "BuildLinkTarget",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

        return (string?)method.Invoke(null, new object[] { url });
    }

    [Theory]
    [InlineData("https://example.com/74ls08.pdf")]
    [InlineData("http://example.com")]
    [InlineData("HTTPS://EXAMPLE.COM")]
    public void A_url_that_already_has_a_scheme_is_used_as_it_stands(string url)
    {
        Assert.Equal(url, BuildLinkTarget(url));
    }

    // The case the add-link dialog actually produces, since it does not normalise what is typed.
    [Theory]
    [InlineData("example.com/thread/12", "https://example.com/thread/12")]
    [InlineData("www.zimmers.net", "https://www.zimmers.net")]
    public void A_url_with_no_scheme_gets_https(string stored, string expected)
    {
        Assert.Equal(expected, BuildLinkTarget(stored));
    }

    [Fact]
    public void Surrounding_whitespace_is_trimmed_off_the_target()
    {
        Assert.Equal("https://example.com", BuildLinkTarget("  https://example.com  "));
    }

    // Nothing usable means NO link rather than one pointing at nonsense - the row still prints as
    // plain text, so the information is not lost.
    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("see the datasheet")]
    public void A_value_that_cannot_be_a_url_produces_no_target(string url)
    {
        Assert.Null(BuildLinkTarget(url));
    }

    // ###########################################################################################
    // Reading an image's pixel dimensions from its HEADER.
    //
    // WHY THIS MATTERS ENOUGH TO TEST: every marked area in the exported PDF is positioned as a
    // fraction of these two numbers. Get them wrong and the overlay lands nowhere near the copper
    // it marks - with nothing thrown and nothing logged, because a wrong size is still a size.
    //
    // The JPEG side is a marker walk, and the trap is which markers are frame headers. 0xC4, 0xC8
    // and 0xCC sit INSIDE the 0xC0-0xCF run without being frame headers, and 0xCD/0xCE/0xCF
    // (DNL/DHP/EXP) carry no dimensions either - an earlier version treated all of C0-CF except
    // three as frames, so a file carrying any of those before its real SOF0 reported garbage.
    // ###########################################################################################
    private static (int Width, int Height)? TryReadImageSize(string path)
    {
        var method = typeof(WorkbookPdfExporter).GetMethod(
            "TryReadImageSize",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

        return ((int, int)?)method.Invoke(null, new object[] { path });
    }

    private string WriteBytes(string name, byte[] bytes)
    {
        string path = Path.Combine(this.thisWorkspace.Root, name);
        File.WriteAllBytes(path, bytes);
        return path;
    }

    [Fact]
    public void A_png_reports_the_size_from_its_ihdr()
    {
        string path = this.WriteBytes("board.png", BuildPng(4220, 2941));

        Assert.Equal((4220, 2941), TryReadImageSize(path));
    }

    [Fact]
    public void A_plain_jpeg_reports_the_size_from_its_frame_header()
    {
        string path = this.WriteBytes("photo.jpg", BuildJpeg(new byte[0], 0xC0, 1600, 1200));

        Assert.Equal((1600, 1200), TryReadImageSize(path));
    }

    // A progressive JPEG (SOF2) is an ordinary frame header and must be read, not skipped.
    [Fact]
    public void A_progressive_jpeg_reports_its_size()
    {
        string path = this.WriteBytes("progressive.jpg", BuildJpeg(new byte[0], 0xC2, 800, 600));

        Assert.Equal((800, 600), TryReadImageSize(path));
    }

    // ###########################################################################################
    // THE REGRESSION: a segment whose marker falls inside 0xC0-0xCF but is NOT a frame header must
    // be skipped by its length, not read as dimensions.
    //
    // Each of these is placed BEFORE the real SOF0 carrying 1600x1200. Against the version that
    // treated C0-CF (minus three) as frames, the DNL/DHP/EXP cases return the segment's own
    // payload bytes as a size instead.
    // ###########################################################################################
    [Theory]
    [InlineData(0xC4)] // DHT - Huffman tables
    [InlineData(0xCC)] // DAC - arithmetic coding conditioning
    [InlineData(0xCD)] // DNL
    [InlineData(0xCE)] // DHP
    [InlineData(0xCF)] // EXP
    public void A_non_frame_marker_in_the_c0_range_does_not_supply_the_size(int marker)
    {
        // A payload that would read as 0x1111 x 0x2222 if mistaken for a frame header.
        var decoy = new byte[] { 0x08, 0x11, 0x11, 0x22, 0x22, 0x00, 0x00 };

        string path = this.WriteBytes(
            $"decoy-{marker:X2}.jpg",
            BuildJpeg(BuildSegment((byte)marker, decoy), 0xC0, 1600, 1200));

        Assert.Equal((1600, 1200), TryReadImageSize(path));
    }

    // A standalone marker carries NO length word, so reading two bytes after it consumes image
    // data as a length and seeks to an arbitrary offset.
    [Fact]
    public void A_standalone_marker_does_not_derail_the_scan()
    {
        // TEM (0xFF01) followed by bytes that would be a wild length if read as one.
        var tem = new byte[] { 0xFF, 0x01, 0xFF, 0xFF };

        string path = this.WriteBytes("tem.jpg", BuildJpeg(tem, 0xC0, 1024, 768));

        Assert.Equal((1024, 768), TryReadImageSize(path));
    }

    [Fact]
    public void A_file_that_is_not_a_png_or_jpeg_has_no_size()
    {
        Assert.Null(TryReadImageSize(this.WriteBytes("notes.txt",
            System.Text.Encoding.ASCII.GetBytes("this is not an image"))));
    }

    [Fact]
    public void A_missing_file_has_no_size()
    {
        Assert.Null(TryReadImageSize(Path.Combine(this.thisWorkspace.Root, "absent.png")));
    }

    // A JPEG with no frame header at all (truncated before SOF) returns null rather than a guess.
    [Fact]
    public void A_jpeg_with_no_frame_header_has_no_size()
    {
        var bytes = new byte[] { 0xFF, 0xD8 }
            .Concat(BuildSegment(0xE0, new byte[] { 0x00, 0x01, 0x02, 0x03 }))
            .ToArray();

        Assert.Null(TryReadImageSize(this.WriteBytes("headless.jpg", bytes)));
    }

    // One APPn/COMn-style segment: 0xFF, marker, big-endian length (including the length word).
    private static byte[] BuildSegment(byte marker, byte[] payload)
    {
        int length = payload.Length + 2;

        return new byte[] { 0xFF, marker, (byte)(length >> 8), (byte)(length & 0xFF) }
            .Concat(payload)
            .ToArray();
    }

    // SOI, then whatever leading bytes the test wants, then a frame header of the given marker
    // carrying the given dimensions.
    private static byte[] BuildJpeg(byte[] leading, byte frameMarker, int width, int height)
    {
        var frame = new byte[]
        {
            0x08,                                   // sample precision
            (byte)(height >> 8), (byte)(height & 0xFF),
            (byte)(width >> 8), (byte)(width & 0xFF),
            0x03                                    // component count
        };

        return new byte[] { 0xFF, 0xD8 }
            .Concat(leading)
            .Concat(BuildSegment(frameMarker, frame))
            .ToArray();
    }

    // The temp PDF WriteZip generates on its way to the archive is cleaned up, rather than left
    // accumulating in the user's temp folder on every export.
    [Fact]
    public void No_temp_pdf_is_left_behind()
    {
        var before = Directory.GetFiles(Path.GetTempPath(), "*.pdf").Length;

        WorkbookPdfExporter.WriteZip(this.BuildDocument(Entry(1, "Just a note")), this.ZipPath());

        Assert.Equal(before, Directory.GetFiles(Path.GetTempPath(), "*.pdf").Length);
    }
}
