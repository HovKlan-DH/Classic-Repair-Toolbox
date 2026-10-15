using System;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What a draft holds that is not a row - files and KiCad calibrations (code review, 2026-10-13).
//
// *** THE BUG THIS EXISTS FOR. *** The Drafts row counted rows only, so a draft whose only work
// was re-calibrating a KiCad overlay, replacing a picture under its own name, or adding KiCad files
// read "0 rows changed" and could not be submitted ("Add some data to this board first") - though
// the Wiki's KiCad walkthrough sends exactly such a draft. These pin that such work is counted, and
// that what the PUBLISHED update replaced since is not counted as the contributor's.
//
// Shares "BoardData": the workbooks are read through BoardDataReader.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftOtherChangesTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";
    private const string Image = "Commodore/C64/250407/Sheet1.png";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string PublishedWorkbook => DraftBoardSource.PublishedPathOf(this.DataRoot, BoardKey);

    private string DraftWorkbook => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private void PublishAndSeed()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.PublishedWorkbook)!);
        BoardWorkbookWriter.Write(this.PublishedWorkbook, new BoardData
        {
            RevisionDate = "2026-09-01",
            Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = Image }],
            Components = [new ComponentEntry { BoardLabel = "U8" }],
        });
        this.WritePublished(Image, "sheet v1");
        BoardDataReader.ClearAllCache();

        Assert.True(DraftSeeder.SeedFromPublishedFile(this.DraftsRoot, this.DataRoot, BoardKey).Created);
    }

    private void WritePublished(string relative, string text)
    {
        string full = Path.Combine(this.DataRoot, relative.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        File.WriteAllText(full, text);
    }

    private DraftOtherChanges Count() =>
        DraftStatusReader.CountOtherChangesCached(
            DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, BoardKey),
            this.DraftsRoot);

    [Fact]
    public void A_fresh_draft_holds_nothing_but_rows()
    {
        this.PublishAndSeed();

        Assert.False(this.Count().Any);
    }

    [Fact]
    public void A_picture_replaced_under_its_own_name_is_a_change()
    {
        this.PublishAndSeed();

        File.WriteAllText(DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, Image), "sheet, better scan");

        Assert.Equal(new DraftOtherChanges(1, 0), this.Count());
    }

    [Fact]
    public void KiCad_files_added_to_the_draft_are_a_change()
    {
        this.PublishAndSeed();

        string kiCad = Path.Combine(Path.GetDirectoryName(this.DraftWorkbook)!, "KiCad data", "board.kicad_pcb");
        Directory.CreateDirectory(Path.GetDirectoryName(kiCad)!);
        File.WriteAllText(kiCad, "(kicad_pcb)");

        Assert.Equal(1, this.Count().Files);
    }

    [Fact]
    public void A_KiCad_calibration_is_a_change()
    {
        this.PublishAndSeed();

        BoardComponentHighlightStorage.SaveKiCadCalibration(this.DraftWorkbook, "Sheet 1", "pcb", 1, 2, 1.5, 1.5, false, false);

        Assert.Equal(new DraftOtherChanges(0, 1), this.Count());
    }

    // The published update replacing a picture is not the contributor's change - measured against
    // the base, the draft's untouched copy is still what it was copied from.
    [Fact]
    public void A_picture_the_update_replaced_is_not_counted_as_the_contributors()
    {
        this.PublishAndSeed();

        this.WritePublished(Image, "sheet v2 from somebody else");

        Assert.False(this.Count().Any);
    }

    // A draft seeded before bases existed is measured against the published copy as it is now -
    // the rule retirement uses.
    [Fact]
    public void A_draft_without_a_base_is_measured_against_the_published_copy()
    {
        this.PublishAndSeed();
        File.Delete(DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey));

        Assert.False(this.Count().Any);

        File.WriteAllText(DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, Image), "sheet, better scan");

        Assert.Equal(1, this.Count().Files);
    }

    // ###########################################################################################
    // *** AN UNTOUCHED COPY IS NOT READ (code review, 2026-10-14). *** The Drafts tab counts this on
    // the UI thread whenever CRT's window comes to the front, and for a draft without a base it
    // hashed EVERY file on both sides - a C64-sized draft is about 1,100 files a side - freezing the
    // window for seconds. A seeded copy keeps the published file's length AND write time (a file
    // copy carries the write time across), so such a pair is the same without a byte being read;
    // only a file whose stamp moved is hashed.
    // ###########################################################################################
    [Fact]
    public void A_draft_without_a_base_reads_no_copy_that_still_has_the_published_length_and_write_time()
    {
        this.PublishAndSeed();
        File.Delete(DraftFolderLayout.GetBasePath(this.DraftsRoot, BoardKey));

        int before = DraftFingerprint.FilesHashedForTests;

        Assert.False(this.Count().Any);
        Assert.Equal(before, DraftFingerprint.FilesHashedForTests);

        // Replaced with bytes of the same length: its write time moved, so it is read - and counted.
        string copy = DraftFolderLayout.GetReferencedFilePath(this.DraftsRoot, BoardKey, Image);
        File.WriteAllText(copy, "sheet v9");
        File.SetLastWriteTimeUtc(copy, DateTime.UtcNow.AddMinutes(-1));

        Assert.Equal(1, this.Count().Files);
    }

    [Fact]
    public void Bookkeeping_and_lock_files_are_not_the_contributors_work()
    {
        this.PublishAndSeed();
        string folder = Path.GetDirectoryName(this.DraftWorkbook)!;

        File.WriteAllText(Path.Combine(folder, "~$Data C64 250407.xlsx"), "excel owner file");
        File.WriteAllText(Path.Combine(folder, ".crt-writing-Data C64 250407.xlsx"), "stranded temporary");

        Assert.False(this.Count().Any);
    }

    [Theory]
    [InlineData(1, 0, "1 file changed")]
    [InlineData(3, 0, "3 files changed")]
    [InlineData(0, 1, "KiCad calibration changed")]
    [InlineData(0, 2, "2 KiCad calibrations changed")]
    [InlineData(2, 1, "2 files and a KiCad calibration changed")]
    [InlineData(1, 3, "1 file and 3 KiCad calibrations changed")]
    [InlineData(0, 0, "")]
    public void The_changes_are_described_in_plain_words(int files, int calibrations, string expected)
    {
        Assert.Equal(expected, new DraftOtherChanges(files, calibrations).Describe());
    }
}
