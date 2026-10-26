using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A 4.0 board's KiCad files copied under their own names (BoardKiCadFolder). KiCad's sheets name
// each other ("Pages/vic.kicad_sch"), so the loaders only see a project whole when its files sit
// side by side under the names the board gives them - not as checksum-named bytes in a store.
// ###########################################################################################
public sealed class BoardKiCadFolderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");
    private string CacheRoot => Path.Combine(this.thisWorkspace.Root, "Cache");
    private string PublishedBoardFile => Path.Combine(this.DataRoot, "Commodore", "C128", "310378", DataTreeLayout.BoardFileName);
    private string DraftFolder => Path.Combine(this.thisWorkspace.Root, "Drafts", "Commodore", "C128", "310378");

    private static string Read(string folder, params string[] parts) =>
        File.ReadAllText(Path.Combine(new[] { folder }.Concat(parts).ToArray()));

    [Fact]
    public void A_published_boards_KiCad_files_are_copied_under_their_names_with_their_pages()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_sch", "root sheet"),
            StoredFiles.Put(this.DataRoot, "KiCad data/Pages/vic.kicad_sch", "vic page"),
            StoredFiles.Put(this.DataRoot, "Schematics/U1.png", "not KiCad"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal(DataTreeUsage.KiCadFolderName, Path.GetFileName(folder));
        Assert.Equal("root sheet", Read(folder, "Open128.kicad_sch"));
        Assert.Equal("vic page", Read(folder, "Pages", "vic.kicad_sch"));
        Assert.Equal(2, Directory.GetFiles(folder, "*", SearchOption.AllDirectories).Length);
    }

    // ###########################################################################################
    // *** A REPLACED COPY NEVER KEEPS THE WRITE TIME OF THE ONE IT REPLACED (2026-10-25). *** The
    // KiCad loader keeps its parse by path and write time (KiCadProjectLoader). File.Copy keeps the
    // source's time, so two different files written in one clock tick - a contributor's copy and the
    // published one, a few milliseconds apart - gave the copy the stamp it already had, and the
    // loader answered with the old file's parse: "Those KiCad files could not be read." after a
    // Remove that had put a good file back (found by a test that failed in some orders only).
    // ###########################################################################################
    [Fact]
    public void A_replaced_copy_gets_a_later_write_time_than_the_one_it_replaced()
    {
        BoardDocument.StoredFileRow first = StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "first bytes");
        BoardDocument.StoredFileRow second = StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "second, other bytes");

        // Both stored in the same clock tick, as a fast machine does.
        var sameTime = new DateTime(2026, 10, 25, 12, 0, 0, DateTimeKind.Utc);
        foreach (BoardDocument.StoredFileRow row in new[] { first, second })
        {
            File.SetLastWriteTimeUtc(Path.Combine(this.DataRoot, DataTreeLayout.StoredFilePath(row.Sha256, row.Name).Replace('/', Path.DirectorySeparatorChar)), sameTime);
        }

        StoredFiles.WriteBoard(this.PublishedBoardFile, first);
        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        DateTime before = File.GetLastWriteTimeUtc(Path.Combine(folder, "Open128.kicad_pcb"));

        StoredFiles.WriteBoard(this.PublishedBoardFile, second);
        BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("second, other bytes", Read(folder, "Open128.kicad_pcb"));
        Assert.True(File.GetLastWriteTimeUtc(Path.Combine(folder, "Open128.kicad_pcb")) > before);
    }

    [Fact]
    public void A_board_with_no_KiCad_files_has_no_folder()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "Schematics/U1.png", "picture"));

        Assert.Equal(string.Empty, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
        Assert.False(Directory.Exists(Path.Combine(this.CacheRoot, BoardKiCadFolder.FolderNameFor(this.PublishedBoardFile))));
    }

    [Fact]
    public void A_missing_board_file_has_no_folder()
    {
        Assert.Equal(string.Empty, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
    }

    // A draft keeps the bytes the tree does not hold in its own store - and those are the ones
    // the contributor imported, so they must be found there first.
    [Fact]
    public void A_listed_files_bytes_are_taken_from_the_drafts_own_store_before_the_trees()
    {
        string draftBoardFile = Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);
        StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "in both");
        StoredFiles.WriteBoard(
            draftBoardFile,
            StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_pcb", "in both"),
            StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_sch", "draft only"));

        string folder = BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("in both", Read(folder, "Open128.kicad_pcb"));
        Assert.Equal("draft only", Read(folder, "Open128.kicad_sch"));
    }

    // A KiCad project imported into a draft lies in "<draft>/KiCad data" until the draft's next
    // write takes it into the store - the traces must show straight away, the new copy winning.
    [Fact]
    public void A_new_file_lying_in_the_draft_folder_is_copied_and_wins_over_a_listed_one()
    {
        string draftBoardFile = Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);
        StoredFiles.WriteBoard(draftBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "published"));
        this.thisWorkspace.WriteFile(Path.Combine(this.DraftFolder, "KiCad data", "Open128.kicad_pcb"), "imported");
        this.thisWorkspace.WriteFile(Path.Combine(this.DraftFolder, "KiCad data", "Pages", "cpu.kicad_sch"), "new page");

        string folder = BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("imported", Read(folder, "Open128.kicad_pcb"));
        Assert.Equal("new page", Read(folder, "Pages", "cpu.kicad_sch"));
    }

    [Fact]
    public void A_file_no_longer_listed_is_removed_from_the_copy_and_a_changed_one_copied_again()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_sch", "first"),
            StoredFiles.Put(this.DataRoot, "KiCad data/Pages/old.kicad_sch", "dropped later"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_sch", "second"));

        Assert.Equal(folder, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
        Assert.Equal("second", Read(folder, "Open128.kicad_sch"));
        Assert.False(File.Exists(Path.Combine(folder, "Pages", "old.kicad_sch")));
        Assert.False(Directory.Exists(Path.Combine(folder, "Pages")));
    }

    // An unchanged entry is not copied again - the KiCad loaders' own cache is keyed by each
    // file's write time, and a fresh copy on every board shown would parse a 45 MB PCB each time.
    [Fact]
    public void An_unchanged_file_is_not_copied_again()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        string copy = Path.Combine(folder, "Open128.kicad_pcb");
        DateTime earlier = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(copy, earlier);

        BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal(earlier, File.GetLastWriteTimeUtc(copy));
    }

    // A copy deleted from the cache (the operating system tidying its temporary folder) comes back.
    [Fact]
    public void A_copy_deleted_from_the_cache_is_copied_again()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        File.Delete(Path.Combine(folder, "Open128.kicad_pcb"));

        BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("pcb", Read(folder, "Open128.kicad_pcb"));
    }

    // A board file can come from a draft folder put in place by hand: a name climbing out of
    // "KiCad data" must never be written outside the copy's own folder.
    [Fact]
    public void A_name_climbing_out_of_the_KiCad_folder_is_not_copied()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"),
            StoredFiles.Put(this.DataRoot, "KiCad data/../../outside.kicad_pcb", "escaped"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Single(Directory.GetFiles(folder, "*", SearchOption.AllDirectories));
        Assert.Empty(Directory.GetFiles(this.thisWorkspace.Root, "outside.kicad_pcb", SearchOption.AllDirectories));
    }

    // Only the types the KiCad loaders read are copied (code review, 2026-10-25): a converted
    // board's generated 21 MB "KiCad-traces.json" was copied for nobody, on the UI thread.
    [Fact]
    public void Only_the_KiCad_files_the_loaders_read_are_copied()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"),
            StoredFiles.Put(this.DataRoot, "KiCad data/KiCad-traces.json", "generated, read by nothing"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal(["Open128.kicad_pcb"], Directory.GetFiles(folder, "*", SearchOption.AllDirectories).Select(Path.GetFileName));
        Assert.Equal(["KiCad data/Open128.kicad_pcb"], BoardKiCadFolder.ListedNames(this.PublishedBoardFile));
    }

    // Asked again with nothing changed - every board shown, every KiCad view - it writes nothing,
    // not even its own record.
    [Fact]
    public void Asking_again_with_nothing_changed_writes_nothing()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        string record = Path.Combine(Path.GetDirectoryName(folder)!, BoardKiCadFolder.RecordFileName);
        DateTime earlier = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(record, earlier);

        BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal(earlier, File.GetLastWriteTimeUtc(record));
    }

    // The copy deletes every file it does not want - never through a folder linked elsewhere,
    // where those would be somebody's own files (code review, 2026-10-25).
    [Fact]
    public void A_copy_folder_that_links_elsewhere_is_neither_written_nor_tidied()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"));

        string theirs = this.thisWorkspace.WriteFile(Path.Combine("Someone else's", "notes.txt"), "keep me");
        string copyFolder = Path.Combine(this.CacheRoot, BoardKiCadFolder.FolderNameFor(this.PublishedBoardFile));
        string link = Path.Combine(copyFolder, DataTreeUsage.KiCadFolderName);
        Directory.CreateDirectory(copyFolder);

        try
        {
            Directory.CreateSymbolicLink(link, Path.GetDirectoryName(theirs)!);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Assert.Skip("This account may not create a folder link (Windows outside Developer Mode).");
        }

        try
        {
            Assert.Equal(string.Empty, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
            Assert.Equal("keep me", File.ReadAllText(theirs));
            Assert.False(File.Exists(Path.Combine(Path.GetDirectoryName(theirs)!, "Open128.kicad_pcb")));
        }
        finally
        {
            TempWorkspace.RemoveLink(link);
        }
    }

    [Fact]
    public void A_listed_file_whose_bytes_are_nowhere_is_left_out()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"),
            StoredFiles.EntryFor("KiCad data/Open128.kicad_sch", "never stored"));

        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Single(Directory.GetFiles(folder, "*", SearchOption.AllDirectories));
    }

    // A draft is the whole board: its file lists the published KiCad files as well as its own,
    // so the board is shown from the draft alone when there is one.
    [Fact]
    public void ForBoard_reads_the_drafts_board_file_when_there_is_one_else_the_published_one()
    {
        const string identity = "Commodore/C128/310378/board.json";
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "published"));

        string published = BoardKiCadFolder.ForBoard(this.DataRoot, this.DraftFolder, identity, this.CacheRoot);
        Assert.Equal("published", Read(published, "Open128.kicad_pcb"));

        StoredFiles.WriteBoard(
            Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName),
            StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_pcb", "drafted"));

        string drafted = BoardKiCadFolder.ForBoard(this.DataRoot, this.DraftFolder, identity, this.CacheRoot);
        Assert.Equal("drafted", Read(drafted, "Open128.kicad_pcb"));
        Assert.NotEqual(published, drafted);

        Assert.Equal(published, BoardKiCadFolder.ForBoard(this.DataRoot, string.Empty, identity, this.CacheRoot));
    }
}
