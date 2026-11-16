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
    // *** BYTES WRITTEN IN ONE CLOCK TICK ARE NEVER TAKEN FOR EACH OTHER (2026-10-25). *** The KiCad
    // loader keeps its parse by path and write time (KiCadProjectLoader). File.Copy keeps the
    // source's time, so two different files written in one clock tick - a contributor's copy and the
    // published one, a few milliseconds apart - gave the copy the stamp it already had, and the
    // loader answered with the old file's parse: "Those KiCad files could not be read." after a
    // Remove that had put a good file back (found by a test that failed in some orders only). Since
    // 2026-11-16 other bytes are another folder (BoardKiCadFolder.FolderNameOf), so another path - and
    // the copy the board file held before goes as it gets the new one, as no other board holds it.
    // ###########################################################################################
    [Fact]
    public void Bytes_stored_in_the_same_clock_tick_are_copied_to_another_path()
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
        string firstFolder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("first bytes", Read(firstFolder, "Open128.kicad_pcb"));

        StoredFiles.WriteBoard(this.PublishedBoardFile, second);
        string secondFolder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.NotEqual(firstFolder, secondFolder);
        Assert.Equal("second, other bytes", Read(secondFolder, "Open128.kicad_pcb"));
        Assert.False(Directory.Exists(firstFolder));
    }

    [Fact]
    public void A_board_with_no_KiCad_files_has_no_folder()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "Schematics/U1.png", "picture"));

        Assert.Equal(string.Empty, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
        Assert.False(Directory.Exists(this.CacheRoot) && Directory.EnumerateFileSystemEntries(this.CacheRoot).Any());
    }

    // ###########################################################################################
    // *** A BOARD AND ITS DRAFT SHARE ONE COPY (2026-11-16, performance walkthrough). *** A draft lists
    // the published KiCad files as they are, so the first change on a board - which makes the draft -
    // must not copy a 45 MB PCB again, nor have the loader parse it again under another path. The same
    // files under the same names are the same folder, whichever board file lists them.
    // ###########################################################################################
    [Fact]
    public void A_board_and_a_draft_listing_the_same_KiCad_files_share_one_copy()
    {
        string draftBoardFile = Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);
        BoardDocument.StoredFileRow pcb = StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb");
        StoredFiles.WriteBoard(this.PublishedBoardFile, pcb);
        StoredFiles.WriteBoard(draftBoardFile, pcb);

        string published = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        string copy = Path.Combine(published, "Open128.kicad_pcb");
        DateTime earlier = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        File.SetLastWriteTimeUtc(copy, earlier);

        Assert.Equal(published, BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot));
        Assert.Equal(earlier, File.GetLastWriteTimeUtc(copy));
    }

    // ###########################################################################################
    // *** A COPY GOES AS SOON AS NO BOARD'S KiCad FILES ARE WHAT IT HOLDS (owner request, 2026-11-16: "a temporary file
    // should be deleted when it is not used anymore" - one nothing had used for 14 days went until then). *** The sweep after
    // launch: the copy of a board shown stays, whatever its age; one no board file holds goes, and so do the copy of a board
    // file gone, the copy of a board whose KiCad files changed since, and a folder an earlier pre-release named by board file
    // (16 hex digits) - never anything not CRT's.
    // ###########################################################################################
    [Fact]
    public void The_sweep_keeps_the_copies_boards_hold_and_takes_away_every_other()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "in use"));
        string held = Path.GetDirectoryName(BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot))!;
        Directory.SetLastWriteTimeUtc(held, DateTime.UtcNow.AddDays(-400));

        // A draft discarded with CRT closed, and another board whose KiCad files a data update changed.
        string goneBoardFile = Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);
        StoredFiles.WriteBoard(goneBoardFile, StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_pcb", "drafted"));
        string ofGone = Path.GetDirectoryName(BoardKiCadFolder.For(goneBoardFile, this.DataRoot, this.CacheRoot))!;
        Directory.Delete(this.DraftFolder, recursive: true);

        string changedBoardFile = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", DataTreeLayout.BoardFileName);
        StoredFiles.WriteBoard(changedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/C64.kicad_pcb", "before the update"));
        string ofChanged = Path.GetDirectoryName(BoardKiCadFolder.For(changedBoardFile, this.DataRoot, this.CacheRoot))!;
        StoredFiles.WriteBoard(changedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/C64.kicad_pcb", "after the update"));

        string unheld = Path.Combine(this.CacheRoot, BoardKiCadFolder.FolderPrefix + "0123456789abcdef");
        Directory.CreateDirectory(Path.Combine(unheld, DataTreeUsage.KiCadFolderName));

        string byBoardFile = Path.Combine(this.CacheRoot, "0123456789abcdef");
        Directory.CreateDirectory(Path.Combine(byBoardFile, DataTreeUsage.KiCadFolderName));

        string notCrts = Path.Combine(this.CacheRoot, "Someone's notes");
        Directory.CreateDirectory(notCrts);

        Assert.Equal(4, BoardKiCadFolder.Sweep(this.CacheRoot, this.DataRoot));

        Assert.True(Directory.Exists(held));
        Assert.False(Directory.Exists(ofGone));
        Assert.False(Directory.Exists(ofChanged));
        Assert.False(Directory.Exists(unheld));
        Assert.False(Directory.Exists(byBoardFile));
        Assert.True(Directory.Exists(notCrts));

        // The board changed is copied anew as it is shown.
        Assert.Equal("after the update", Read(BoardKiCadFolder.For(changedBoardFile, this.DataRoot, this.CacheRoot), "C64.kicad_pcb"));
    }

    [Fact]
    public void The_sweep_of_a_cache_never_made_takes_nothing_away()
    {
        Assert.Equal(0, BoardKiCadFolder.Sweep(this.CacheRoot, this.DataRoot));
        Assert.Equal(0, BoardKiCadFolder.Sweep(string.Empty, this.DataRoot));
    }

    // ###########################################################################################
    // *** A board file whose KiCad files changed lets go of its copy at once - unless another board file holds it (owner
    // request, 2026-11-16). *** A draft lists the published KiCad files as they are, so a board and its draft share one copy:
    // the draft importing its own keeps the published board's copy, which the board still holds. A draft discarded lets go of
    // its own copy (LetGoOf), never the one the board holds too.
    // ###########################################################################################
    [Fact]
    public void A_copy_goes_as_its_board_file_lets_go_of_it_unless_another_holds_it()
    {
        string draftBoardFile = Path.Combine(this.DraftFolder, DataTreeLayout.BoardFileName);
        BoardDocument.StoredFileRow pcb = StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "published");
        StoredFiles.WriteBoard(this.PublishedBoardFile, pcb);
        StoredFiles.WriteBoard(draftBoardFile, pcb);

        string shared = Path.GetDirectoryName(BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot))!;
        Assert.Equal(shared, Path.GetDirectoryName(BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot)));

        // The draft imports its own: the board still holds the shared copy.
        StoredFiles.WriteBoard(draftBoardFile, StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_pcb", "imported"));
        string own = Path.GetDirectoryName(BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot))!;
        Assert.True(Directory.Exists(shared));

        // Imported again: the draft's first copy is nobody's, and goes at once.
        StoredFiles.WriteBoard(draftBoardFile, StoredFiles.Put(this.DraftFolder, "KiCad data/Open128.kicad_pcb", "imported again"));
        string again = Path.GetDirectoryName(BoardKiCadFolder.For(draftBoardFile, this.DataRoot, this.CacheRoot))!;
        Assert.False(Directory.Exists(own));

        // Discarded: its copy goes, the board's stays.
        BoardKiCadFolder.LetGoOf(draftBoardFile, this.CacheRoot);
        Assert.False(Directory.Exists(again));
        Assert.True(Directory.Exists(shared));
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
    public void A_file_no_longer_listed_is_not_in_the_copy_and_a_changed_one_is_copied_again()
    {
        StoredFiles.WriteBoard(
            this.PublishedBoardFile,
            StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_sch", "first"),
            StoredFiles.Put(this.DataRoot, "KiCad data/Pages/old.kicad_sch", "dropped later"));

        BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_sch", "second"));
        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);

        Assert.Equal("second", Read(folder, "Open128.kicad_sch"));
        Assert.False(File.Exists(Path.Combine(folder, "Pages", "old.kicad_sch")));
        Assert.False(Directory.Exists(Path.Combine(folder, "Pages")));
    }

    // A file a copy did not want - one half copied when CRT stopped ("<name>.crt-copying") - goes as the copy is asked for.
    [Fact]
    public void A_file_the_copy_does_not_want_is_taken_out_of_it()
    {
        StoredFiles.WriteBoard(this.PublishedBoardFile, StoredFiles.Put(this.DataRoot, "KiCad data/Open128.kicad_pcb", "pcb"));
        string folder = BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot);
        string halfCopied = Path.Combine(folder, "Open128.kicad_pcb.crt-copying");
        File.WriteAllText(halfCopied, "half");

        Assert.Equal(folder, BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot));
        Assert.False(File.Exists(halfCopied));
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

        // The folder the copy would be made in, found by making it once and taking its files away.
        string copyFolder = Path.GetDirectoryName(BoardKiCadFolder.For(this.PublishedBoardFile, this.DataRoot, this.CacheRoot))!;
        Directory.Delete(copyFolder, recursive: true);
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
