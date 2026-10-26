using Handlers.DataHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for DataManager's orphan/non-used file cleanup - the only code path in the application
// that DELETES files from the user's data root. The rule these tests pin down is that the
// cleanup must FAIL CLOSED: it may only delete a file when the referenced-file map is complete,
// because every gap in that map turns into a deletion.
//
// The map has two sources, and losing either one silently must abort the cleanup:
//   1. The online manifest snapshot - the authority on which files the server provides. With no
//      snapshot (offline launch, DNS failure, server down) "orphan" cannot be determined at all.
//   2. The walk of the lists - in CRT 4.0 (2026-10-24) the main file, the board files it lists,
//      and the stored files each board file names. A list that fails to read references nothing,
//      so a corrupt or locked file must abort rather than count as "references nothing".
//
// (CRT 3.x also walked a "_UserContribution" workbook naming the user's own boards in the data
// folder. 4.0 has none: a board of the user's own is a draft, outside the data folder - which the
// draft tests at the bottom hold.)
//
// The tests call the internal DeleteOrphanAndUnusedFilesAsync(manifestSnapshot) seam directly.
// The public overload only adds the two UserSettings gates on top; going through it would force
// this class to mutate UserSettings static state as well, and a test class can only join one
// xUnit collection ("DataManager" here), so the settings statics could race with the parallel
// "UserSettings" collection.
//
// Static state again, so this file shares the sequential "DataManager" collection.
[Collection("DataManager")]
public sealed class DataManagerOrphanCleanupTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // Leave the static lists empty so nothing later mistakes test data for real data.
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");

        // The draft-safety tests below point DraftManager's static root at a temp folder. Reset it
        // to the unloaded state so it never outlives this workspace - the same rule
        // DraftManagerTests and TabDraftsTests follow. Collections never run in parallel
        // (xunit.runner.json), so the "Drafts" collection cannot be running concurrently, but a
        // stale root left pointing at a deleted folder would still confuse whatever ran next.
        DraftManager.LoadFrom(string.Empty);

        this.thisWorkspace.Dispose();
        this.thisBesideData.Dispose();
    }

    // ------------------------------------------------------------------ fixture helpers

    private const string MainName = MainFiles.FileName;

    private const string BoardId = "Commodore/C64/250407";

    private static readonly string BoardFile = DataTreeLayout.BoardFilePath(BoardId);

    // The main file listing the given boards, in `dataRoot` (the workspace by default).
    private void WriteMain(string? dataRoot = null, params string[] boardIds) =>
        MainFiles.Write(
            dataRoot ?? this.thisWorkspace.Root,
            boardIds.Select((id, index) => new MainFiles.Board(id, "C64", $"Board {index + 1}")).ToArray());

    // A board file naming a schematic and a manual, their bytes in the tree's store - the stored
    // paths, data-root relative.
    private (string Schematic, string Manual) WriteBoard(string? dataRoot = null, string boardId = BoardId)
    {
        string root = dataRoot ?? this.thisWorkspace.Root;
        BoardDocument.StoredFileRow schematic = StoredFiles.Put(root, "schematic.png", "png of " + boardId);
        BoardDocument.StoredFileRow manual = StoredFiles.Put(root, "manual.pdf", "pdf of " + boardId);

        StoredFiles.WriteBoard(
            Path.Combine(root, DataTreeLayout.BoardFilePath(boardId).Replace('/', Path.DirectorySeparatorChar)),
            new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "schematic.png" }],
                BoardLocalFiles = [new BoardLocalFileEntry { Category = "Schematics", Name = "Service manual", File = "manual.pdf" }],
            },
            schematic,
            manual);

        return (DataTreeLayout.StoredFilePath(schematic.Sha256, schematic.Name), DataTreeLayout.StoredFilePath(manual.Sha256, manual.Name));
    }

    private static DataFileEntry ManifestEntry(string file) => new()
    {
        File = file,
        Checksum = new string('a', 64),
        Url = "https://example.org/" + file
    };

    private bool FileExists(string relativePath)
        => File.Exists(Path.Combine(this.thisWorkspace.Root, relativePath.Replace('/', Path.DirectorySeparatorChar)));

    // ------------------------------------------------- fail closed: no manifest snapshot

    [Fact]
    public async Task Cleanup_without_a_manifest_snapshot_deletes_nothing()
    {
        // The offline-launch scenario: the manifest fetch failed, so the snapshot is null. The
        // map would be missing every server-provided file that no board names (READMEs, other
        // main files' boards, ...), and all of them would be deleted.
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        this.thisWorkspace.WriteFile("only-the-manifest-knows-me.txt", "server-provided file");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(manifestSnapshot: null);

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("only-the-manifest-knows-me.txt"),
            "a file was deleted although no manifest snapshot existed to say whether it is an orphan");
    }

    [Fact]
    public async Task Cleanup_with_an_empty_manifest_snapshot_deletes_nothing()
    {
        // An empty snapshot carries the same problem as a missing one: it cannot vouch for
        // anything, so nothing may be deleted on its authority.
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        this.thisWorkspace.WriteFile("only-the-manifest-knows-me.txt", "server-provided file");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>());

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("only-the-manifest-knows-me.txt"));
    }

    // ------------------------------------------- fail closed: incomplete mapping

    [Fact]
    public async Task An_unreadable_board_file_aborts_the_cleanup()
    {
        // A board file that exists but cannot be parsed names an unknown set of files. Treating
        // it as "names nothing" would delete every stored file of that board, so the whole
        // cleanup must abort instead.
        this.WriteMain(boardIds: BoardId);
        this.thisWorkspace.WriteFile(BoardFile, "this is not a board file");
        this.thisWorkspace.WriteFile("Files/ab/stored.png", "bytes of the unreadable board");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName) });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("Files/ab/stored.png"),
            "an unreadable board file must abort the cleanup, not orphan that board's files");
    }

    [Fact]
    public async Task A_missing_resolved_main_file_aborts_the_cleanup()
    {
        // With the resolved main file absent (e.g. a fresh data root where the download failed),
        // the board list is unknown, so nothing can be classified as an orphan.
        this.thisWorkspace.WriteFile("Files/ab/stored.png", "board bytes");

        DataManager.LoadFrom(this.thisWorkspace.Root, "ghost.json");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry("ghost.json") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("Files/ab/stored.png"));
    }

    [Fact]
    public async Task An_unreadable_main_file_aborts_the_cleanup()
    {
        this.WriteMain(boardIds: BoardId);
        (string schematic, _) = this.WriteBoard();
        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        // Damaged after the load - the cleanup reads it again to build its map.
        this.thisWorkspace.WriteFile(MainName, "{ not json");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName) });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists(schematic));
    }

    // -------------------------------------------------- the intended behaviour still works

    [Fact]
    public async Task An_orphan_file_is_deleted_and_its_empty_folder_removed()
    {
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        this.thisWorkspace.WriteFile("Stale/orphan.bin", "in no manifest, referenced by nothing");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName), ManifestEntry(BoardFile) });

        Assert.Equal(1, deleted);
        Assert.False(this.FileExists("Stale/orphan.bin"));
        Assert.False(Directory.Exists(Path.Combine(this.thisWorkspace.Root, "Stale")),
            "the emptied folder should be removed too");
    }

    // ------------------------------------------------------------ oscilloscope captures
    //
    // Code review, 2026-10-13: a capture was saved, by default, straight into the data folder - and
    // a capture is a file no manifest lists, so the cleanup deleted every one of them.

    private readonly TempWorkspace thisBesideData = new();

    private string CaptureFolderBesideData => this.thisBesideData.Path_("Oscilloscope images");

    private List<DataFileEntry> MinimalManifest() =>
        new() { ManifestEntry(MainName), ManifestEntry(BoardFile) };

    private void WriteMinimalBoard()
    {
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);
    }

    [Fact]
    public async Task Captures_the_old_default_left_in_the_data_folder_are_moved_out_before_the_cleanup()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("U1_5.png", "a capture");
        this.thisWorkspace.WriteFile("U14_Pin 3_PAL.png", "another capture");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            this.MinimalManifest(),
            new OscilloscopeCaptureRelocation.Plan(this.CaptureFolderBesideData, this.CaptureFolderBesideData));

        Assert.Equal(0, deleted);
        Assert.False(this.FileExists("U1_5.png"));
        Assert.Equal("a capture", File.ReadAllText(Path.Combine(this.CaptureFolderBesideData, "U1_5.png")));
        Assert.Equal("another capture", File.ReadAllText(Path.Combine(this.CaptureFolderBesideData, "U14_Pin 3_PAL.png")));
    }

    // A capture whose name is already taken in the capture folder is not ours to overwrite - so it
    // stays in the data folder, and the cleanup leaves it alone.
    [Fact]
    public async Task A_capture_that_cannot_be_moved_stays_and_the_cleanup_keeps_away_from_it()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("U1_5.png", "the older capture");
        Directory.CreateDirectory(this.CaptureFolderBesideData);
        File.WriteAllText(Path.Combine(this.CaptureFolderBesideData, "U1_5.png"), "a newer capture");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            this.MinimalManifest(),
            new OscilloscopeCaptureRelocation.Plan(this.CaptureFolderBesideData, this.CaptureFolderBesideData));

        Assert.Equal(0, deleted);
        Assert.Equal("the older capture", this.thisWorkspace.ReadFile("U1_5.png"));
        Assert.Equal("a newer capture", File.ReadAllText(Path.Combine(this.CaptureFolderBesideData, "U1_5.png")));
    }

    // An image in the data folder's top level that does not look like a capture could still be the
    // user's: not moved, and not deleted.
    [Fact]
    public async Task An_image_in_doubt_in_the_data_folder_is_neither_moved_nor_deleted()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("my board photo.jpg", "a photo");
        this.thisWorkspace.WriteFile("Stale/orphan.bin", "a real orphan");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            this.MinimalManifest(),
            new OscilloscopeCaptureRelocation.Plan(this.CaptureFolderBesideData, this.CaptureFolderBesideData));

        Assert.Equal(1, deleted);
        Assert.True(this.FileExists("my board photo.jpg"));
        Assert.False(this.FileExists("Stale/orphan.bin"));
        Assert.False(File.Exists(Path.Combine(this.CaptureFolderBesideData, "my board photo.jpg")));
    }

    // A capture folder the user chose INSIDE the data folder: nothing in it is deleted, nor the
    // folder itself when it is empty - and its captures are not moved anywhere.
    [Fact]
    public async Task Nothing_inside_a_capture_folder_within_the_data_folder_is_deleted()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("Scope shots/U1_5.png", "a capture");
        this.thisWorkspace.WriteFile("Scope shots/notes.txt", "the user's notes");
        Directory.CreateDirectory(this.thisWorkspace.Path_("Scope shots", "empty for now"));
        this.thisWorkspace.WriteFile("Stale/orphan.bin", "a real orphan");

        string captureFolder = this.thisWorkspace.Path_("Scope shots");
        var plan = new OscilloscopeCaptureRelocation.Plan(captureFolder, MoveTarget: null);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(this.MinimalManifest(), plan);

        Assert.Equal(1, deleted);
        Assert.True(this.FileExists("Scope shots/U1_5.png"));
        Assert.True(this.FileExists("Scope shots/notes.txt"));
        Assert.True(Directory.Exists(this.thisWorkspace.Path_("Scope shots", "empty for now")));
        Assert.False(this.FileExists("Stale/orphan.bin"));
    }

    // The capture folder IS the data folder (picked there by hand): every file in it could be a
    // capture, so the cleanup deletes nothing at all.
    [Fact]
    public async Task A_capture_folder_that_is_the_data_folder_stops_the_cleanup()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("U1_5.png", "a capture");
        this.thisWorkspace.WriteFile("Stale/orphan.bin", "would be an orphan");

        var plan = OscilloscopeCaptureRelocation.Plan.ForFolder(this.thisWorkspace.Root, this.thisWorkspace.Root);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(this.MinimalManifest(), plan);

        Assert.Null(plan.MoveTarget);
        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("U1_5.png"));
        Assert.True(this.FileExists("Stale/orphan.bin"));
    }

    // A file the manifest lists is data, whatever its name looks like.
    [Fact]
    public async Task A_capture_shaped_file_the_manifest_lists_is_not_moved()
    {
        this.WriteMinimalBoard();
        this.thisWorkspace.WriteFile("Logo_small.png", "published data");

        var manifest = this.MinimalManifest();
        manifest.Add(ManifestEntry("Logo_small.png"));

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            manifest,
            new OscilloscopeCaptureRelocation.Plan(this.CaptureFolderBesideData, this.CaptureFolderBesideData));

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("Logo_small.png"));
        Assert.False(File.Exists(Path.Combine(this.CaptureFolderBesideData, "Logo_small.png")));
    }

    [Fact]
    public async Task A_file_listed_in_the_manifest_survives_even_when_no_board_names_it()
    {
        // README-style files are provided by the server but named by no board; the manifest
        // snapshot is what keeps them alive.
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        this.thisWorkspace.WriteFile("!README.txt", "about this data");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName), ManifestEntry("!README.txt") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("!README.txt"));
    }

    // The stored bytes a board file names are kept by the walk alone - the manifest here lists
    // nothing but the main file, as a manifest fetched before a publish would.
    [Fact]
    public async Task The_stored_files_a_board_names_survive_the_cleanup()
    {
        this.WriteMain(boardIds: BoardId);
        (string schematic, string manual) = this.WriteBoard();

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName) });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists(BoardFile));
        Assert.True(this.FileExists(schematic));
        Assert.True(this.FileExists(manual));
    }

    // Bytes no board names any more - a picture replaced by a publish - go once the manifest no
    // longer lists them either.
    [Fact]
    public async Task Stored_bytes_no_board_names_any_more_are_deleted()
    {
        this.WriteMain(boardIds: BoardId);
        this.WriteBoard();
        BoardDocument.StoredFileRow replaced = StoredFiles.Put(this.thisWorkspace.Root, "schematic.png", "the old scan");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainName), ManifestEntry(BoardFile) });

        Assert.Equal(1, deleted);
        Assert.False(File.Exists(StoredFiles.PathOf(this.thisWorkspace.Root, replaced)));
    }

    // ------------------------------------------------- A sync never destroys a draft

    // THE first line of Phase 2's own definition of done in NewContributeStrategy.md: "A sync never
    // destroys a draft (test this explicitly, with a real sync)." Nothing pinned it until now.
    //
    // The property is structural rather than defensive, which is exactly why it is worth a test:
    // cleanup only ever enumerates inside the DATA root (and re-checks containment per file), and
    // "Drafts/" is a SIBLING of the data folder in AppData, never inside it. So no guard in the
    // cleanup code mentions drafts at all - the safety comes entirely from that separation, and a
    // future change that moved drafts under the data root, or widened the enumeration, would
    // silently start deleting a contributor's unpublished work with nothing to restore it from.
    // These tests fail the moment that separation stops holding.
    //
    // Deliberately run with a manifest snapshot that vouches for everything and NO board naming
    // the draft, which is the harshest case: every drafted file is, by the cleanup's own
    // definition, unreferenced - and must still survive because it is out of scope entirely.
    private string SetUpDraftBesideDataRoot(string excelDataFile)
    {
        // Inside the workspace so it is cleaned up with it, but deliberately NOT inside the data
        // root that DataManager is pointed at (the workspace's "Data" subfolder below) - that
        // sibling relationship is exactly the real AppData layout, and is the whole reason the
        // cleanup cannot reach a draft.
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

        string draftFolder = DraftManager.GetBoardFolder(excelDataFile);

        // Written the way DraftSeeder would - a board file plus the marker that makes the folder a
        // draft - so what the cleanup is proven not to touch is the real thing.
        var board = new BoardData { RevisionDate = "2026-01-15" };
        board.Components.Add(new ComponentEntry
        {
            BoardLabel = "U8",
            FriendlyName = "PLA (my correction)",
        });

        CachedWorkbooks.Write(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile), board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { BoardKey = excelDataFile, BaseRevision = "2026-01-15" });

        // A DRAFTED ATTACHMENT - bytes that exist in no other copy. The cleanup walks the data
        // folder for files the manifest does not name; reaching into Drafts/ would destroy this
        // one, and the contributor would have nothing to restore it from.
        string attachment = DraftFileResolver.BuildDraftFileDestination(draftFolder, "docs/datasheet.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(attachment)!);
        File.WriteAllText(attachment, "a datasheet the contributor attached");

        return draftFolder;
    }

    // A data root that is a SUBFOLDER of the workspace, so "Drafts/" can sit beside it rather than
    // inside it. The other tests in this class point DataManager at the workspace root itself,
    // which would put the drafts folder within the very tree the cleanup walks - inverting the
    // property under test.
    private string SetUpDataRootWithBoard()
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        Directory.CreateDirectory(dataRoot);

        this.WriteMain(dataRoot, BoardId);
        this.WriteBoard(dataRoot);

        return dataRoot;
    }

    [Fact]
    public async Task The_cleanup_never_touches_a_draft_or_its_attachments()
    {
        string dataRoot = this.SetUpDataRootWithBoard();
        string draftFolder = this.SetUpDraftBesideDataRoot(BoardFile);

        // An unreferenced file inside the data root, to prove the cleanup actually runs here.
        File.WriteAllText(Path.Combine(dataRoot, "stray-orphan.txt"), "not referenced by anything");

        DataManager.LoadFrom(dataRoot, MainName);

        await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>
        {
            ManifestEntry(MainName),
            ManifestEntry(BoardFile),
        });

        // Sanity: the cleanup really is ACTIVE in this setup rather than short-circuiting, so the
        // survival assertions below mean something. An unreferenced file inside the DATA root is
        // deleted by the same call that leaves the draft alone.
        Assert.False(
            File.Exists(Path.Combine(dataRoot, "stray-orphan.txt")),
            "the cleanup deleted nothing at all here, so this test could not prove a draft survives it");

        Assert.True(
            File.Exists(DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, BoardFile)),
            "the cleanup deleted a draft's board file - a contributor's unpublished work, which nothing can restore");

        Assert.True(
            File.Exists(Path.Combine(draftFolder, DraftFolderLayout.DraftMarkerFileName)),
            "the cleanup deleted a draft's marker - a contributor's unpublished work, which nothing can restore");

        Assert.True(
            File.Exists(DraftFileResolver.BuildDraftFileDestination(draftFolder, "docs/datasheet.pdf")),
            "the cleanup deleted a drafted attachment, which exists in no other copy");
    }

    // The draft's rows must still read back intact afterwards, not merely have a file left behind:
    // a truncated or rewritten board file would pass a File.Exists check while having lost the work.
    [Fact]
    public async Task A_draft_still_reads_back_intact_after_a_cleanup()
    {
        string dataRoot = this.SetUpDataRootWithBoard();
        this.SetUpDraftBesideDataRoot(BoardFile);

        DataManager.LoadFrom(dataRoot, MainName);

        await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>
        {
            ManifestEntry(MainName),
            ManifestEntry(BoardFile),
        });

        // The contributor's own work, still there. This is the guarantee the whole test exists
        // for: the sync owns the data folder, and a cleanup that reached into Drafts/ would destroy
        // unpublished work with no way to recover it.
        BoardData? draftBoard = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, BoardFile);

        Assert.NotNull(draftBoard);
        Assert.Equal("PLA (my correction)", Assert.Single(draftBoard!.Components).FriendlyName);

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, BoardFile));

        Assert.NotNull(marker);
        Assert.Equal("2026-01-15", marker!.BaseRevision);
    }
}
