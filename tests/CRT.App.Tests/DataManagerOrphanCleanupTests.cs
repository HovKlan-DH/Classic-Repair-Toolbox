using Handlers.DataHandling;
using Handlers.OnlineHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// Tests for DataManager's orphan/non-used file cleanup - the only code path in the application
// that DELETES files from the user's data root. The rule these tests pin down is that the
// cleanup must FAIL CLOSED: it may only delete a file when the referenced-file map is complete,
// because every gap in that map turns into a deletion.
//
// The map has two sources, and losing either one silently must abort the cleanup:
//   1. The online manifest snapshot - the authority on which files the server provides. With no
//      snapshot (offline launch, DNS failure, server down) "orphan" cannot be determined at all.
//   2. The workbook walk - main workbook, user-contribution sidecar, and board workbooks. A
//      workbook that fails to read references nothing, so a corrupt or locked file must abort
//      rather than count as "references nothing". This is the only protection user-contributed
//      local-only files have - no re-sync can restore those.
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

    static DataManagerOrphanCleanupTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public void Dispose()
    {
        // Leave the static lists empty so nothing later mistakes test data for real data.
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");

        // The draft-safety tests below point DraftManager's static root at a temp folder. Reset it
        // to the unloaded state so it never outlives this workspace - the same rule
        // DraftManagerTests and TabDraftsTests follow. Collections never run in parallel
        // (xunit.runner.json), so the "Drafts" collection cannot be running concurrently, but a
        // stale root left pointing at a deleted folder would still confuse whatever ran next.
        DraftManager.LoadFrom(string.Empty);

        this.thisWorkspace.Dispose();
    }

    // ------------------------------------------------------------------ fixture helpers

    private const string MainExcelName = "master.xlsx";

    private static readonly string[] HardwareHeaders =
    {
        "Hardware name in drop-down", "Board name in drop-down",
        "Excel data file", "Hardware notes in \"Overview\" tab"
    };

    /// <summary>Writes a master workbook listing the given board workbook paths (data-root relative).</summary>
    private void WriteMasterWorkbook(string fileName, params string[] boardExcelFiles)
    {
        string?[][] rows = boardExcelFiles
            .Select((file, index) => new string?[] { "C64", $"Board {index + 1}", file, "" })
            .ToArray();

        new BoardWorkbookBuilder()
            .Sheet("Hardware & Board", HardwareHeaders, rows)
            .SaveTo(this.thisWorkspace.Path_(fileName));
    }

    /// <summary>Writes a board workbook whose sheets reference the given data-root-relative files.</summary>
    private void WriteBoardWorkbook(string relativePath, string schematicImage, string boardLocalFile)
    {
        new BoardWorkbookBuilder()
            .Sheet("Board schematics", BoardWorkbookBuilder.SchematicsHeaders,
                new[] { "uuid-1", "Sheet 1", "", schematicImage, "", "", "", "", "" })
            .Sheet("Board local files", BoardWorkbookBuilder.BoardLocalFilesHeaders,
                new[] { "Schematics", "Service manual", boardLocalFile })
            .SaveTo(this.thisWorkspace.Path_(relativePath.Replace('/', Path.DirectorySeparatorChar)));
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
        // map would be missing every server-provided file that no workbook references (READMEs,
        // other main-workbook versions' boards, ...), and all of them would be deleted.
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("only-the-manifest-knows-me.txt", "server-provided file");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

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
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("only-the-manifest-knows-me.txt", "server-provided file");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>());

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("only-the-manifest-knows-me.txt"));
    }

    // ------------------------------------------- fail closed: incomplete workbook mapping

    [Fact]
    public async Task An_unreadable_board_workbook_aborts_the_cleanup()
    {
        // A board workbook that exists but cannot be parsed references an unknown set of files.
        // Treating it as "references nothing" would delete every local-only asset of that board,
        // so the whole cleanup must abort instead.
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.thisWorkspace.WriteFile("C64/Data.xlsx", "this is not an Excel file");
        this.thisWorkspace.WriteFile("C64/photo.png", "asset of the unreadable board");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName) });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("C64/photo.png"),
            "an unreadable board workbook must abort the cleanup, not orphan that board's files");
    }

    [Fact]
    public async Task A_missing_resolved_main_workbook_aborts_the_cleanup()
    {
        // With the resolved main workbook absent (e.g. a fresh data root where the download
        // failed), the board list is unknown, so nothing can be classified as an orphan.
        this.thisWorkspace.WriteFile("C64/photo.png", "board asset");

        DataManager.LoadFrom(this.thisWorkspace.Root, "ghost.xlsx");

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry("ghost.xlsx") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("C64/photo.png"));
    }

    [Fact]
    public async Task An_unreadable_user_contribution_sidecar_aborts_the_cleanup()
    {
        // The contribution sidecar names the user's OWN boards - files that exist nowhere but on
        // this machine. If the sidecar is corrupt while the main workbook still reads fine, the
        // contribution's files would be the ones deleted, and no re-sync could bring them back.
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("master_UserContribution.xlsx", "this is not an Excel file");
        this.thisWorkspace.WriteFile("Contrib/photo.png", "user-contributed, local-only");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName), ManifestEntry("C64/Data.xlsx") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("Contrib/photo.png"),
            "a corrupt contribution sidecar must abort the cleanup, not forfeit the user's own data");
    }

    // -------------------------------------------------- the intended behaviour still works

    [Fact]
    public async Task An_orphan_file_is_deleted_and_its_empty_folder_removed()
    {
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("Stale/orphan.bin", "in no manifest, referenced by nothing");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName), ManifestEntry("C64/Data.xlsx") });

        Assert.Equal(1, deleted);
        Assert.False(this.FileExists("Stale/orphan.bin"));
        Assert.False(Directory.Exists(Path.Combine(this.thisWorkspace.Root, "Stale")),
            "the emptied folder should be removed too");
    }

    [Fact]
    public async Task A_file_listed_in_the_manifest_survives_even_when_no_workbook_references_it()
    {
        // README-style files are provided by the server but referenced by no workbook sheet;
        // the manifest snapshot is what keeps them alive.
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("!README.txt", "about this data");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName), ManifestEntry("!README.txt") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("!README.txt"));
    }

    [Fact]
    public async Task Files_referenced_by_a_board_workbook_survive_the_cleanup()
    {
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.thisWorkspace.WriteFile("C64/schematic.png", "png");
        this.thisWorkspace.WriteFile("C64/manual.pdf", "pdf");
        // A file physically under the board's "KiCad data" folder is kept by the folder walk
        // even though no workbook sheet names it.
        this.thisWorkspace.WriteFile("C64/KiCad data/board.kicad_pcb", "kicad");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName) });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("C64/schematic.png"));
        Assert.True(this.FileExists("C64/manual.pdf"));
        Assert.True(this.FileExists("C64/KiCad data/board.kicad_pcb"));
    }

    [Fact]
    public async Task A_user_contribution_sidecars_boards_and_assets_survive_the_cleanup()
    {
        // Contribution files are local-only, so they are never in the manifest; the sidecar
        // workbook walk is their only protection.
        this.WriteMasterWorkbook(MainExcelName, "C64/Data.xlsx");
        this.WriteBoardWorkbook("C64/Data.xlsx", "C64/schematic.png", "C64/manual.pdf");
        this.WriteMasterWorkbook("master_UserContribution.xlsx", "Contrib/Data.xlsx");
        this.WriteBoardWorkbook("Contrib/Data.xlsx", "Contrib/schematic.png", "Contrib/manual.pdf");
        this.thisWorkspace.WriteFile("Contrib/schematic.png", "png");
        this.thisWorkspace.WriteFile("Contrib/manual.pdf", "pdf");

        DataManager.LoadFrom(this.thisWorkspace.Root, MainExcelName);

        int deleted = await DataManager.DeleteOrphanAndUnusedFilesAsync(
            new List<DataFileEntry> { ManifestEntry(MainExcelName), ManifestEntry("C64/Data.xlsx") });

        Assert.Equal(0, deleted);
        Assert.True(this.FileExists("master_UserContribution.xlsx"));
        Assert.True(this.FileExists("Contrib/Data.xlsx"));
        Assert.True(this.FileExists("Contrib/schematic.png"));
        Assert.True(this.FileExists("Contrib/manual.pdf"));
    }

    // ------------------------------------------------- A sync never destroys a draft

    // THE first line of Phase 2's own definition of done in NewContributeStrategy.md: "A sync never
    // destroys a draft (test this explicitly, with a real sync)." Nothing pinned it until now.
    //
    // The property is structural rather than defensive, which is exactly why it is worth a test:
    // cleanup only ever enumerates inside the DATA root (and re-checks containment per file), and
    // "Drafts/" is a SIBLING of "Data/" in AppData, never inside it. So no guard in the cleanup
    // code mentions drafts at all - the safety comes entirely from that separation, and a future
    // change that moved drafts under the data root, or widened the enumeration, would silently
    // start deleting a contributor's unpublished work with nothing to restore it from. These tests
    // fail the moment that separation stops holding.
    //
    // Deliberately run with a manifest snapshot that vouches for everything and NO workbook
    // referencing the draft, which is the harshest case: every drafted file is, by the cleanup's
    // own definition, unreferenced - and must still survive because it is out of scope entirely.
    private string SetUpDraftBesideDataRoot(string excelDataFile)
    {
        // Inside the workspace so it is cleaned up with it, but deliberately NOT inside the data
        // root that DataManager is pointed at (the workspace's "Data" subfolder below) - that
        // sibling relationship is exactly the real AppData layout, and is the whole reason the
        // cleanup cannot reach a draft.
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));

        string draftFolder = DraftManager.GetSystemFolder(excelDataFile);

        // *** A DRAFT IS A BOARD FOLDER SINCE PHASE 6 (2026-09-23). *** Written the way
        // DraftSeeder would - a workbook plus the marker that makes the folder a draft - so what
        // the cleanup is proven not to touch is the real thing rather than a stand-in.
        string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);

        var board = new BoardData { RevisionDate = "2026-01-15" };
        board.Components.Add(new ComponentEntry
        {
            BoardLabel = "U8",
            FriendlyName = "PLA (my correction)",
        });

        BoardWorkbookWriter.Write(workbook, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, excelDataFile),
            new DraftMarker { SystemKey = excelDataFile, BaseRevision = "2026-01-15" });

        // A DRAFTED ATTACHMENT - bytes that exist in no other copy. The cleanup walks Data/ for
        // files the manifest does not name; reaching into Drafts/ would destroy this one, and the
        // contributor would have nothing to restore it from.
        string attachment = DraftFileResolver.BuildDraftFileDestination(draftFolder, "docs/datasheet.pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(attachment)!);
        File.WriteAllText(attachment, "a datasheet the contributor attached");

        return draftFolder;
    }

    // A data root that is a SUBFOLDER of the workspace, so "Drafts/" can sit beside it rather than
    // inside it. The other tests in this class point DataManager at the workspace root itself,
    // which would put the drafts folder within the very tree the cleanup walks - inverting the
    // property under test.
    private string SetUpDataRootWithBoard(string excelDataFile)
    {
        string dataRoot = this.thisWorkspace.Path_("Data");
        Directory.CreateDirectory(dataRoot);

        string?[][] rows = { new string?[] { "C64", "Board 1", excelDataFile, "" } };
        new BoardWorkbookBuilder()
            .Sheet("Hardware & Board", HardwareHeaders, rows)
            .SaveTo(Path.Combine(dataRoot, MainExcelName));

        string boardPath = Path.Combine(dataRoot, excelDataFile.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(boardPath)!);

        new BoardWorkbookBuilder()
            .Sheet("Board schematics", BoardWorkbookBuilder.SchematicsHeaders,
                new[] { "uuid-1", "Sheet 1", "", "C64/schematic.png", "", "", "", "", "" })
            .Sheet("Board local files", BoardWorkbookBuilder.BoardLocalFilesHeaders,
                new[] { "Schematics", "Service manual", "C64/manual.pdf" })
            .SaveTo(boardPath);

        return dataRoot;
    }

    [Fact]
    public async Task The_cleanup_never_touches_a_draft_or_its_attachments()
    {
        const string ExcelDataFile = "C64/Data.xlsx";

        string dataRoot = this.SetUpDataRootWithBoard(ExcelDataFile);
        string draftFolder = this.SetUpDraftBesideDataRoot(ExcelDataFile);

        // An unreferenced file inside the data root, to prove the cleanup actually runs here.
        File.WriteAllText(Path.Combine(dataRoot, "stray-orphan.txt"), "not referenced by anything");

        DataManager.LoadFrom(dataRoot, MainExcelName);

        await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>
        {
            ManifestEntry(MainExcelName),
            ManifestEntry(ExcelDataFile),
        });

        // Sanity: the cleanup really is ACTIVE in this setup rather than short-circuiting, so the
        // survival assertions below mean something. An unreferenced file inside the DATA root is
        // deleted by the same call that leaves the draft alone.
        Assert.False(
            File.Exists(Path.Combine(dataRoot, "stray-orphan.txt")),
            "the cleanup deleted nothing at all here, so this test could not prove a draft survives it");

        Assert.True(
            File.Exists(Path.Combine(draftFolder, DraftFolderLayout.DraftMarkerFileName)),
            "the cleanup deleted a draft's marker - a contributor's unpublished work, which nothing can restore");

        Assert.True(
            File.Exists(DraftFileResolver.BuildDraftFileDestination(draftFolder, "docs/datasheet.pdf")),
            "the cleanup deleted a drafted attachment, which exists in no other copy");
    }

    // The draft's rows must still read back intact afterwards, not merely have a file left behind:
    // a truncated or rewritten draft.json would pass a File.Exists check while having lost the work.
    [Fact]
    public async Task A_draft_still_reads_back_intact_after_a_cleanup()
    {
        const string ExcelDataFile = "C64/Data.xlsx";

        string dataRoot = this.SetUpDataRootWithBoard(ExcelDataFile);
        this.SetUpDraftBesideDataRoot(ExcelDataFile);

        DataManager.LoadFrom(dataRoot, MainExcelName);

        await DataManager.DeleteOrphanAndUnusedFilesAsync(new List<DataFileEntry>
        {
            ManifestEntry(MainExcelName),
            ManifestEntry(ExcelDataFile),
        });

        // The contributor's own work, still there. This is the guarantee the whole test exists
        // for: the sync owns Data/, and a cleanup that reached into Drafts/ would destroy
        // unpublished work with no way to recover it.
        BoardData? draftBoard = DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, ExcelDataFile);

        Assert.NotNull(draftBoard);
        Assert.Equal("PLA (my correction)", Assert.Single(draftBoard!.Components).FriendlyName);

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, ExcelDataFile));

        Assert.NotNull(marker);
        Assert.Equal("2026-01-15", marker!.BaseRevision);
    }
}
