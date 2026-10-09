using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// How DataManager.LoadBoardDataAsync resolves a board that has a local draft.
//
// *** THE MECHANISM CHANGED IN PHASE 6; THE BEHAVIOURS DID NOT. *** This file used to set up
// draft.json row deltas and prove the load OVERLAID them onto the published board. A draft is
// now a real board folder with its own workbook, so the load CHOOSES a file instead of merging
// one (DraftBoardSource) - and every test below was rewritten to drive the new mechanism while
// asserting the same thing it always did:
//
//   - a draft changes what the board reads as;
//   - another board's draft never bleeds in;
//   - "view boards as officially published" shows the published rows but still REPORTS the draft;
//   - a board that exists only as a draft loads even with no published file.
//
// The class name is kept even though "overlay" is now the wrong word, because it is the file
// anyone looking for this wiring will search for. The header says what actually happens.
//
// Rewritten rather than deleted, deliberately: deleting the tests along with the mechanism would
// have removed the only proof that the app-level wiring still behaves as promised.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataManagerDraftOverlayTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // Unique per TEST INSTANCE (xUnit constructs a fresh instance per [Fact]), not a shared
    // constant - BoardDataReader's own load cache is a process-wide static keyed by this exact
    // string, and this class is not in the "BoardData" collection that clears it between tests. A
    // shared ExcelDataFile let an earlier test's cached "PLA" survive into a later test that
    // rewrote the Excel file, which then read back the stale value - a test ordering bug, not a
    // product one. A fresh path per test means every test is a guaranteed cache miss.
    private readonly string thisExcelDataFile =
        $"Commodore/C64/{Guid.NewGuid():N}/Data C64 test v1.0.0.xlsx";

    static DataManagerDraftOverlayTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public DataManagerDraftOverlayTests()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        DraftManager.LoadFrom(Path.Combine(this.thisWorkspace.Root, "Drafts"));

        // Redirected before any test here can write ViewOfficialPublishedOnly. The file is written
        // BEFORE LoadFrom deliberately - UserSettings.LoadFrom's "file does not exist" branch
        // returns WITHOUT resetting the static _data, so pointing it at a fresh path would leave
        // the PREVIOUS test's settings in memory. Writing "{}" first forces the
        // deserialize-and-replace path, which is the only one that actually resets it.
        string settingsPath = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(settingsPath, "{}");
        UserSettings.LoadFrom(settingsPath);
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        this.thisWorkspace.Dispose();
    }

    private HardwareBoardEntry BoardEntry => new() { ExcelDataFile = this.thisExcelDataFile };

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private void WriteBoardExcel()
    {
        string path = Path.Combine(
            this.thisWorkspace.Root,
            this.thisExcelDataFile.Replace('/', Path.DirectorySeparatorChar));

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookBuilder.WriteCompleteBoard(path);
    }

    // ###########################################################################################
    // Writes a DRAFT of this board: a real board workbook in the drafts tree, plus the marker
    // that makes the folder a draft.
    //
    // This is what replaced DraftDataStore.Save in every test below. Note it writes a COMPLETE
    // board rather than just the changed rows - that is the seeding model, and it is what makes an
    // absent row mean "deleted" rather than "not copied yet".
    // ###########################################################################################
    private void WriteDraft(BoardData board, NewBoardRegistration? registration = null)
    {
        string workbook = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, this.thisExcelDataFile);
        Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);

        CachedWorkbooks.Write(workbook, board);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, this.thisExcelDataFile),
            new DraftMarker
            {
                BoardKey = this.thisExcelDataFile,
                BaseRevision = registration is null ? "2026-09-01" : string.Empty,
                NewBoard = registration,
            });
    }

    // The published board as BoardWorkbookBuilder writes it, with one component's friendly name
    // changed - the "drafted correction" every test below is built around.
    private static BoardData BoardWithComponent(string boardLabel, string friendlyName) => new()
    {
        RevisionDate = "2026-09-01",
        Components =
        [
            new ComponentEntry { BoardLabel = boardLabel, FriendlyName = friendlyName, Category = "IC" },
        ],
    };

    // ------------------------------------------------------------------------ No draft

    [Fact]
    public async Task LoadBoardDataAsync_returns_the_official_data_when_there_is_no_draft()
    {
        this.WriteBoardExcel();

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("PLA", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ------------------------------------------------------------------------ A draft is read

    [Fact]
    public async Task LoadBoardDataAsync_reads_the_DRAFTS_workbook_when_one_exists()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (drafted correction)"));

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("PLA (drafted correction)", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ###########################################################################################
    // The drafted rows the UI marks are now DERIVED by comparing the draft against the published
    // copy, rather than read from a recorded delta list. This proves that derivation reaches
    // LastLoadedDraftSummary, which is what the component list's "drafted" chip reads.
    // ###########################################################################################
    [Fact]
    public async Task A_changed_row_is_reported_in_LastLoadedDraftSummary()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (drafted correction)"));

        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.Contains("U1", DataManager.LastLoadedDraftSummary.DraftedComponentBoardLabels);
    }

    [Fact]
    public async Task An_UNCHANGED_drafted_row_is_NOT_reported_as_drafted()
    {
        // The other half, and the one that would break loudly if seeding or comparison were wrong:
        // a draft whose rows match the published board exactly has changed nothing, so nothing
        // should be marked. Without this, a freshly seeded draft would light up every row.
        this.WriteBoardExcel();

        BoardData? published = await DataManager.LoadBoardDataAsync(this.BoardEntry);
        Assert.NotNull(published);

        this.WriteDraft(published!);

        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.Empty(DataManager.LastLoadedDraftSummary.DraftedComponentBoardLabels);
    }

    [Fact]
    public async Task LoadBoardDataAsync_still_returns_the_official_data_for_a_board_with_a_different_draft()
    {
        // A draft for a DIFFERENT board must never bleed into this one's load - proves the source
        // is resolved per-board (by ExcelDataFile), not applied blindly.
        this.WriteBoardExcel();

        string otherWorkbook = DraftFolderLayout.GetWorkbookPath(
            this.DraftsRoot, "Commodore/C64/250425/Data C64 250425 v1.0.0.xlsx");

        Directory.CreateDirectory(Path.GetDirectoryName(otherWorkbook)!);
        CachedWorkbooks.Write(
            otherWorkbook,
            DataManagerDraftOverlayTests.BoardWithComponent("U1", "Should not appear here"));

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, "Commodore/C64/250425/Data C64 250425 v1.0.0.xlsx"),
            new DraftMarker { BoardKey = "Commodore/C64/250425/Data C64 250425 v1.0.0.xlsx" });

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("PLA", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ------------------------------------------------------------------------ ViewOfficialPublishedOnly

    [Fact]
    public async Task ViewOfficialPublishedOnly_reads_the_published_workbook_instead()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (should be hidden)"));

        UserSettings.ViewOfficialPublishedOnly = true;

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("PLA", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ###########################################################################################
    // The toggle changes which FILE is read, not whether the draft is known about - so a "you are
    // viewing the published version" banner still has something to read, and the Drafts tab does
    // not appear to lose the draft while the toggle is on.
    //
    // *** WHAT IS REPORTED CHANGED, AND THIS EXPECTATION WAS UPDATED DELIBERATELY. *** It used to
    // assert LastLoadedDraftSummary still listed the drafted ROWS. It cannot now: the summary is
    // derived by comparing what was LOADED against the published copy, and with this toggle on
    // the published copy IS what was loaded - so there is nothing to compare and no rows to list.
    // The draft is instead reported through the marker-derived facts, which is the honest answer:
    // "a draft exists, and you are not looking at it".
    // ###########################################################################################
    [Fact]
    public async Task ViewOfficialPublishedOnly_still_knows_the_draft_is_there()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (drafted)"));

        UserSettings.ViewOfficialPublishedOnly = true;

        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        // The base revision comes off the marker, which is read whether or not the draft is shown.
        Assert.Equal("2026-09-01", DataManager.LastLoadedDraftBaseRevision);

        // And nothing is marked as drafted, because the published board is what is on screen -
        // marking rows there would tint a board that carries none of the contributor's edits.
        Assert.Empty(DataManager.LastLoadedDraftSummary.DraftedComponentBoardLabels);
    }

    [Fact]
    public async Task Turning_ViewOfficialPublishedOnly_off_again_reads_the_draft_once_more()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (drafted)"));

        UserSettings.ViewOfficialPublishedOnly = true;
        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        UserSettings.ViewOfficialPublishedOnly = false;
        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.Equal("PLA (drafted)", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ###########################################################################################
    // *** THE PUBLISHED WORKBOOK IS PARSED ONCE, NOT TWICE (code review, 2026-09-25). ***
    //
    // A drafted load also reads the published copy, to work out which rows are drafted. That
    // read used its own "published:" cache key - a leftover from when the drafted load was cached
    // under the board identity - so toggling "view boards as officially published" parsed and
    // held the same workbook a second time under its bare path.
    //
    // Proved by what the second load can still see: after the drafted load, the published file
    // is overwritten with garbage. The toggled load is then only right if it is served from the
    // entry the comparison already made - a fresh parse of the garbage would answer null.
    // ###########################################################################################
    [Fact]
    public async Task The_published_view_reuses_the_workbook_the_drafted_load_already_parsed()
    {
        this.WriteBoardExcel();
        this.WriteDraft(DataManagerDraftOverlayTests.BoardWithComponent("U1", "PLA (drafted)"));

        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        string publishedPath = Path.Combine(
            this.thisWorkspace.Root,
            this.thisExcelDataFile.Replace('/', Path.DirectorySeparatorChar));

        File.WriteAllText(publishedPath, "this is no longer a workbook");

        UserSettings.ViewOfficialPublishedOnly = true;
        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("PLA", result!.Components.Single(c => c.BoardLabel == "U1").FriendlyName);
    }

    // ---------------------------------- A board that exists only as a draft

    private void SaveNewBoardDraft(params ComponentEntry[] components)
    {
        var board = new BoardData();
        board.Components.AddRange(components);

        this.WriteDraft(board, new NewBoardRegistration
        {
            HardwareName = "Commodore 64",
            BoardName = "MyBoard",
            ExcelDataFile = this.thisExcelDataFile,
        });
    }

    // Deliberately NO WriteBoardExcel call in any of these - the absence of the published file is
    // the whole point. A board created through "Add a new board" has none and never will.
    [Fact]
    public async Task A_draft_only_board_loads_even_though_it_has_no_official_board_file()
    {
        this.SaveNewBoardDraft();

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Empty(result!.Components);
    }

    [Fact]
    public async Task A_draft_only_boards_own_rows_are_what_it_renders()
    {
        this.SaveNewBoardDraft(
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" });

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Equal("CPU", result!.Components.Single(c => c.BoardLabel == "U8").FriendlyName);
    }

    [Fact]
    public async Task A_draft_only_board_is_reported_as_a_new_board()
    {
        this.SaveNewBoardDraft();

        await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.True(DataManager.LastLoadedDraftIsNewBoard);
    }

    // ###########################################################################################
    // With the toggle on, a draft-only board renders BLANK rather than failing to open.
    //
    // Officially it does not exist, so "nothing here" is the truthful answer - and the alternative
    // (refusing to load) would leave the contributor unable to turn the toggle back off from a
    // board they cannot open.
    // ###########################################################################################
    [Fact]
    public async Task A_draft_only_board_shows_as_a_blank_board_when_viewing_officially_published_only()
    {
        this.SaveNewBoardDraft(
            new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", Category = "IC" });

        UserSettings.ViewOfficialPublishedOnly = true;

        BoardData? result = await DataManager.LoadBoardDataAsync(this.BoardEntry);

        Assert.NotNull(result);
        Assert.Empty(result!.Components);
    }
}
