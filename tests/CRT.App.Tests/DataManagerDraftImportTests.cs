using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A board folder put into Drafts/ by hand, taken all the way through the board-list load
// (owner report, 2026-09-27: "I added one new board into ...\Drafts but as I cannot see the
// 'Draft' tab, then I cannot submit it").
//
// DraftFolderImportTests pins each rule of the import on its own. THIS file pins the pieces the
// owner's case needs TOGETHER, because each one alone still leaves the board unsubmittable:
//
//   - the import - in CRT 4.0 converting an Excel-era copy into a 4.0 draft (2026-10-24) - or
//     the folder is not a draft and the Drafts tab never lists it;
//   - one entry per folder (DataManager.MergeDraftOnlyBoards), or the board is listed twice;
//   - the empty baseline for a new board, or its rows count as nothing and Submit stays disabled.
//
// A board made the old way, listed by a "_UserContribution" workbook, is no longer read from the
// data folder in 4.0: its copy in Drafts/ is how it is carried over.
//
// "DataManager" collection: DataManager's lists and DraftManager's root are both statics.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataManagerDraftImportTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private static readonly MainFiles.Board C64 = new("Commodore/C64/250407", "Commodore 64", "250407 (long board)");

    private const string LegacyBoardId = "Commodore/C128/310378 Open128";

    // The Excel workbook a board made the old way was copied with.
    private const string LegacyWorkbook = LegacyBoardId + "/Data C128 310378 Open128.xlsx";

    public DataManagerDraftImportTests()
    {
        DraftManager.LoadFrom(this.DraftsRoot);
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.json");
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => this.thisWorkspace.Root;

    [Fact]
    public void A_board_made_the_old_way_and_copied_into_Drafts_is_listed_once_and_can_be_submitted()
    {
        string main = MainFiles.Write(this.DataRoot, C64);
        this.WriteHandPlacedWorkbook(LegacyWorkbook, components: 3);

        DataManager.LoadFrom(this.DataRoot, main);

        // Once, under its folder names - the ones a submission must send.
        HardwareBoardEntry entry = Assert.Single(
            DataManager.HardwareBoards,
            board => BoardDescriptorRules.BoardIdFromExcelDataFile(board.ExcelDataFile) == LegacyBoardId);
        Assert.True(entry.IsDraftOnly);
        Assert.Equal("C128", entry.HardwareName);
        Assert.Equal(DataTreeLayout.BoardFilePath(LegacyBoardId), entry.ExcelDataFile);

        // On the Drafts tab - once.
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));

        // A new board with all three rows to send, so Submit is enabled - and a 4.0 draft now.
        DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile)!;
        Assert.True(status.IsNewBoard);
        Assert.Equal(3, DraftStatusReader.CountChanges(status));
        Assert.Equal(DataTreeLayout.BoardFileName, Path.GetFileName(status.WorkbookPath));

        Assert.Equal("C128", status.NewBoard!.HardwareName);
        Assert.Equal("310378 Open128", status.NewBoard.BoardName);
    }

    [Fact]
    public void A_hand_placed_board_nothing_lists_appears_as_a_new_board_of_its_own()
    {
        string main = MainFiles.Write(this.DataRoot, C64);
        this.WriteHandPlacedWorkbook("Amstrad/CPC 6128/MC0020/Data CPC 6128 MC0020.xlsx", components: 2);

        DataManager.LoadFrom(this.DataRoot, main);

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == "Amstrad/CPC 6128/MC0020/board.json");
        Assert.True(entry.IsDraftOnly);
        Assert.Equal("CPC 6128", entry.HardwareName);
        Assert.Equal("MC0020", entry.BoardName);
        Assert.Contains(entry, DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards));
    }

    [Fact]
    public void A_hand_placed_copy_of_a_published_board_becomes_its_draft_and_counts_only_what_differs()
    {
        string main = MainFiles.Write(this.DataRoot, C64);
        this.WritePublishedBoard(C64.Identity, components: 3);
        this.WriteHandPlacedWorkbook(C64.BoardId + "/Data C64 250407 v3.0.0.xlsx", components: 4);

        DataManager.LoadFrom(this.DataRoot, main);

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards);
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));

        DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, C64.Identity)!;
        Assert.False(status.IsNewBoard);
        Assert.Equal(1, DraftStatusReader.CountChanges(status));
    }

    // ###########################################################################################
    // *** THE DRAFT THAT VANISHED (owner report, 2026-09-27). *** Open128 - made the old way, copied
    // into Drafts/ and submitted - was merged into BETA at 12:48, and at 12:49 CRT deleted the
    // draft "because its changes are now in the published data". They were not. Only a board the
    // MAIN file lists is published; a draft of a board it does not list is never retired, whatever
    // its receipt says.
    // ###########################################################################################
    [Fact]
    public async Task A_board_the_main_file_does_not_list_is_not_retired()
    {
        string main = MainFiles.Write(this.DataRoot, C64);
        this.WriteHandPlacedWorkbook(LegacyWorkbook, components: 3);

        DataManager.LoadFrom(this.DataRoot, main);

        Assert.Equal([C64.Identity], DataManager.PublishedExcelDataFiles());

        IReadOnlyList<RetirableDraft> retirable = await PublishedDraftRetirer.FindAsync(
            [Published(LegacyBoardId)], DataManager.DataRoot, DraftManager.DraftsRoot, DataManager.PublishedExcelDataFiles());

        Assert.Empty(retirable);
    }

    // The other half: a draft of a board that IS published, identical to it and published, goes -
    // the Excel files its conversion set aside notwithstanding.
    [Fact]
    public async Task A_draft_of_a_really_published_board_is_still_retired_once_it_matches()
    {
        string main = MainFiles.Write(this.DataRoot, C64);
        this.WritePublishedBoard(C64.Identity, components: 3);
        this.WriteHandPlacedWorkbook(C64.BoardId + "/Data C64 250407 v3.0.0.xlsx", components: 3);

        DataManager.LoadFrom(this.DataRoot, main);

        IReadOnlyList<RetirableDraft> retirable = await PublishedDraftRetirer.FindAsync(
            [Published(C64.BoardId)],
            DataManager.DataRoot,
            DraftManager.DraftsRoot,
            DataManager.PublishedExcelDataFiles());

        Assert.Single(retirable);
    }

    // ###########################################################################################
    // *** A NEW BOARD'S DRAFT STAYS REACHABLE ONCE BETA LISTS IT (code review, 2026-09-27). *** A
    // draft made by CRT 3.x keeps the workbook name it was created with, while the main file lists
    // the board by its board file. In 4.0 both name the same folder, and every identity of a folder
    // reads its one board file - so the LISTED entry reads the draft, and it is ONE row on the
    // Drafts tab (2026-10-24; in 3.x the listed entry could not read it, and the draft kept an
    // entry of its own).
    // ###########################################################################################
    [Fact]
    public void A_new_boards_draft_stays_reachable_after_the_main_file_lists_the_board()
    {
        const string draftKey = LegacyWorkbook;
        var listed = new MainFiles.Board(LegacyBoardId, "Commodore 128", "310378 Open128");

        string main = MainFiles.Write(this.DataRoot, C64, listed);
        this.WritePublishedBoard(listed.Identity, components: 3);

        // The contributor's new-board draft, as "Add a new board" made it.
        Assert.True(DraftSeeder.CreateNewBoard(this.DraftsRoot, new NewBoardRegistration
        {
            HardwareName = "C128",
            BoardName = "310378 Open128",
            ExcelDataFile = draftKey,
        }).Created);
        Assert.True(DraftWorkbookStore.Edit(this.DraftsRoot, draftKey, board => WithComponents(board, 4)));

        DataManager.LoadFrom(this.DataRoot, main);

        Assert.DoesNotContain(DataManager.HardwareBoards, board => board.ExcelDataFile == draftKey);
        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == listed.Identity);

        // One row on the Drafts tab - the listed entry, which reads the draft.
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));
        BoardSourceSelection source = DraftBoardSource.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile);
        Assert.True(source.IsDraft);
        Assert.Equal(4, DraftWorkbookStore.LoadDraftBoard(DraftManager.DraftsRoot, entry.ExcelDataFile)!.Components.Count);
    }

    // ###########################################################################################
    // *** A DRAFT CRT 3.x MADE IS CONVERTED WHEN THE BOARDS ARE LISTED (2026-10-24). *** Its
    // workbook would otherwise be a draft with no board file - listed, unreadable, unsendable. The
    // load converts it before anything lists the boards, so it is an ordinary 4.0 draft at once.
    // ###########################################################################################
    [Fact]
    public void A_draft_CRT_3_made_is_a_4_0_draft_once_the_boards_are_listed()
    {
        const string crt3Workbook = "Commodore/C64/250407/Data C64 250407 v3.0.0.xlsx";

        string main = MainFiles.Write(this.DataRoot, C64);
        this.WritePublishedBoard(C64.Identity, components: 3);
        this.WriteHandPlacedWorkbook(crt3Workbook, components: 4);
        DraftMarkerStore.Save(
            Path.Combine(DraftFolderLayout.GetBoardFolder(this.DraftsRoot, crt3Workbook), DraftFolderLayout.DraftMarkerFileName),
            new DraftMarker { BoardKey = crt3Workbook, BaseRevision = "2026-August-7" });

        DataManager.LoadFrom(this.DataRoot, main);

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards);
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));

        DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, C64.Identity)!;
        Assert.True(File.Exists(status.WorkbookPath));
        Assert.Equal(DataTreeLayout.BoardFileName, Path.GetFileName(status.WorkbookPath));
        Assert.Equal(1, DraftStatusReader.CountChanges(status));
        Assert.Equal("2026-August-7", status.BaseRevision);
    }

    // "published" - in the production data, the only state retirement acts on.
    private static SubmissionReceipt Published(string boardId) => new()
    {
        SubmissionId = 8,
        BoardId = boardId,
        LastKnownState = "published",
        SentUtc = DateTimeOffset.UtcNow.AddMinutes(-3),
    };

    // ------------------------------------------------------------------ helpers

    private static BoardData WithComponents(BoardData board, int components)
    {
        var updated = new BoardData { RevisionDate = board.RevisionDate, HardwareName = board.HardwareName, BoardName = board.BoardName };

        for (int i = 0; i < components; i++)
        {
            updated.Components.Add(new ComponentEntry { BoardLabel = $"U{i}", Description = "part" });
        }

        return updated;
    }

    private static BoardData Board(int components) => WithComponents(new BoardData { RevisionDate = "2026-August-7" }, components);

    // An Excel workbook put into Drafts/ by hand, under the name it was copied with - no marker.
    private void WriteHandPlacedWorkbook(string workbookPath, int components) =>
        CachedWorkbooks.Write(Path.Combine([this.DraftsRoot, .. workbookPath.Split('/')]), Board(components));

    // The published board's file in the 4.0 tree.
    private void WritePublishedBoard(string identity, int components) =>
        CachedWorkbooks.Write(DraftBoardSource.PublishedPathOf(this.DataRoot, identity), Board(components));
}
