using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A board folder put into Drafts/ by hand, taken all the way through the board-list load
// (owner report, 2026-09-27: "I added one new board into ...\Drafts but as I cannot see the
// 'Draft' tab, then I cannot submit it").
//
// DraftFolderImportTests pins each rule of the import on its own. THIS file pins the three pieces
// the owner's case needs TOGETHER, because each one alone still leaves the board unsubmittable:
//
//   - the import, or the folder is not a draft and the Drafts tab never lists it;
//   - the one-entry-per-folder rule in DataManager.MergeDraftOnlyBoards, or a board made the old
//     way is listed twice - under its _UserContribution names and under its folder names;
//   - the empty baseline for a new board, or its rows are compared against the contributor's
//     own legacy copy in Data/, count as nothing, and Submit stays disabled.
//
// "DataManager" collection: DataManager's lists and DraftManager's root are both statics.
// ###########################################################################################
[Collection("DataManager")]
public sealed class DataManagerDraftImportTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private const string MasterName = "Classic-Repair-Toolbox.xlsx";

    // What DataManager derives from MasterName - see thisGetUserContributionMainExcelRelativePath.
    private const string ContributionName = "Classic-Repair-Toolbox_UserContribution.xlsx";

    private const string PublishedKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    private const string LegacyKey = "Commodore/C128/310378 Open128/Data C128 310378 Open128.xlsx";

    private static readonly string[] HardwareHeaders =
    {
        "Hardware name in drop-down", "Board name in drop-down",
        "Excel data file", "Hardware notes in \"Overview\" tab"
    };

    static DataManagerDraftImportTests()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Classic Repair Toolbox tests");
    }

    public DataManagerDraftImportTests()
    {
        DraftManager.LoadFrom(this.DraftsRoot);
    }

    public void Dispose()
    {
        DataManager.LoadFrom(this.thisWorkspace.Root, "does-not-exist.xlsx");
        DraftManager.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    [Fact]
    public void A_board_made_the_old_way_and_copied_into_Drafts_is_listed_once_and_can_be_submitted()
    {
        this.WriteHardwareList(MasterName, new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" });
        this.WriteHardwareList(ContributionName, new[] { "Commodore 128", "310378 Open128", LegacyKey, "Has 6581 SID." });

        // The legacy copy in Data/, and the same board copied into Drafts/ with no marker.
        this.WriteBoard(DraftBoardSource.PublishedPathOf(this.thisWorkspace.Root, LegacyKey), components: 3);
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, LegacyKey), components: 3);

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        // Once, and where the contributor already knows it - under its _UserContribution names.
        HardwareBoardEntry entry = Assert.Single(
            DataManager.HardwareBoards,
            board => BoardDescriptorRules.BoardIdFromExcelDataFile(board.ExcelDataFile) == "Commodore/C128/310378 Open128");
        Assert.Equal("Commodore 128", entry.HardwareName);

        // On the Drafts tab - once.
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));

        // A new board with all three rows to send, so Submit is enabled.
        DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, entry.ExcelDataFile)!;
        Assert.True(status.IsNewBoard);
        Assert.Equal(3, DraftStatusReader.CountChanges(status));

        // Registered under its FOLDER names - the ones a submission must send.
        Assert.Equal("C128", status.NewBoard!.HardwareName);
        Assert.Equal("310378 Open128", status.NewBoard.BoardName);
    }

    [Fact]
    public void A_hand_placed_board_nothing_lists_appears_as_a_new_board_of_its_own()
    {
        const string key = "Amstrad/CPC 6128/MC0020/Data CPC 6128 MC0020.xlsx";

        this.WriteHardwareList(MasterName, new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" });
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, key), components: 2);

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == key);
        Assert.True(entry.IsDraftOnly);
        Assert.Equal("CPC 6128", entry.HardwareName);
        Assert.Equal("MC0020", entry.BoardName);
        Assert.Contains(entry, DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards));
    }

    [Fact]
    public void A_hand_placed_copy_of_a_published_board_becomes_its_draft_and_counts_only_what_differs()
    {
        this.WriteHardwareList(MasterName, new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" });
        this.WriteBoard(DraftBoardSource.PublishedPathOf(this.thisWorkspace.Root, PublishedKey), components: 3);
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, PublishedKey), components: 4);

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        HardwareBoardEntry entry = Assert.Single(DataManager.HardwareBoards);
        Assert.Same(entry, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));

        DraftStatus status = DraftStatusReader.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, PublishedKey)!;
        Assert.False(status.IsNewBoard);
        Assert.Equal(1, DraftStatusReader.CountChanges(status));
    }

    // ###########################################################################################
    // *** THE DRAFT THAT VANISHED (owner report, 2026-09-27). *** Open128 - made the old way, copied
    // into Drafts/ and submitted - was merged into BETA at 12:48, and at 12:49 CRT deleted the
    // draft "because its changes are now in the published data". They were not: retirement
    // compared the draft with the file at its path in Data/, which for a _UserContribution board is
    // the contributor's OWN copy - identical, since the draft was copied from it. Only a board the
    // MAIN workbook lists is published; the legacy board's draft must stay.
    // ###########################################################################################
    [Fact]
    public async Task A_board_made_the_old_way_is_not_retired_against_the_contributors_own_copy()
    {
        this.WriteHardwareList(MasterName, new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" });
        this.WriteHardwareList(ContributionName, new[] { "Commodore 128", "310378 Open128", LegacyKey, "" });
        this.WriteBoard(DraftBoardSource.PublishedPathOf(this.thisWorkspace.Root, LegacyKey), components: 3);
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, LegacyKey), components: 3);

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        HardwareBoardEntry legacy = Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == LegacyKey);
        Assert.True(legacy.IsUserContribution);
        Assert.False(legacy.IsPublished);
        Assert.True(Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == PublishedKey).IsPublished);

        Assert.Equal([PublishedKey], DataManager.PublishedExcelDataFiles());

        List<SubmissionReceipt> receipts = [Published("Commodore/C128/310378 Open128")];

        IReadOnlyList<RetirableDraft> retirable = await PublishedDraftRetirer.FindAsync(
            receipts, DataManager.DataRoot, DraftManager.DraftsRoot, DataManager.PublishedExcelDataFiles());

        Assert.Empty(retirable);

        // What the list passed before did: every non-draft-only board, the legacy one included -
        // and the draft came back as retirable. This is the deletion the owner saw.
        IReadOnlyList<RetirableDraft> withTheOldList = await PublishedDraftRetirer.FindAsync(
            receipts,
            DataManager.DataRoot,
            DraftManager.DraftsRoot,
            DataManager.HardwareBoards.Where(board => !board.IsDraftOnly).Select(board => board.ExcelDataFile).ToList());

        Assert.Single(withTheOldList);
    }

    // The other half: a draft of a board that IS published, identical to it and merged, still goes.
    [Fact]
    public async Task A_draft_of_a_really_published_board_is_still_retired_once_it_matches()
    {
        this.WriteHardwareList(MasterName, new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" });
        this.WriteBoard(DraftBoardSource.PublishedPathOf(this.thisWorkspace.Root, PublishedKey), components: 3);
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, PublishedKey), components: 3);

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        IReadOnlyList<RetirableDraft> retirable = await PublishedDraftRetirer.FindAsync(
            [Published("Commodore/C64/250407")],
            DataManager.DataRoot,
            DraftManager.DraftsRoot,
            DataManager.PublishedExcelDataFiles());

        Assert.Single(retirable);
    }

    // ###########################################################################################
    // *** A NEW BOARD'S DRAFT STAYS REACHABLE ONCE BETA LISTS IT (code review, 2026-09-27). *** The
    // publish names the workbook with the tree's generation ("... v2.0.0.xlsx"), so the listed
    // entry cannot read the draft, which keeps its own name. Skipping the draft's entry because its
    // FOLDER was listed left the draft unreachable until production retired it - no edit, no
    // resubmission, even after a push-back asked for one. The draft keeps its entry, reads its own
    // workbook, and is ONE row on the Drafts tab.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_draft_stays_reachable_after_BETA_lists_the_board_under_another_workbook_name()
    {
        const string draftKey = "Commodore/C128/310378 Open128/Data C128 310378 Open128.xlsx";
        const string listedKey = "Commodore/C128/310378 Open128/Data C128 310378 Open128 v2.0.0.xlsx";

        this.WriteHardwareList(MasterName,
            new[] { "Commodore 64", "250407 (long board)", PublishedKey, "" },
            new[] { "Commodore 128", "310378 Open128", listedKey, "" });
        this.WriteBoard(DraftBoardSource.PublishedPathOf(this.thisWorkspace.Root, listedKey), components: 3);

        // The contributor's new-board draft, as "Add a new board" made it.
        this.WriteBoard(DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, draftKey), components: 4);
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, draftKey), new DraftMarker
        {
            BoardKey = draftKey,
            NewBoard = new NewBoardRegistration { HardwareName = "C128", BoardName = "310378 Open128", ExcelDataFile = draftKey },
        });

        DataManager.LoadFrom(this.thisWorkspace.Root, MasterName);

        HardwareBoardEntry draft = Assert.Single(DataManager.HardwareBoards, board => board.ExcelDataFile == draftKey);
        Assert.True(draft.IsDraftOnly);
        Assert.Contains(DataManager.HardwareBoards, board => board.ExcelDataFile == listedKey);

        // One row on the Drafts tab - the entry that reads the draft.
        Assert.Same(draft, Assert.Single(DraftManager.EnumerateDraftedBoards(DataManager.HardwareBoards)));
        Assert.True(DraftBoardSource.Resolve(DataManager.DataRoot, DraftManager.DraftsRoot, draft.ExcelDataFile).IsDraft);
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

    private void WriteHardwareList(string fileName, params string?[][] rows) =>
        new BoardWorkbookBuilder()
            .Sheet("Hardware & Board", HardwareHeaders, rows)
            .SaveTo(Path.Combine(this.thisWorkspace.Root, fileName));

    private void WriteBoard(string path, int components)
    {
        var board = new BoardData { RevisionDate = "2026-August-7" };

        for (int i = 0; i < components; i++)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = $"U{i}", Description = "part" });
        }

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        CachedWorkbooks.Write(path, board);
    }
}
