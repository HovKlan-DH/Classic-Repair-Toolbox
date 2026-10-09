using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftFolderImport - a board folder put into Drafts/ by hand becomes a draft (owner request,
// 2026-09-27: "the CRT App can kind of 'import' that into its Draft so it can be submitted").
//
// Without it such a folder has no marker, so nothing in the application sees it: no Drafts tab,
// no Submit. The import writes the marker and nothing else decides anything differently - which
// is why several tests below finish by asking the ORDINARY draft readers (HasDraft,
// DraftStatusReader) rather than the import's own outcome.
//
// "BoardData" collection: these tests write and read board workbooks, like DraftStatusReaderTests.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftFolderImportTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

    private const string PublishedKey = "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx";

    // The owner's own case: a board made the old way, listed only by a _UserContribution workbook.
    private const string LegacyKey = "Commodore/C128/310378 Open128/Data C128 310378 Open128 v2.0.0.xlsx";

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    public void Dispose() => this.thisWorkspace.Dispose();

    // ------------------------------------------------------------------ the two kinds of draft

    [Fact]
    public void A_hand_placed_copy_of_a_PUBLISHED_board_becomes_its_draft()
    {
        this.WriteDraftWorkbook(PublishedKey, BoardWith(3, "2026-August-7"));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));

        Assert.Equal(DraftFolderImportKind.DraftOfPublishedBoard, outcome.Kind);
        Assert.Equal(PublishedKey, outcome.ExcelDataFile);
        Assert.True(DraftBoardSource.HasDraft(this.DraftsRoot, PublishedKey));

        DraftStatus? status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, PublishedKey);
        Assert.NotNull(status);
        Assert.False(status!.IsNewBoard);
    }

    // ###########################################################################################
    // *** THE BASE REVISION IS THE ONE INSIDE THE COPY, never the published board's current one. ***
    // The copy was taken from whatever revision it says it was; stamping today's published
    // revision would claim the contributor had seen changes they never saw, and the drift warning
    // could then never tell them the board has moved on under their copy.
    // ###########################################################################################
    [Fact]
    public void A_published_boards_copy_keeps_the_revision_written_inside_it_as_its_base()
    {
        this.WriteDraftWorkbook(PublishedKey, BoardWith(1, "2026-August-7"));
        this.WritePublishedWorkbook(PublishedKey, BoardWith(1, "2026-September-20"));

        this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true));

        Assert.Equal("2026-August-7", this.MarkerOf(PublishedKey)!.BaseRevision);
    }

    [Fact]
    public void A_folder_nothing_lists_becomes_a_NEW_board_named_after_its_folders()
    {
        const string key = "Amstrad/CPC 6128/MC0020/Data CPC 6128 MC0020.xlsx";
        this.WriteDraftWorkbook(key, BoardWith(2));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import());

        Assert.Equal(DraftFolderImportKind.NewBoard, outcome.Kind);

        NewBoardRegistration registration = this.MarkerOf(key)!.NewBoard!;
        Assert.Equal("CPC 6128", registration.HardwareName);
        Assert.Equal("MC0020", registration.BoardName);
        Assert.Equal(key, registration.ExcelDataFile);
        Assert.Equal(string.Empty, this.MarkerOf(key)!.BaseRevision);
    }

    // ###########################################################################################
    // *** THE REGISTRATION'S NAMES MUST REBUILD THE FOLDER'S BOARD ID. *** A submission sends them
    // as the board's parts, and the server rebuilds the id from those parts and compares - a
    // display name there ("Commodore 128" for the "C128" folder) is refused as
    // identity.board_id_mismatch, which is what happened on 2026-09-23.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_registered_names_rebuild_its_folder_so_a_submission_is_accepted()
    {
        this.WriteDraftWorkbook(LegacyKey, BoardWith(2));

        this.Import(new KnownDraftBoard(LegacyKey, IsPublished: false));

        NewBoardRegistration registration = this.MarkerOf(LegacyKey)!.NewBoard!;

        Assert.Equal(
            BoardDescriptorRules.BoardIdFromExcelDataFile(LegacyKey),
            $"{NewBoardIdentity.ExtractManufacturer(registration.ExcelDataFile)}/{registration.HardwareName}/{registration.BoardName}");
    }

    // ###########################################################################################
    // *** THE OWNER'S CASE (2026-09-27). *** A board made the old way is listed by a
    // "_UserContribution" workbook and lives in Data/ - but it has never been PUBLISHED, so its copy
    // in Drafts/ is a NEW board. And its rows must count against NOTHING: compared with the legacy
    // copy in Data/ (the same board) every row was unchanged, the tab said "nothing added yet" and
    // Submit was disabled. Fails against the version that compared against the file at the
    // published path.
    // ###########################################################################################
    [Fact]
    public void A_board_only_a_legacy_UserContribution_workbook_lists_is_a_NEW_board_whose_rows_all_count()
    {
        BoardData board = BoardWith(4);
        this.WritePublishedWorkbook(LegacyKey, board);
        this.WriteDraftWorkbook(LegacyKey, board);

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(new KnownDraftBoard(LegacyKey, IsPublished: false)));

        Assert.Equal(DraftFolderImportKind.NewBoard, outcome.Kind);
        Assert.Equal(LegacyKey, outcome.ExcelDataFile);

        DraftStatus status = DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, LegacyKey)!;
        Assert.True(status.IsNewBoard);
        Assert.Equal(string.Empty, status.PublishedWorkbookPath);
        Assert.Equal(4, DraftStatusReader.CountChanges(status));
    }

    [Fact]
    public void A_published_listing_wins_over_a_local_one_for_the_same_folder()
    {
        this.WriteDraftWorkbook(PublishedKey, BoardWith(1));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(
            new KnownDraftBoard(PublishedKey, IsPublished: false),
            new KnownDraftBoard(PublishedKey, IsPublished: true)));

        Assert.Equal(DraftFolderImportKind.DraftOfPublishedBoard, outcome.Kind);
    }

    // ------------------------------------------------------------------ what is left alone

    [Fact]
    public void A_folder_that_is_already_a_draft_is_left_exactly_as_it_is()
    {
        this.WriteDraftWorkbook(PublishedKey, BoardWith(1, "2026-August-7"));
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, PublishedKey),
            new DraftMarker { BoardKey = PublishedKey, BaseRevision = "the original" });

        Assert.Empty(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));
        Assert.Equal("the original", this.MarkerOf(PublishedKey)!.BaseRevision);
    }

    [Fact]
    public void Running_it_again_imports_nothing_more()
    {
        this.WriteDraftWorkbook(PublishedKey, BoardWith(1));

        Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));
        Assert.Empty(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));
    }

    // ###########################################################################################
    // A shared folder can hold an .xlsx datasheet, and three levels down it looks exactly like a
    // board folder. Taking it for one would put a board called "Shared files" in the lists - and
    // the server refuses that name anyway (identity.reserved_folder).
    // ###########################################################################################
    [Fact]
    public void The_shared_folders_are_never_taken_for_a_board()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "Shared files", "Board local files", "Pinouts.xlsx"), "x");
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Generic shared files", "Datasheets", "TTL", "Table.xlsx"), "x");

        Assert.Empty(this.Import());
        Assert.False(File.Exists(Path.Combine(this.DraftsRoot, "Commodore", "Shared files", "Board local files", DraftFolderLayout.DraftMarkerFileName)));
    }

    [Fact]
    public void A_folder_with_no_workbook_is_reported_and_not_imported()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "C64", "250407", "Sheet1.png"), "x");

        DraftFolderImportOutcome outcome = Assert.Single(this.Import());

        Assert.False(outcome.Imported);
        Assert.Contains("no board workbook", outcome.Reason);
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, PublishedKey));
    }

    // Excel writes a "~$" owner file beside every workbook it has open.
    [Fact]
    public void Excels_owner_file_is_not_a_second_workbook()
    {
        const string key = "Amstrad/CPC 6128/MC0020/Data CPC 6128 MC0020.xlsx";
        this.WriteDraftWorkbook(key, BoardWith(1));
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Amstrad", "CPC 6128", "MC0020", "~$Data CPC 6128 MC0020.xlsx"), "x");

        Assert.True(Assert.Single(this.Import()).Imported);
        Assert.Equal(key, this.MarkerOf(key)!.NewBoard!.ExcelDataFile);
    }

    [Fact]
    public void Several_workbooks_in_a_new_boards_folder_are_not_guessed_between()
    {
        this.WriteDraftWorkbook("Amstrad/CPC 6128/MC0020/Data A.xlsx", BoardWith(1));
        this.WriteDraftWorkbook("Amstrad/CPC 6128/MC0020/Data B.xlsx", BoardWith(1));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import());

        Assert.False(outcome.Imported);
        Assert.False(File.Exists(Path.Combine(this.DraftsRoot, "Amstrad", "CPC 6128", "MC0020", DraftFolderLayout.DraftMarkerFileName)));
    }

    // ###########################################################################################
    // On Linux "commodore/c64" is a different folder from "Commodore/C64", and the published tree
    // is case-sensitive - so a folder named like a known board in other capitals is neither that
    // board's draft nor a new board beside it. Refused, with the reason, rather than guessed.
    // ###########################################################################################
    [Fact]
    public void Folder_names_differing_only_in_capitals_from_a_known_board_are_not_imported()
    {
        this.thisWorkspace.WriteFile(Path.Combine("Drafts", "Commodore", "c64", "250407", "Data C64 250407 v2.0.0.xlsx"), "x");

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));

        Assert.False(outcome.Imported);
        Assert.Contains("capitals", outcome.Reason);
    }

    // The server refuses a board whose parts carry extra spaces (identity.parts_not_canonical).
    [Fact]
    public void A_new_board_whose_folder_name_the_server_would_refuse_is_not_imported()
    {
        this.WriteDraftWorkbook("Amstrad/CPC 6128/MC0020  rev B/Data.xlsx", BoardWith(1));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import());

        Assert.False(outcome.Imported);
        Assert.Contains("extra spaces", outcome.Reason);
    }

    [Fact]
    public void Nothing_happens_without_a_drafts_folder()
    {
        Assert.Empty(DraftFolderImport.ImportUnmarkedFolders(string.Empty, [], Now));
        Assert.Empty(DraftFolderImport.ImportUnmarkedFolders(this.DraftsRoot, [], Now));
    }

    // ------------------------------------------------------------------ renaming to the known name

    // ###########################################################################################
    // Every draft lookup builds the workbook's path from the known board's file name, so a copy
    // under another name (an older data generation, say) would be a draft whose workbook is never
    // found. With exactly one workbook it is renamed - its sidecar with it, or the highlights would
    // be left beside a name nothing reads.
    // ###########################################################################################
    [Fact]
    public void A_known_boards_copy_under_another_name_is_renamed_together_with_its_sidecar()
    {
        this.WriteDraftWorkbook("Commodore/C64/250407/Data C64 250407.xlsx", BoardWith(2));
        string folder = Path.Combine(this.DraftsRoot, "Commodore", "C64", "250407");
        File.WriteAllText(Path.Combine(folder, "Data C64 250407.json"), "{}");

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));

        Assert.True(outcome.Imported);
        Assert.Equal("Data C64 250407.xlsx", outcome.RenamedFrom);
        Assert.True(File.Exists(Path.Combine(folder, "Data C64 250407 v2.0.0.xlsx")));
        Assert.True(File.Exists(Path.Combine(folder, "Data C64 250407 v2.0.0.json")));
        Assert.False(File.Exists(Path.Combine(folder, "Data C64 250407.xlsx")));
        Assert.False(File.Exists(Path.Combine(folder, "Data C64 250407.json")));

        Assert.True(File.Exists(DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, PublishedKey)!.WorkbookPath));
    }

    [Fact]
    public void A_known_boards_folder_with_several_workbooks_none_rightly_named_is_left_alone()
    {
        this.WriteDraftWorkbook("Commodore/C64/250407/Data A.xlsx", BoardWith(1));
        this.WriteDraftWorkbook("Commodore/C64/250407/Data B.xlsx", BoardWith(1));

        DraftFolderImportOutcome outcome = Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));

        Assert.False(outcome.Imported);
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, PublishedKey));
    }

    // ###########################################################################################
    // A workbook open in Excel cannot be renamed on Windows. The import must then change NOTHING -
    // in particular the sidecar, which moves first, has to be moved back - so the next start can
    // simply try again.
    // ###########################################################################################
    [Fact]
    public void A_rename_blocked_by_an_open_workbook_changes_nothing_and_is_retried_next_time()
    {
        Assert.SkipUnless(OperatingSystem.IsWindows(), "File locks are only mandatory on Windows.");

        this.WriteDraftWorkbook("Commodore/C64/250407/Data C64 250407.xlsx", BoardWith(1));
        string folder = Path.Combine(this.DraftsRoot, "Commodore", "C64", "250407");
        string workbook = Path.Combine(folder, "Data C64 250407.xlsx");
        File.WriteAllText(Path.Combine(folder, "Data C64 250407.json"), "{}");

        DraftFolderImportOutcome outcome;

        using (new FileStream(workbook, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            outcome = Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true)));
        }

        Assert.False(outcome.Imported);
        Assert.True(File.Exists(workbook));
        Assert.True(File.Exists(Path.Combine(folder, "Data C64 250407.json")));
        Assert.False(File.Exists(Path.Combine(folder, "Data C64 250407 v2.0.0.json")));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, PublishedKey));

        Assert.True(Assert.Single(this.Import(new KnownDraftBoard(PublishedKey, IsPublished: true))).Imported);
    }

    // ###########################################################################################
    // *** AN UNEXPECTED FAILURE COSTS ONE FOLDER, NEVER THE BOARD LIST (code review, 2026-09-27). ***
    // The import runs inside DataManager's load of the main workbook, before the list of boards is
    // set, so an exception escaping it left the application with no boards at all. Forced here with
    // a known workbook name holding a NUL - which File.Move refuses with an ArgumentException, not
    // an IOException - so the rename of one folder fails in a way the import never expected.
    // ###########################################################################################
    [Fact]
    public void An_unexpected_failure_in_one_folder_is_reported_for_it_and_the_others_are_still_taken_in()
    {
        const string broken = "Commodore/C64/250407/Data C64 250407\0 v2.0.0.xlsx";
        const string newBoard = "Amstrad/CPC 6128/MC0020/Data CPC 6128 MC0020.xlsx";

        this.WriteDraftWorkbook("Commodore/C64/250407/Data C64 250407.xlsx", BoardWith(1));
        this.WriteDraftWorkbook(newBoard, BoardWith(1));

        IReadOnlyList<DraftFolderImportOutcome> outcomes = this.Import(new KnownDraftBoard(broken, IsPublished: true));

        Assert.Equal(2, outcomes.Count);
        Assert.Contains(outcomes, outcome => !outcome.Imported && outcome.BoardFolder.EndsWith("250407", StringComparison.Ordinal));
        Assert.Contains(outcomes, outcome => outcome.Imported && outcome.ExcelDataFile == newBoard);
    }

    // ------------------------------------------------------------------ helpers

    private IReadOnlyList<DraftFolderImportOutcome> Import(params KnownDraftBoard[] known) =>
        DraftFolderImport.ImportUnmarkedFolders(this.DraftsRoot, known, Now);

    private DraftMarker? MarkerOf(string key) =>
        DraftMarkerStore.Load(DraftFolderLayout.GetMarkerPath(this.DraftsRoot, key));

    private static BoardData BoardWith(int components, string revision = "2026-09-01")
    {
        var board = new BoardData { RevisionDate = revision };

        for (int i = 0; i < components; i++)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = $"U{i}", Description = "part" });
        }

        return board;
    }

    // A workbook put into Drafts/ by hand - NO marker, which is the whole point.
    private void WriteDraftWorkbook(string key, BoardData board)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, key);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
    }

    private void WritePublishedWorkbook(string key, BoardData board)
    {
        string path = DraftBoardSource.PublishedPathOf(this.DataRoot, key);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
    }
}
