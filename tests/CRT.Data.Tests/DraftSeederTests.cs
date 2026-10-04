using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Seeding a draft folder as a full copy of the published board
// (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
//
// *** WHY "COMPLETE COPY" IS THE THING BEING TESTED. *** The draft is compared against the
// published board to work out what the contributor changed (BoardDataDiffer), and that
// comparison reads an ABSENT row as a deletion. Seed a draft partially and every untouched row
// reads as deleted - the change list becomes noise and a submission built from it asks the
// server to delete most of the board. So completeness is not tidiness here, it is correctness.
//
// Shares the "BoardData" collection because BoardDataReader caches loaded boards in a static
// dictionary keyed by path, and these tests read workbooks back.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftSeederTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    private static BoardData PublishedBoard() => new()
    {
        RevisionDate = "2026-09-01",
        Schematics =
        [
            new BoardSchematicEntry
            {
                SchematicName = "Sheet 1",
                SchematicImageFile = "Commodore/C64/250407/Sheet1.png",
            },
        ],
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", Description = "CPU" },
            new ComponentEntry { BoardLabel = "U9", Description = "VIC" },
        ],
        ComponentImages =
        [
            new ComponentImageEntry
            {
                BoardLabel = "U8",
                Pin = "1",
                Name = "Clock",
                // A manufacturer shared file - referenced, but NOT this system's to own.
                File = "Commodore/Shared files/scope.png",
            },
        ],
    };

    // Writes the published side to disk: the files the board's rows reference.
    private void WritePublishedFiles(params string[] relativePaths)
    {
        foreach (string relative in relativePaths)
        {
            string full = Path.Combine(this.DataRoot, relative.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(full)!);
            File.WriteAllText(full, "bytes for " + relative);
        }
    }

    // ------------------------------------------------------------------ Seeding a published board

    // ###########################################################################################
    // *** A DRAFT IS A COPY OF THE PUBLISHED FILE AS IT IS NOW, NOT AS CRT FIRST READ IT (owner
    // report, 2026-10-02). *** The project owner edited a published workbook in Excel with CRT
    // open, then saved a component - and the new draft had none of the Excel edits: it was seeded
    // from the board CRT had cached when the board was first shown. Here the cache holds the old
    // board, the file is changed behind it, and the draft must hold the change.
    // ###########################################################################################
    [Fact]
    public async Task Seeding_from_the_file_copies_it_as_it_is_now_not_as_it_was_cached()
    {
        string published = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftSeederTests.SystemKey);
        Directory.CreateDirectory(Path.GetDirectoryName(published)!);

        BoardWorkbookWriter.Write(published, DraftSeederTests.PublishedBoard());
        BoardDataReader.ClearCache(published);
        BoardData? cached = await BoardDataReader.LoadAsync(published, published);
        Assert.Equal(2, cached!.Components.Count);

        BoardData edited = DraftSeederTests.PublishedBoard();
        edited.Components.Add(new ComponentEntry { BoardLabel = "U99", Description = "Typed in Excel" });
        BoardWorkbookWriter.Write(published, edited);

        try
        {
            DraftSeedResult seeded = DraftSeeder.SeedFromPublishedFile(this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey);

            Assert.True(seeded.Created, seeded.Reason);
            Assert.Contains(
                DraftWorkbookStore.LoadDraftBoard(this.DraftsRoot, DraftSeederTests.SystemKey)!.Components,
                component => component.BoardLabel == "U99");
        }
        finally
        {
            BoardDataReader.ClearCache(published);
        }
    }

    [Fact]
    public void Seeding_from_a_file_that_is_not_there_says_so_and_creates_nothing()
    {
        DraftSeedResult seeded = DraftSeeder.SeedFromPublishedFile(this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey);

        Assert.False(seeded.Created);
        Assert.NotEmpty(seeded.Reason);
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftSeederTests.SystemKey));
    }

    [Fact]
    public void Seeding_creates_a_workbook_the_app_can_read_back()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.True(result.Created, result.Reason);
        Assert.True(File.Exists(result.WorkbookPath));

        // The workbook carries the PUBLISHED file name, so the folder is indistinguishable from a
        // real board - see DraftFolderLayoutTests for why that matters.
        Assert.Equal("Data C64 250407.xlsx", Path.GetFileName(result.WorkbookPath));
    }

    // ###########################################################################################
    // *** THE LOAD-BEARING ONE: EVERY ROW IS COPIED, NOT JUST THE ONES SOMEBODY TOUCHED. ***
    //
    // A freshly seeded draft must compare EQUAL to the board it was seeded from. If it does not,
    // the Drafts tab reports changes the contributor never made, and every one of those is a row
    // a submission would ask the server to alter.
    // ###########################################################################################
    [Fact]
    public async Task A_freshly_seeded_draft_reports_NO_changes_against_the_published_board()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        BoardData published = DraftSeederTests.PublishedBoard();

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, published);

        Assert.True(result.Created, result.Reason);

        BoardData seeded = await DraftSeederTests.ReadWorkbookAsync(result.WorkbookPath);

        Assert.Empty(BoardDataDiffer.Compare(published, seeded));
    }

    [Fact]
    public void The_sidecar_is_copied_when_the_published_board_has_one()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        // The published sidecar, beside the published workbook.
        string publishedSidecar = Path.Combine(
            this.DataRoot, "Commodore", "C64", "250407", "Data C64 250407.json");

        Directory.CreateDirectory(Path.GetDirectoryName(publishedSidecar)!);
        File.WriteAllText(publishedSidecar, "{\"Component highlights\":{}}");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.True(result.Created, result.Reason);

        // Byte-copied rather than rewritten: the sidecar holds roots this app does not model
        // ("KiCad calibration points", and whatever a future version adds), and rewriting it from
        // a partial model would discard them.
        string draftSidecar = DraftFolderLayout.GetSidecarPath(this.DraftsRoot, DraftSeederTests.SystemKey);
        Assert.True(File.Exists(draftSidecar));
        Assert.Equal(File.ReadAllText(publishedSidecar), File.ReadAllText(draftSidecar));
    }

    [Fact]
    public void A_board_with_NO_sidecar_seeds_perfectly_well()
    {
        // Plenty of boards have no highlights at all. A missing sidecar is ordinary, not an error.
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.True(result.Created, result.Reason);
        Assert.False(File.Exists(DraftFolderLayout.GetSidecarPath(this.DraftsRoot, DraftSeederTests.SystemKey)));
    }

    // ------------------------------------------------------------------ Which files travel

    [Fact]
    public void The_boards_OWN_files_are_copied_into_the_draft()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        string copied = Path.Combine(
            this.DraftsRoot, "Commodore", "C64", "250407", "Sheet1.png");

        Assert.True(File.Exists(copied));
        Assert.Equal(1, result.FilesCopied);
    }

    // ###########################################################################################
    // *** A SHARED FILE IS LEFT WHERE IT IS. ***
    //
    // A manufacturer "Shared files" image belongs to many boards. Copying one into this draft
    // would FORK it: a later edit to the draft's copy would silently fail to reach the boards that
    // share it. It keeps resolving against Data/ as it always did, and is REPORTED so the caller
    // can say so rather than leaving it looking like a missing file.
    // ###########################################################################################
    [Fact]
    public void A_SHARED_file_is_not_copied_and_is_reported_as_shared()
    {
        this.WritePublishedFiles(
            "Commodore/C64/250407/Sheet1.png",
            "Commodore/Shared files/scope.png");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.Contains("Commodore/Shared files/scope.png", result.FilesShared);

        Assert.False(Directory.Exists(Path.Combine(this.DraftsRoot, "Commodore", "Shared files")));
    }

    // ###########################################################################################
    // A published board CAN reference a file the tree does not hold - a partial sync, a
    // hand-deleted image. That row was already broken before the draft existed, so seeding
    // reports it and carries on rather than refusing to create the draft at all.
    // ###########################################################################################
    [Fact]
    public void A_referenced_file_that_is_MISSING_is_reported_but_does_not_fail_the_seed()
    {
        // Deliberately writes nothing: every referenced file is absent.
        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.True(result.Created, result.Reason);
        Assert.Contains("Commodore/C64/250407/Sheet1.png", result.FilesMissing);
        Assert.Equal(0, result.FilesCopied);
    }

    // ###########################################################################################
    // Files come from the ROWS, never from a directory walk - the same rule a submission follows.
    // A walk would sweep up editor backups, thumbnail caches and whatever else sits in a board
    // folder, and in a draft that rubbish would then be submitted.
    // ###########################################################################################
    [Fact]
    public void A_file_no_row_references_is_NOT_copied()
    {
        this.WritePublishedFiles(
            "Commodore/C64/250407/Sheet1.png",
            "Commodore/C64/250407/~$Data C64 250407.xlsx",
            "Commodore/C64/250407/thumbs.db");

        DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        string draftFolder = Path.Combine(this.DraftsRoot, "Commodore", "C64", "250407");

        Assert.False(File.Exists(Path.Combine(draftFolder, "~$Data C64 250407.xlsx")));
        Assert.False(File.Exists(Path.Combine(draftFolder, "thumbs.db")));
    }

    // ------------------------------------------------------------------ The marker

    [Fact]
    public void The_marker_records_the_revision_the_draft_was_taken_from()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftSeederTests.SystemKey));

        Assert.NotNull(marker);

        // Without this the drift warning cannot tell "the published board moved on underneath you"
        // from "you edited these rows yourself" - and nothing else in the folder records it.
        Assert.Equal("2026-09-01", marker!.BaseRevision);
        Assert.Equal(DraftSeederTests.SystemKey, marker.SystemKey);
        Assert.False(marker.IsNewSystem);
    }

    // ###########################################################################################
    // *** RE-SEEDING WOULD SILENTLY DISCARD THE CONTRIBUTOR'S WORK. ***
    //
    // The single most destructive thing this class could do, and the easiest to trigger by
    // accident from a caller that does not check first. So an existing marker means "leave it
    // alone", and the caller is told why rather than being allowed to assume success.
    // ###########################################################################################
    [Fact]
    public async Task Seeding_REFUSES_to_overwrite_a_draft_that_already_exists()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        // The contributor's own edit, which must survive.
        string workbook = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftSeederTests.SystemKey);
        var edited = new BoardData
        {
            RevisionDate = "2026-09-01",
            Components = [new ComponentEntry { BoardLabel = "U8", Description = "MY EDIT" }],
        };

        BoardWorkbookWriter.Write(workbook, edited);

        DraftSeedResult second = DraftSeeder.SeedFromPublished(
            this.DraftsRoot, this.DataRoot, DraftSeederTests.SystemKey, DraftSeederTests.PublishedBoard());

        Assert.False(second.Created);
        Assert.Contains("already exists", second.Reason, StringComparison.OrdinalIgnoreCase);

        // And the edit is still there.
        BoardData after = await DraftSeederTests.ReadWorkbookAsync(workbook);
        Assert.Equal("MY EDIT", after.Components.Single(c => c.BoardLabel == "U8").Description);
    }

    // ------------------------------------------------------------------ A brand-new system

    [Fact]
    public async Task A_NEW_system_gets_an_empty_but_readable_workbook()
    {
        var registration = new NewSystemRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            ExcelDataFile = DraftSeederTests.SystemKey,
            CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
        };

        DraftSeedResult result = DraftSeeder.CreateNewSystem(this.DraftsRoot, registration);

        Assert.True(result.Created, result.Reason);
        Assert.True(File.Exists(result.WorkbookPath));

        // Readable and empty - the headers are there, which is what makes it hand-editable from
        // the moment it is created. Phase 2 rejected writing an .xlsx here; see DraftSeeder's own
        // header for why this is not that mistake.
        BoardData board = await DraftSeederTests.ReadWorkbookAsync(result.WorkbookPath);
        Assert.Empty(board.Components);
        Assert.Empty(board.Schematics);
    }

    // ###########################################################################################
    // *** A NEW SYSTEM IS SEEDED WITH NO REVISION DATE, AND THAT IS FINE - THE SERVER STAMPS IT.
    // *** It once was not: approving a new system answered "This submission cannot be published: A
    // publish must carry a revision" (reported by the project owner, 2026-09-26), because
    // PublishPlan was built from the SUBMITTED date and a new system has none to give.
    //
    // The project owner settled it the same day: "when the maintainer publish it to BETA, the
    // revision date gets updated from server. Same happens when it gets published to real
    // production, so server always wins, and what is typed by user is not important." So
    // ApprovePublishFlow.BuildPlan stamps the publish date and nothing downstream needs the draft
    // to carry one.
    //
    // This test holds the CLIENT half of that: the seeded workbook deliberately carries no
    // revision date, and the merge leaves it empty for a new system. Stamping one here instead
    // would be wrong in its own right - PublishExecutor's header explains that a locally stamped
    // draft immediately looks newer than the board it came from, so DraftRevisionComparer would
    // report drift on every draft. The server half is
    // ApprovePublishFlowTests.A_NEW_system_that_carries_no_revision_date_can_still_be_planned.
    // ###########################################################################################
    [Fact]
    public async Task A_NEW_system_is_seeded_with_no_revision_date_because_the_server_stamps_it()
    {
        var registration = new NewSystemRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            ExcelDataFile = DraftSeederTests.SystemKey,
        };

        DraftSeedResult result = DraftSeeder.CreateNewSystem(this.DraftsRoot, registration);
        Assert.True(result.Created, result.Reason);

        // 1. Nothing stamped a revision date into the seeded workbook.
        BoardData seeded = await DraftSeederTests.ReadWorkbookAsync(result.WorkbookPath);
        Assert.True(string.IsNullOrWhiteSpace(seeded.RevisionDate));

        // 2. With no published board to fall back to, the merged board carries none either.
        var manifest = new SubmissionManifest
        {
            SystemId = "Commodore/C64/250407",
            Manufacturer = "Commodore",
            Hardware = "C64",
            Board = "250407",
        };

        BoardData merged = PublishMerge.Build(manifest, published: null);

        Assert.Equal(string.Empty, PublishMerge.RevisionOf(merged));
    }

    // ###########################################################################################
    // *** A NEW SYSTEM GETS AN EMPTY "Scope baseline" FOLDER (owner request, 2026-09-24). ***
    //
    // Beside the workbook, so the contributor can pick it straight away when referencing their
    // first baseline image rather than first creating it by hand with the exact name.
    // ###########################################################################################
    [Fact]
    public void A_NEW_system_gets_an_empty_Scope_baseline_folder_beside_its_workbook()
    {
        var registration = new NewSystemRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            ExcelDataFile = DraftSeederTests.SystemKey,
        };

        DraftSeedResult result = DraftSeeder.CreateNewSystem(this.DraftsRoot, registration);

        Assert.True(result.Created, result.Reason);

        string expected = Path.Combine(
            Path.GetDirectoryName(result.WorkbookPath)!,
            "Scope baseline");

        Assert.True(Directory.Exists(expected), $"Expected [{expected}] to exist.");
        Assert.Empty(Directory.EnumerateFileSystemEntries(expected));
    }

    // A draft SEEDED from a published board is a copy of what the board actually has - it is not
    // handed an extra folder the published board never had.
    [Fact]
    public void A_SEEDED_draft_is_not_given_a_Scope_baseline_folder_its_board_does_not_have()
    {
        this.WritePublishedFiles("Commodore/C64/250407/Sheet1.png");

        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            this.DraftsRoot,
            this.DataRoot,
            DraftSeederTests.SystemKey,
            DraftSeederTests.PublishedBoard());

        Assert.True(result.Created, result.Reason);
        Assert.False(Directory.Exists(DraftFolderLayout.GetScopeBaselineFolder(this.DraftsRoot, DraftSeederTests.SystemKey)));
    }

    [Fact]
    public void A_NEW_systems_marker_carries_its_registration_and_NO_base_revision()
    {
        var registration = new NewSystemRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            HardwareNotes = "Test board",
            ExcelDataFile = DraftSeederTests.SystemKey,
        };

        DraftSeeder.CreateNewSystem(this.DraftsRoot, registration);

        DraftMarker? marker = DraftMarkerStore.Load(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftSeederTests.SystemKey));

        Assert.NotNull(marker);
        Assert.True(marker!.IsNewSystem);
        Assert.Equal("C64", marker.NewSystem!.HardwareName);

        // EMPTY, deliberately: this system has no published counterpart, so there is no revision
        // it is based on. Stamping today's date would make a later drift check compare against a
        // revision that never existed.
        Assert.Equal(string.Empty, marker.BaseRevision);
    }

    [Fact]
    public void Creating_a_NEW_system_also_refuses_to_overwrite_an_existing_draft()
    {
        var registration = new NewSystemRegistration { ExcelDataFile = DraftSeederTests.SystemKey };

        Assert.True(DraftSeeder.CreateNewSystem(this.DraftsRoot, registration).Created);
        Assert.False(DraftSeeder.CreateNewSystem(this.DraftsRoot, registration).Created);
    }

    [Fact]
    public void An_unresolvable_location_is_refused_with_a_reason_rather_than_throwing()
    {
        DraftSeedResult result = DraftSeeder.SeedFromPublished(
            draftsRoot: string.Empty,
            this.DataRoot,
            DraftSeederTests.SystemKey,
            DraftSeederTests.PublishedBoard());

        Assert.False(result.Created);
        Assert.NotEqual(string.Empty, result.Reason);
    }

    // ###########################################################################################
    // Reads a workbook back the way the app does.
    //
    // The cache key is the PATH rather than the system identity, so each test's own temp folder
    // gets its own cache entry - BoardDataReader keeps loaded boards in a static dictionary, and
    // sharing a key across tests would hand one test another's board.
    // ###########################################################################################
    private static async Task<BoardData> ReadWorkbookAsync(string path)
    {
        BoardData? board = await BoardDataReader.LoadAsync(path, path);
        Assert.NotNull(board);

        return board!;
    }
}
