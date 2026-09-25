using System;
using System.Collections.Generic;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// When a local draft has done its job and can be deleted without asking (maintainer request,
// 2026-09-23).
//
// *** THIS DECIDES WHETHER TO DESTROY A FOLDER, WITH NO CONFIRMATION, so every test here is
// really asking one question: can this rule ever delete work that exists nowhere else? ***
// TabDrafts' manual Discard button confirms first precisely because it can; this path is only
// allowed to run when the draft has been PROVED redundant, so the bytes it removes are a
// duplicate of data the sync can hand back.
//
// The gap it fills: TabDrafts' own header said a draft "goes away when the work is actually
// published and the synced data starts carrying it", and nothing ever implemented that sentence.
// DraftManager.DiscardDraft had exactly one caller - the button.
//
// Shares the "BoardData" collection because BoardDataReader caches loaded boards in a static
// dictionary keyed by path, and these tests write and read workbooks.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftRetirementTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    // ------------------------------------------------------------------ IsPublishedState

    [Theory]
    [InlineData("published")]
    [InlineData("merged")]
    [InlineData("PUBLISHED")]
    [InlineData("  merged  ")]
    public void A_state_meaning_it_is_in_the_published_tree_counts(string state)
    {
        Assert.True(DraftRetirement.IsPublishedState(state));
    }

    // ###########################################################################################
    // *** "approved" AND "accepted" MUST NOT COUNT, and this is the test that keeps it that way. ***
    //
    // Both are past the reviewer's decision and both render as a cheerful green row, so treating
    // them as "published" is an easy and very damaging mistake: an approved submission has NOT
    // been written to the data tree yet. Retiring the draft at that point deletes the work before
    // anything carries it, and the next sync brings down a board that still lacks the change.
    // ###########################################################################################
    [Theory]
    [InlineData("approved")]
    [InlineData("accepted")]
    [InlineData("pending")]
    [InlineData("changes_requested")]
    [InlineData("rejected")]
    [InlineData("withdrawn")]
    [InlineData("")]
    [InlineData(null)]
    public void A_state_that_is_NOT_yet_in_the_published_tree_does_not_count(string? state)
    {
        Assert.False(DraftRetirement.IsPublishedState(state));
    }

    // ------------------------------------------------------------------ IsRetirable

    [Fact]
    public void A_draft_IDENTICAL_to_the_published_board_is_retirable()
    {
        // The whole point: the contributor's change has shipped, the sync has brought it back
        // down, and the draft now says nothing the published board does not.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        Assert.True(DraftRetirement.IsRetirable(this.Status()));
    }

    // ###########################################################################################
    // *** THE TEST THAT STOPS THIS BEING A DATA-LOSS BUG. ***
    //
    // The contributor kept working after submitting, so the draft holds an edit nobody has
    // published. A rule that trusted the server's "published" state alone would delete it.
    // ###########################################################################################
    [Fact]
    public void A_draft_STILL_DIFFERING_from_the_published_board_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU, second revision"));

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_with_an_EXTRA_row_is_NOT_retirable()
    {
        // Equality has to mean every row, not just the ones that happen to pair up. An added
        // component is work that exists only in the draft.
        BoardData published = DraftRetirementTests.BoardWith("CPU");

        BoardData draft = DraftRetirementTests.BoardWith("CPU");
        draft.Components.Add(new ComponentEntry { BoardLabel = "U9", Description = "VIC" });

        this.WritePublished(published);
        this.WriteDraft(draft);

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // ###########################################################################################
    // *** A SYSTEM THAT EXISTS ONLY AS A DRAFT IS NEVER RETIRED. ***
    //
    // There is no published board to compare against, so "no differences" would be vacuously true
    // against nothing at all - and the folder being deleted would be the only copy of a brand-new
    // system the contributor built from scratch. IsRetirable refuses one outright rather than
    // trusting the comparison to save it.
    // ###########################################################################################
    [Fact]
    public void A_DRAFT_ONLY_system_is_never_retirable()
    {
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftRetirementTests.SystemKey),
            new DraftMarker
            {
                SystemKey = DraftRetirementTests.SystemKey,
                BaseRevision = string.Empty,
                NewSystem = new NewSystemRegistration
                {
                    HardwareName = "C64",
                    BoardName = "250407",
                    ExcelDataFile = DraftRetirementTests.SystemKey,
                },
                CreatedUtc = "2026-09-23T00:00:00Z",
            });

        Assert.False(DraftRetirement.IsRetirable(
            DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftRetirementTests.SystemKey)));
    }

    [Fact]
    public void A_missing_PUBLISHED_workbook_is_NOT_retirable()
    {
        // The publish may be perfectly real and the sync simply not have run yet. Deleting here
        // would remove the draft before the data that replaces it has arrived.
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_missing_or_null_draft_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));

        // No draft workbook written, so Resolve finds no marker and answers null.
        Assert.False(DraftRetirement.IsRetirable(
            DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftRetirementTests.SystemKey)));

        Assert.False(DraftRetirement.IsRetirable(null));
    }

    [Fact]
    public void An_UNREADABLE_draft_workbook_is_NOT_retirable()
    {
        // Failing closed. A workbook that cannot be parsed might hold anything, and "I could not
        // read it" is not evidence that it is redundant.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        File.WriteAllText(
            DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftRetirementTests.SystemKey),
            "this is not a workbook");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // ------------------------------------------------------------------ Beyond the rows

    // ###########################################################################################
    // *** ROWS ARE NOT THE WHOLE DRAFT (code review, 2026-09-25). ***
    //
    // A draft folder also carries the KiCad calibrations (in the sidecar - BoardData has no
    // section for them) and the bytes of every file it references. IsRetirable used to compare
    // rows alone, so re-calibrating an overlay, or replacing a scan with a corrected one under
    // the same name, after submitting was deleted without asking once the published rows caught
    // up - the exact "kept working after submitting" case the class promises to protect.
    // ###########################################################################################
    [Fact]
    public void A_draft_with_a_DIFFERENT_KiCad_calibration_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWithSchematic("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWithSchematic("CPU"));

        this.WriteCalibration(this.PublishedWorkbook, offsetX: 10);
        this.WriteCalibration(this.DraftWorkbook, offsetX: 12);

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_with_a_calibration_the_published_board_LACKS_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWithSchematic("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWithSchematic("CPU"));

        this.WriteCalibration(this.DraftWorkbook, offsetX: 12);

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_whose_calibration_MATCHES_the_published_one_is_retirable()
    {
        // The anti-vacuity partner of the two above: a calibration that shipped must not keep the
        // draft alive for ever.
        this.WritePublished(DraftRetirementTests.BoardWithSchematic("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWithSchematic("CPU"));

        this.WriteCalibration(this.PublishedWorkbook, offsetX: 12);
        this.WriteCalibration(this.DraftWorkbook, offsetX: 12);

        Assert.True(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_file_REPLACED_under_the_same_name_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        this.WriteFile(this.PublishedFolder, "Sheet1.png", "the original scan");
        this.WriteFile(this.DraftFolder, "Sheet1.png", "a corrected scan");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_file_the_published_tree_does_NOT_have_is_NOT_retirable()
    {
        // A new file not yet published - or an Excel lock file, because the workbook is open
        // right now. Either way the folder is not a duplicate of anything.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        this.WriteFile(this.DraftFolder, Path.Combine("KiCad data", "board.kicad_pcb"), "traces");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_draft_whose_files_are_BYTE_IDENTICAL_to_the_published_ones_is_retirable()
    {
        // What seeding produces: every referenced file copied across unchanged.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        this.WriteFile(this.PublishedFolder, Path.Combine("Images", "Sheet1.png"), "same bytes");
        this.WriteFile(this.DraftFolder, Path.Combine("Images", "Sheet1.png"), "same bytes");

        Assert.True(DraftRetirement.IsRetirable(this.Status()));
    }

    // ------------------------------------------------------------------ The stamp

    // ###########################################################################################
    // *** THE CHECK AND THE DELETE HAPPEN AT DIFFERENT MOMENTS. *** The comparison runs off the
    // UI thread and takes a while, long enough for the contributor to save into that very draft.
    // FindRetirableDrafts stamps the folder BEFORE comparing, and the caller deletes only if
    // IsUnchangedSince still holds.
    // ###########################################################################################
    [Fact]
    public void A_draft_edited_after_it_was_found_retirable_is_no_longer_unchanged()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        RetirableDraft found = Assert.Single(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemKey, "published")],
            this.ResolveStatus));

        Assert.True(DraftRetirement.IsUnchangedSince(found));

        // A save lands in the draft between the check and the delete.
        this.WriteFile(this.DraftFolder, "notes.txt", "written just now");

        Assert.False(DraftRetirement.IsUnchangedSince(found));
    }

    [Fact]
    public void The_found_draft_names_its_own_folder()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        RetirableDraft found = Assert.Single(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemKey, "merged")],
            this.ResolveStatus));

        Assert.Equal(DraftRetirementTests.SystemKey, found.SystemId);
        Assert.Equal(this.DraftFolder, found.Folder);
    }

    // ------------------------------------------------------------------ FindRetirableSystems

    [Fact]
    public void Only_a_PUBLISHED_receipt_whose_draft_matches_is_named()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        IReadOnlyList<string> retirable = DraftRetirement.FindRetirableSystems(
            [
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemKey, "published"),
            ],
            this.ResolveStatus);

        Assert.Equal([DraftRetirementTests.SystemKey], retirable);
    }

    [Fact]
    public void A_receipt_that_is_not_published_names_nothing_even_when_the_draft_matches()
    {
        // The draft IS identical here, so the only thing stopping retirement is the state - which
        // is what makes this the anti-vacuity partner of the test above.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        IReadOnlyList<string> retirable = DraftRetirement.FindRetirableSystems(
            [
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemKey, "approved"),
            ],
            this.ResolveStatus);

        Assert.Empty(retirable);
    }

    [Fact]
    public void A_system_with_SEVERAL_published_receipts_is_named_once()
    {
        // Ordinary: submit, get published, edit again, submit again. Two published receipts, one
        // folder - and DiscardDraft on an already-deleted folder would be a second pointless call.
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        IReadOnlyList<string> retirable = DraftRetirement.FindRetirableSystems(
            [
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemKey, "published"),
                DraftRetirementTests.Receipt(2, DraftRetirementTests.SystemKey, "merged"),
            ],
            this.ResolveStatus);

        Assert.Single(retirable);
    }

    [Fact]
    public void An_absent_receipt_list_names_nothing_rather_than_throwing()
    {
        // Called on the launch path for every user, including one who has never submitted.
        Assert.Empty(DraftRetirement.FindRetirableSystems(null, this.ResolveStatus));
        Assert.Empty(DraftRetirement.FindRetirableSystems([], this.ResolveStatus));
    }

    // ------------------------------------------------------------------ helpers

    private DraftStatus? Status() =>
        DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftRetirementTests.SystemKey);

    private DraftStatus? ResolveStatus(string systemId) =>
        DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, systemId);

    private static SubmissionReceipt Receipt(long id, string systemId, string state) => new()
    {
        SubmissionId = id,
        SystemId = systemId,
        LastKnownState = state,
    };

    private string DraftWorkbook =>
        DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftRetirementTests.SystemKey);

    private string PublishedWorkbook =>
        DraftBoardSource.PublishedPathOf(this.DataRoot, DraftRetirementTests.SystemKey);

    private string DraftFolder => Path.GetDirectoryName(this.DraftWorkbook)!;

    private string PublishedFolder => Path.GetDirectoryName(this.PublishedWorkbook)!;

    private void WriteFile(string folder, string relative, string content)
    {
        string path = Path.Combine(folder, relative);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
    }

    private void WriteCalibration(string workbookPath, double offsetX) =>
        BoardComponentHighlightStorage.SaveKiCadCalibration(
            workbookPath, "Sheet 1", "board.kicad_pcb", offsetX, 20, 1.5, 1.5, mirrorX: false, mirrorY: false);

    // Calibrations are keyed per schematic and collected through the board's schematic list, so
    // a calibration test needs a board that has one.
    private static BoardData BoardWithSchematic(string description)
    {
        BoardData board = DraftRetirementTests.BoardWith(description);
        board.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 1" });
        return board;
    }

    private static BoardData BoardWith(string description) => new()
    {
        RevisionDate = "2026-09-01",
        Components =
        [
            new ComponentEntry { BoardLabel = "U8", Description = description },
        ],
    };

    private void WritePublished(BoardData board)
    {
        string path = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftRetirementTests.SystemKey);

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);

        // The reader caches by path, and these tests write several boards to the same path across
        // one run.
        BoardDataReader.ClearCache(path);
    }

    // Writes the draft workbook AND the marker that makes the folder a draft - a workbook alone is
    // not a draft, which is the rule DraftBoardSource exists to enforce.
    private void WriteDraft(BoardData board)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftRetirementTests.SystemKey);

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
        BoardDataReader.ClearCache(path);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftRetirementTests.SystemKey),
            new DraftMarker
            {
                SystemKey = DraftRetirementTests.SystemKey,
                BaseRevision = "2026-09-01",
                NewSystem = null,
                CreatedUtc = "2026-09-23T00:00:00Z",
            });
    }
}
