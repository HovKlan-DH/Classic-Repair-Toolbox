using System;
using System.Collections.Generic;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// When a local draft has done its job and can be deleted without asking (owner request,
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
    // ###########################################################################################
    // *** "returned" IS NOT PUBLISHED (code review, 2026-09-27). *** It was merged and has been
    // taken back OUT of BETA, so its work is no longer in the published tree - and the draft may
    // be the only copy the contributor has to correct it from. Retiring it would delete exactly
    // the thing the rollback's mail tells them to use.
    // ###########################################################################################
    [Fact]
    public void A_submission_taken_back_out_of_BETA_does_not_retire_its_draft()
    {
        Assert.False(DraftRetirement.IsPublishedState("returned"));
        Assert.False(DraftRetirement.IsPublishedState("Returned"));
    }

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    // What a real SubmissionReceipt carries: the system's id, built from the workbook path when it
    // was submitted - never the workbook path itself. Receipts here used to carry SystemKey, which
    // no real receipt ever does, and that is how the app's lookup by id finding nothing went unseen.
    private static readonly string SystemId = SystemDescriptorRules.SystemIdFromExcelDataFile(DraftRetirementTests.SystemKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    // ------------------------------------------------------------------ IsPublishedState

    [Theory]
    [InlineData("published")]
    [InlineData("PUBLISHED")]
    [InlineData("  published  ")]
    public void A_state_meaning_it_is_in_the_PRODUCTION_data_counts(string state)
    {
        Assert.True(DraftRetirement.IsPublishedState(state));
    }

    // ###########################################################################################
    // *** BETA DOES NOT COUNT (owner decision, 2026-09-27). *** "merged" is in the BETA data only,
    // and a maintainer can still push it back to the queue - after which a draft retired at the
    // BETA stage was simply gone from the contributor's machine. It used to count; it must not.
    // ###########################################################################################
    [Theory]
    [InlineData("merged")]
    [InlineData("  MERGED  ")]
    [InlineData("returned")]
    public void A_submission_only_in_BETA_does_not_count(string state)
    {
        Assert.False(DraftRetirement.IsPublishedState(state));
    }

    // ###########################################################################################
    // *** "approved" AND "accepted" MUST NOT COUNT, and this is the test that keeps it that way. ***
    //
    // Both are past the maintainer's decision and both render as a cheerful green row, so treating
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
    // *** A SYSTEM THAT EXISTS ONLY AS A DRAFT IS NOT RETIRED WHILE NOTHING IS PUBLISHED. ***
    //
    // There is no published board to compare against, so "no differences" would be vacuously true
    // against nothing at all - and the folder being deleted would be the only copy of a brand-new
    // system the contributor built from scratch. Since 2026-09-25 a new system's draft IS retired
    // once its published board is here (below); this is the half that must still hold.
    // ###########################################################################################
    [Fact]
    public void A_DRAFT_ONLY_system_is_not_retirable_while_nothing_is_published()
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

    // ###########################################################################################
    // *** A SHARED FILE THE CONTRIBUTOR ATTACHED is compared with where it is PUBLISHED - the
    // shared folder - not with a copy inside the board's folder (2026-09-25). *** Such a file sits
    // in the draft under its whole path ("Commodore/Shared files/Component images/HotCPU.png").
    // Compared against the board's published folder it could never match, so a draft holding one
    // would never be retired after its work was published.
    // ###########################################################################################
    [Fact]
    public void A_drafted_SHARED_file_that_is_now_published_is_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        string shared = Path.Combine("Commodore", "Shared files", "Component images", "HotCPU.png");
        this.WriteFile(this.DataRoot, shared, "the new image");
        this.WriteFile(this.DraftFolder, shared, "the new image");

        Assert.True(DraftRetirement.IsRetirable(this.Status()));
    }

    [Fact]
    public void A_drafted_SHARED_file_that_is_not_published_yet_is_NOT_retirable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        this.WriteFile(this.DraftFolder, Path.Combine("Generic shared files", "Component images", "7408.jpg"), "new");

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

    // ------------------------------------------------------------------ A draft that lost its workbook

    // ###########################################################################################
    // *** A DISCARD THAT STOPPED PART-WAY LEFT A DRAFT WITH NO WORKBOOK (owner report, 2026-10-02;
    // cases agreed with the project owner). *** Retiring the ZX Spectrum Issue 4B draft deleted its
    // workbook and sidecar and then stopped at an image something had open, leaving the marker and
    // three images byte-identical to stable. With no workbook to compare, every later check refused
    // it, so it stayed - "0 rows changed", with the drift bar up - for good. The delete order is
    // fixed (DraftWorkbookStore.Discard); these are the rule that clears a draft ALREADY in that
    // state: retired when nothing left in its folder differs from the published board.
    // ###########################################################################################

    // Case 1.
    [Fact]
    public void A_published_draft_that_lost_its_workbook_and_holds_only_files_identical_to_stable_is_retired()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteFile(this.PublishedFolder, "Issue 4.B.png", "the published scan");

        this.WriteDraftThatLostItsWorkbook();
        this.WriteFile(this.DraftFolder, "Issue 4.B.png", "the published scan");

        RetirableDraft found = Assert.Single(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "published")],
            this.ResolveStatus));

        // And the board stops being a draft once the folder goes.
        Assert.True(DraftWorkbookStore.Discard(this.DraftsRoot, found.ExcelDataFile));
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftRetirementTests.SystemKey));
    }

    // Case 2, first half.
    [Fact]
    public void A_draft_that_lost_its_workbook_is_kept_when_an_image_differs_from_stable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteFile(this.PublishedFolder, "Issue 4.B.png", "the published scan");

        this.WriteDraftThatLostItsWorkbook();
        this.WriteFile(this.DraftFolder, "Issue 4.B.png", "a corrected scan");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // Case 2, second half.
    [Fact]
    public void A_draft_that_lost_its_workbook_is_kept_when_an_image_is_not_in_stable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));

        this.WriteDraftThatLostItsWorkbook();
        this.WriteFile(this.DraftFolder, "Issue 5.png", "a scan nobody has published");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // Case 3, first half: the highlights file was left behind, and it is stable's own.
    [Fact]
    public void A_draft_that_lost_its_workbook_is_retired_when_its_highlights_file_is_identical_to_stable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        File.WriteAllText(BoardComponentHighlightStorage.GetJsonPath(this.PublishedWorkbook), "{ \"highlights\": 1 }");

        this.WriteDraftThatLostItsWorkbook();
        File.WriteAllText(BoardComponentHighlightStorage.GetJsonPath(this.DraftWorkbook), "{ \"highlights\": 1 }");

        Assert.True(DraftRetirement.IsRetirable(this.Status()));
    }

    // Case 3, second half: a highlights file that differs may hold the contributor's own marking.
    [Fact]
    public void A_draft_that_lost_its_workbook_is_kept_when_its_highlights_file_differs_from_stable()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        File.WriteAllText(BoardComponentHighlightStorage.GetJsonPath(this.PublishedWorkbook), "{ \"highlights\": 1 }");

        this.WriteDraftThatLostItsWorkbook();
        File.WriteAllText(BoardComponentHighlightStorage.GetJsonPath(this.DraftWorkbook), "{ \"highlights\": 2 }");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // Case 4: the submission is still pending, or only in BETA.
    [Theory]
    [InlineData("pending")]
    [InlineData("merged")]
    public void A_draft_that_lost_its_workbook_is_kept_while_its_submission_is_not_published_to_stable(string state)
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraftThatLostItsWorkbook();

        Assert.Empty(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, state)],
            this.ResolveStatus));
    }

    // Case 5.
    [Fact]
    public void A_draft_that_lost_its_workbook_is_kept_while_the_stable_workbook_is_not_downloaded()
    {
        this.WriteDraftThatLostItsWorkbook();

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
    }

    // Case 6: with its workbook, the rows still decide, whatever the files say.
    [Fact]
    public void A_draft_that_still_has_its_workbook_is_judged_by_its_rows_as_before()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteFile(this.PublishedFolder, "Issue 4.B.png", "the published scan");

        this.WriteDraft(DraftRetirementTests.BoardWith("CPU, edited after submitting"));
        this.WriteFile(this.DraftFolder, "Issue 4.B.png", "the published scan");

        Assert.False(DraftRetirement.IsRetirable(this.Status()));
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
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "published")],
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
            [DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "published")],
            this.ResolveStatus));

        // The WORKBOOK key, which everything after the check is keyed by - not the receipt's id.
        Assert.Equal(DraftRetirementTests.SystemKey, found.ExcelDataFile);
        Assert.Equal(this.DraftFolder, found.Folder);
    }

    // ###########################################################################################
    // *** THE REPORTED BUG (owner, 2026-09-25): published to BETA, the synced board matched
    // the draft, and the draft stayed in the list. *** The application looked a receipt's system
    // id up as though it were a workbook path, found no draft, and so never retired one. This is
    // the search the application runs, with a receipt shaped as a real one is.
    // ###########################################################################################
    [Fact]
    public void A_published_receipt_naming_its_SYSTEM_finds_its_matching_draft()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        RetirableDraft found = Assert.Single(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, "Commodore/C64/250407", "published")],
            this.DataRoot,
            this.DraftsRoot,
            ["Commodore/C64/250425/Data C64 250425.xlsx", DraftRetirementTests.SystemKey]));

        Assert.Equal(DraftRetirementTests.SystemKey, found.ExcelDataFile);
    }

    // ###########################################################################################
    // *** A NEW SYSTEM'S DRAFT IS RETIRED ONCE ITS PUBLISHED BOARD IS HERE (owner request,
    // 2026-09-25): "People will either not know they can/should remove this or they forget, so
    // better clean-up when we can." ***
    //
    // Its draft is keyed by the name it was created with, and its publish writes the tree's
    // generation - so the two workbooks have DIFFERENT names, and the draft is found by its folder.
    // CRT downloads the published workbook only once the master lists the system, so it being here
    // means the contributor can open the published board instead.
    // ###########################################################################################
    private const string NewDraftKey = "Retro/Home Computer/Rev A/Data Home Computer Rev A.xlsx";
    private const string NewPublishedKey = "Retro/Home Computer/Rev A/Data Home Computer Rev A v2.0.0.xlsx";

    private void WriteNewSystem(BoardData draft, BoardData? published)
    {
        string draftPath = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftRetirementTests.NewDraftKey);
        Directory.CreateDirectory(Path.GetDirectoryName(draftPath)!);
        BoardWorkbookWriter.Write(draftPath, draft);
        BoardDataReader.ClearCache(draftPath);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftRetirementTests.NewDraftKey),
            new DraftMarker
            {
                SystemKey = DraftRetirementTests.NewDraftKey,
                BaseRevision = string.Empty,
                NewSystem = new NewSystemRegistration
                {
                    HardwareName = "Home Computer",
                    BoardName = "Rev A",
                    ExcelDataFile = DraftRetirementTests.NewDraftKey,
                },
                CreatedUtc = "2026-09-25T00:00:00Z",
            });

        if (published is not null)
        {
            string path = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftRetirementTests.NewPublishedKey);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            BoardWorkbookWriter.Write(path, published);
            BoardDataReader.ClearCache(path);
        }
    }

    [Fact]
    public void A_NEW_systems_draft_is_retired_once_its_published_board_is_listed_and_matches()
    {
        this.WriteNewSystem(DraftRetirementTests.BoardWith("CPU"), published: DraftRetirementTests.BoardWith("CPU"));

        RetirableDraft found = Assert.Single(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, "Retro/Home Computer/Rev A", "published")],
            this.DataRoot,
            this.DraftsRoot,
            [DraftRetirementTests.NewPublishedKey]));

        // Named by the DRAFT's own key - what the discard, the table guard and the lists use.
        Assert.Equal(DraftRetirementTests.NewDraftKey, found.ExcelDataFile);
    }

    // Changed after it was sent, or changed by a maintainer before publishing: the draft still holds
    // something the published board does not say.
    [Fact]
    public void A_NEW_systems_draft_that_differs_from_the_published_board_is_kept()
    {
        this.WriteNewSystem(DraftRetirementTests.BoardWith("CPU"), published: DraftRetirementTests.BoardWith("MPU"));

        Assert.Empty(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, "Retro/Home Computer/Rev A", "published")],
            this.DataRoot,
            this.DraftsRoot,
            [DraftRetirementTests.NewPublishedKey]));
    }

    // Published, but the master does not list it yet: CRT has not downloaded the published
    // workbook and cannot show the board, so the draft is the only way to see it. Kept.
    [Fact]
    public void A_NEW_system_the_master_does_not_list_yet_keeps_its_draft()
    {
        this.WriteNewSystem(DraftRetirementTests.BoardWith("CPU"), published: null);

        Assert.Empty(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, "Retro/Home Computer/Rev A", "published")],
            this.DataRoot,
            this.DraftsRoot,
            []));
    }

    // A receipt for a board the application does not know is left alone - there is no way to tell
    // which draft it means, and keeping a draft is always the safe mistake.
    [Fact]
    public void A_receipt_for_a_board_the_app_does_not_know_retires_nothing()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        Assert.Empty(DraftRetirement.FindRetirableDrafts(
            [DraftRetirementTests.Receipt(1, "Commodore/C64/250407", "published")],
            this.DataRoot,
            this.DraftsRoot,
            ["Commodore/C64/250425/Data C64 250425.xlsx"]));
    }

    // ------------------------------------------------------------------ FindRetirableSystems

    [Fact]
    public void Only_a_PUBLISHED_receipt_whose_draft_matches_is_named()
    {
        this.WritePublished(DraftRetirementTests.BoardWith("CPU"));
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));

        IReadOnlyList<string> retirable = DraftRetirement.FindRetirableSystems(
            [
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "published"),
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
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "approved"),
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
                DraftRetirementTests.Receipt(1, DraftRetirementTests.SystemId, "published"),
                DraftRetirementTests.Receipt(2, DraftRetirementTests.SystemId, "published"),
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

    // The lookup the application makes: a receipt's system id, among the boards it knows.
    private DraftStatus? ResolveStatus(string systemId) =>
        DraftStatusReader.ResolveForSystem(this.DataRoot, this.DraftsRoot, systemId, [DraftRetirementTests.SystemKey]);

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

    // What a discard that stopped after the workbook leaves: the marker, and no workbook beside it.
    private void WriteDraftThatLostItsWorkbook()
    {
        this.WriteDraft(DraftRetirementTests.BoardWith("CPU"));
        File.Delete(this.DraftWorkbook);
    }
}
