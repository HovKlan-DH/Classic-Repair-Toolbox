using System;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which workbook a board is read from (NewContributeStrategy.md Phase 6).
//
// *** THIS IS WHERE THE OVERLAY BECOMES A CHOICE. *** The published workbook used to be read and
// a draft's row deltas merged on top of it, so a drafted board was a computation rather than a
// file. Now a draft IS a board folder, and loading a drafted board means reading that folder's
// workbook instead. These tests pin which one wins, and - just as importantly - the cases where
// a draft must NOT win.
// ###########################################################################################
public sealed class DraftBoardSourceTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    // Writes a published workbook (contents irrelevant - this class resolves paths, it never parses).
    private void WritePublishedWorkbook()
    {
        string path = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "published workbook");
    }

    // Writes a draft folder: the marker, and optionally the workbook beside it.
    private void WriteDraft(bool withWorkbook = true, NewBoardRegistration? registration = null)
    {
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey),
            new DraftMarker
            {
                BoardKey = DraftBoardSourceTests.BoardKey,
                BaseRevision = registration is null ? "2026-09-01" : string.Empty,
                NewBoard = registration,
            });

        if (withWorkbook)
        {
            string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(path, "draft workbook");
        }
    }

    private BoardSourceSelection Resolve(bool preferPublished = false) =>
        DraftBoardSource.Resolve(
            this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished);

    // ------------------------------------------------------------------ The basic choice

    [Fact]
    public void With_NO_draft_the_published_workbook_is_read()
    {
        this.WritePublishedWorkbook();

        BoardSourceSelection selection = this.Resolve();

        Assert.False(selection.IsDraft);
        Assert.Null(selection.Marker);
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            selection.WorkbookPath);
    }

    [Fact]
    public void With_a_draft_the_DRAFTS_workbook_is_read_instead()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft();

        BoardSourceSelection selection = this.Resolve();

        Assert.True(selection.IsDraft);
        Assert.NotNull(selection.Marker);
        Assert.Equal(
            DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey),
            selection.WorkbookPath);
    }

    [Fact]
    public void The_published_path_is_reported_EVEN_WHEN_the_draft_is_the_one_being_read()
    {
        // So a caller can compare the draft against what it was drafted from (BoardDataDiffer)
        // without re-deriving the path, and so "view as published" has somewhere to point.
        this.WritePublishedWorkbook();
        this.WriteDraft();

        BoardSourceSelection selection = this.Resolve();

        Assert.True(selection.IsDraft);
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            selection.PublishedWorkbookPath);
    }

    // ------------------------------------------------------------------ What makes a draft a draft

    // ###########################################################################################
    // *** THE MARKER IS WHAT MAKES A FOLDER A DRAFT, NOT THE PRESENCE OF A WORKBOOK. ***
    //
    // A folder under Drafts/ holding an .xlsx but no marker is not a draft HERE. One a contributor
    // put there on purpose is given its marker when the board list loads (DraftFolderImport,
    // 2026-09-27) - at that one point, so this class keeps deciding from the marker alone and never
    // reads an unmarked folder in place of the published board.
    // ###########################################################################################
    [Fact]
    public void A_workbook_under_Drafts_with_NO_marker_is_NOT_treated_as_a_draft()
    {
        this.WritePublishedWorkbook();

        string strayWorkbook = DraftFolderLayout.GetWorkbookPath(
            this.DraftsRoot, DraftBoardSourceTests.BoardKey);

        Directory.CreateDirectory(Path.GetDirectoryName(strayWorkbook)!);
        File.WriteAllText(strayWorkbook, "a file somebody copied here");

        BoardSourceSelection selection = this.Resolve();

        Assert.False(selection.IsDraft);
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            selection.WorkbookPath);
    }

    // ###########################################################################################
    // A MARKER WITH NO WORKBOOK falls back to the published copy rather than refusing to load.
    //
    // Reachable by a crash between creating the folder and writing the workbook, or by someone
    // deleting the .xlsx by hand. Showing the published board is both true and recoverable;
    // refusing would leave a board that cannot be opened and a draft that cannot be discarded
    // from inside the application.
    // ###########################################################################################
    [Fact]
    public void A_marker_with_NO_workbook_falls_back_to_the_published_copy()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft(withWorkbook: false);

        BoardSourceSelection selection = this.Resolve();

        Assert.False(selection.IsDraft);
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            selection.WorkbookPath);

        // The marker is STILL reported, so the draft is still listed and can be discarded.
        Assert.NotNull(selection.Marker);
    }

    // ------------------------------------------------------------------ View as published

    // ###########################################################################################
    // "View boards as officially published" is honoured HERE, at the one place the choice is
    // made - the same single-point rule DataManager applied when the overlay still existed.
    // ###########################################################################################
    [Fact]
    public void VIEW_AS_PUBLISHED_reads_the_published_workbook_even_when_a_draft_exists()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft();

        BoardSourceSelection selection = this.Resolve(preferPublished: true);

        Assert.False(selection.IsDraft);
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            selection.WorkbookPath);
    }

    [Fact]
    public void VIEW_AS_PUBLISHED_still_reports_the_marker_so_the_draft_is_not_forgotten()
    {
        // The draft is still THERE and still listed; it is just not what is on screen. Hiding the
        // marker too would make the toggle look like it had discarded the draft.
        this.WritePublishedWorkbook();
        this.WriteDraft();

        Assert.NotNull(this.Resolve(preferPublished: true).Marker);
    }

    // ------------------------------------------------------------------ A draft-only board

    [Fact]
    public void A_DRAFT_ONLY_board_is_read_from_its_draft_and_says_so()
    {
        // No published workbook written at all - this board exists nowhere else.
        this.WriteDraft(registration: new NewBoardRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            ExcelDataFile = DraftBoardSourceTests.BoardKey,
        });

        BoardSourceSelection selection = this.Resolve();

        Assert.True(selection.IsDraft);
        Assert.True(selection.IsNewBoard);
    }

    // ###########################################################################################
    // *** IsNewBoard COMES OFF THE MARKER, NEVER FROM "the published file is missing". ***
    //
    // For a board the main workbook DOES list, a missing file is a real sync failure that must
    // keep failing loudly rather than quietly rendering an empty board. That distinction was
    // already load-bearing in DataManager before this change, and it survives it.
    // ###########################################################################################
    // ###########################################################################################
    // *** THE LISTING DECIDES THE BASELINE WHEN THE CALLER KNOWS IT (code review, 2026-09-27). ***
    // HardwareBoardEntry.IsPublished is the rule; the marker only stands in for it. A new board's
    // draft whose board has since been published is compared with the published board, and a
    // legacy board's draft with an ORDINARY marker ("Save to draft" seeds one) with nothing - not
    // with the contributor's own copy in Data/, against which nothing ever counted.
    // ###########################################################################################
    [Fact]
    public void A_new_boards_draft_of_a_board_now_listed_as_published_is_compared_with_the_published_board()
    {
        this.WriteDraft(registration: new NewBoardRegistration { ExcelDataFile = DraftBoardSourceTests.BoardKey });

        BoardSourceSelection selection = DraftBoardSource.Resolve(
            this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, listedAsPublished: true);

        Assert.Equal(DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey), selection.PublishedWorkbookPath);
        Assert.Equal(
            selection.PublishedWorkbookPath,
            DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, listedAsPublished: true)!.PublishedWorkbookPath);
    }

    [Fact]
    public void A_draft_of_a_board_not_listed_as_published_is_compared_with_nothing_whatever_its_marker()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft();

        Assert.Equal(string.Empty, DraftBoardSource.Resolve(
            this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, listedAsPublished: false).PublishedWorkbookPath);
        Assert.Equal(string.Empty, DraftStatusReader.Resolve(
            this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, listedAsPublished: false)!.PublishedWorkbookPath);

        // With no listing to ask, the marker decides as before.
        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            this.Resolve().PublishedWorkbookPath);
    }

    [Fact]
    public void A_PUBLISHED_board_whose_file_is_missing_is_NOT_mistaken_for_a_new_board()
    {
        // A draft with no registration, and no published file on disk - a broken sync, not a new
        // board.
        this.WriteDraft();

        Assert.False(this.Resolve().IsNewBoard);
    }

    [Fact]
    public void VIEW_AS_PUBLISHED_on_a_draft_only_board_points_at_a_file_that_does_not_exist()
    {
        // Correct rather than a bug: officially this board does not exist, and the load path
        // renders that as a blank board rather than as a failure.
        this.WriteDraft(registration: new NewBoardRegistration
        {
            ExcelDataFile = DraftBoardSourceTests.BoardKey,
        });

        BoardSourceSelection selection = this.Resolve(preferPublished: true);

        Assert.False(selection.IsDraft);
        Assert.False(File.Exists(selection.WorkbookPath));
        Assert.True(selection.IsNewBoard);
    }

    // ###########################################################################################
    // *** A NEW BOARD IS COMPARED AGAINST NOTHING, even with a file at its published path
    // (2026-09-27). *** That file can only be the contributor's own - a board made the old way,
    // listed by a _UserContribution workbook, lives in Data/ at exactly this path - and comparing a
    // copy of it against itself marked none of its rows as drafted. The ordinary draft beside it
    // (The_published_path_is_reported_EVEN_WHEN...) keeps its published baseline.
    // ###########################################################################################
    [Fact]
    public void A_DRAFT_ONLY_board_has_NO_baseline_even_with_a_file_at_its_published_path()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft(registration: new NewBoardRegistration
        {
            HardwareName = "C64",
            BoardName = "250407",
            ExcelDataFile = DraftBoardSourceTests.BoardKey,
        });

        BoardSourceSelection selection = this.Resolve();

        Assert.True(selection.IsDraft);
        Assert.Equal(string.Empty, selection.PublishedWorkbookPath);
    }

    // ------------------------------------------------------------------ HasDraft

    [Fact]
    public void HasDraft_answers_from_the_MARKER_alone()
    {
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftBoardSourceTests.BoardKey));

        this.WriteDraft();

        Assert.True(DraftBoardSource.HasDraft(this.DraftsRoot, DraftBoardSourceTests.BoardKey));
    }

    [Fact]
    public void HasDraft_is_FALSE_for_a_stray_workbook_with_no_marker()
    {
        // Same rule as the load itself, and it has to be: a Drafts tab that listed stray folders
        // would offer to submit data the contributor never drafted.
        string strayWorkbook = DraftFolderLayout.GetWorkbookPath(
            this.DraftsRoot, DraftBoardSourceTests.BoardKey);

        Directory.CreateDirectory(Path.GetDirectoryName(strayWorkbook)!);
        File.WriteAllText(strayWorkbook, "stray");

        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, DraftBoardSourceTests.BoardKey));
    }

    // ------------------------------------------------------------------ ViewedDraftFolder / ResolveViewedPath

    // ###########################################################################################
    // *** "VIEW AS PUBLISHED" MUST SWITCH THE FILES AND THE SIDECAR, NOT ONLY THE WORKBOOK (code
    // review, 2026-09-25). ***
    //
    // The toggle switched which workbook was read, while images, local files, KiCad data and
    // calibrations kept being looked up draft-first - so the screen claimed to be the published
    // board while drawing the contributor's own unpublished bytes. These two rules are what every
    // display surface now asks, so they are pinned for all three states.
    // ###########################################################################################
    [Fact]
    public void The_viewed_draft_folder_is_the_draft_while_the_draft_is_shown()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft();

        Assert.Equal(
            DraftFolderLayout.GetBoardFolder(this.DraftsRoot, DraftBoardSourceTests.BoardKey),
            DraftBoardSource.ViewedDraftFolder(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished: false));

        Assert.Equal(
            DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey),
            DraftBoardSource.ResolveViewedPath(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished: false));
    }

    [Fact]
    public void VIEW_AS_PUBLISHED_gives_NO_draft_folder_and_the_PUBLISHED_sidecar()
    {
        this.WritePublishedWorkbook();
        this.WriteDraft();

        Assert.Equal(
            string.Empty,
            DraftBoardSource.ViewedDraftFolder(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished: true));

        Assert.Equal(
            DraftBoardSource.PublishedPathOf(this.DataRoot, DraftBoardSourceTests.BoardKey),
            DraftBoardSource.ResolveViewedPath(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished: true));

        // The WRITE path is unaffected: an edit made while viewing the published board still
        // belongs in the draft, which is the only place a contributor's change may land.
        Assert.Equal(
            DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey),
            DraftBoardSource.ResolveWritablePath(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey));
    }

    [Fact]
    public void A_stray_folder_with_no_marker_gives_no_draft_folder_to_search()
    {
        // Same rule as the workbook: a folder this application did not create is not a draft, so
        // its files must not replace the published ones on screen either.
        this.WritePublishedWorkbook();

        string strayWorkbook = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftBoardSourceTests.BoardKey);
        Directory.CreateDirectory(Path.GetDirectoryName(strayWorkbook)!);
        File.WriteAllText(strayWorkbook, "stray");

        Assert.Equal(
            string.Empty,
            DraftBoardSource.ViewedDraftFolder(this.DataRoot, this.DraftsRoot, DraftBoardSourceTests.BoardKey, preferPublished: false));
    }

    [Fact]
    public void An_unresolvable_board_answers_empty_rather_than_throwing()
    {
        BoardSourceSelection selection = DraftBoardSource.Resolve(
            this.DataRoot, this.DraftsRoot, excelDataFile: string.Empty);

        Assert.Equal(string.Empty, selection.WorkbookPath);
        Assert.False(selection.IsDraft);
        Assert.False(DraftBoardSource.HasDraft(this.DraftsRoot, string.Empty));
    }
}
