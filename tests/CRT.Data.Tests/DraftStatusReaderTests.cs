using System;
using System.Collections.Generic;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DraftStatusReader.CountChangesCached - the "N rows changed" number the Drafts tab shows for
// every drafted system (code review, 2026-09-25).
//
// *** WHY IT IS CACHED. *** The tab re-lists every drafted system after every save of any kind,
// and each count is two full workbook parses on the UI thread. With a few large boards drafted,
// each save froze the window for seconds re-counting boards nobody had touched.
//
// *** WHY THE CACHE CANNOT GO STALE. *** CountChanges was uncached on purpose: a contributor may
// edit the draft in Excel. So the cache is keyed by each input file's length and write time, and
// these tests pin both halves - an untouched board is answered without a parse, and any change to
// any of the four inputs (both workbooks, both sidecars) is counted afresh.
//
// "BoardData" collection: BoardDataReader's static cache is keyed by path, and these tests write
// workbooks.
// ###########################################################################################
[Collection("BoardData")]
public sealed class DraftStatusReaderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private const string SystemKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    public void Dispose() => this.thisWorkspace.Dispose();

    [Fact]
    public void The_cached_count_matches_a_fresh_count()
    {
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(3));

        DraftStatus? status = this.Status();

        Assert.Equal(2, DraftStatusReader.CountChanges(status));
        Assert.Equal(2, DraftStatusReader.CountChangesCached(status));
    }

    // ###########################################################################################
    // *** THE PROOF THAT AN UNTOUCHED BOARD IS NOT RE-PARSED. ***
    //
    // The workbook's bytes are replaced with garbage of the SAME length and its write time is put
    // back, so the stamp is unchanged. A fresh count now fails to parse and answers 0; the cached
    // count still answers the real number - which it can only do without reading the file.
    // ###########################################################################################
    [Fact]
    public void An_untouched_draft_is_answered_from_the_cache_without_a_parse()
    {
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(3));

        DraftStatus? status = this.Status();
        Assert.Equal(2, DraftStatusReader.CountChangesCached(status));

        string workbook = status!.WorkbookPath;
        long length = new FileInfo(workbook).Length;
        DateTime written = File.GetLastWriteTimeUtc(workbook);

        File.WriteAllBytes(workbook, new byte[length]);
        File.SetLastWriteTimeUtc(workbook, written);

        Assert.Equal(0, DraftStatusReader.CountChanges(status));
        Assert.Equal(2, DraftStatusReader.CountChangesCached(status));
    }

    // ###########################################################################################
    // *** A DRAFT'S ERRORS AND WARNINGS, FOR ITS ROW (owner report, 2026-10-02: "It must check for
    // errors when creating the draft, and if the board changes "offline", outside of app"). *** The
    // table's own checks, counted without opening the table - and counted again after an edit in
    // Excel, like the change count.
    // ###########################################################################################
    [Fact]
    public void A_drafts_errors_and_warnings_are_counted_and_follow_an_edit_in_Excel()
    {
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));

        // Every component is unmarked on any schematic: one warning each.
        this.WriteDraft(DraftStatusReaderTests.BoardWith(2));

        DraftStatus? status = this.Status();
        string draftFolder = DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftStatusReaderTests.SystemKey);

        Assert.Equal(new BoardProblemCounts(0, 2), DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder));

        // In Excel: a link that is not a web address - an error.
        BoardData edited = DraftStatusReaderTests.BoardWith(2);
        edited.ComponentLinks.Add(new ComponentLinkEntry { BoardLabel = "U0", Name = "Bad", Url = "ftp://example.com" });
        this.WriteDraft(edited);
        File.SetLastWriteTimeUtc(status!.WorkbookPath, DateTime.UtcNow.AddMinutes(1));

        Assert.Equal(new BoardProblemCounts(1, 2), DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder));
    }

    // ###########################################################################################
    // Case 18 (owner request, 2026-10-03: "Flagged" became warnings; cases agreed with the project
    // owner). A draft's row counts its duplicate rows as warnings, one per row. A row the save
    // leaves out (an important signal missing its net) is never in the FILE as such - reading the
    // workbook drops it, exactly as saving does - so only an open table can count it.
    // ###########################################################################################
    [Fact]
    public void A_drafts_row_counts_its_duplicate_rows_as_warnings_but_not_a_row_the_save_leaves_out()
    {
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));

        // One component, unmarked on any schematic: one warning.
        BoardData draft = DraftStatusReaderTests.BoardWith(1);

        // Its pinout twice - note rows, so no file is looked for.
        draft.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U0", Name = "Pinout", Note = "Same as 74LS08" });
        draft.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U0", Name = "Pinout", Note = "Same as 7408" });
        draft.KiCadImportantSignals.Add(new KiCadImportantSignalEntry { DisplayName = "CLK", KiCadNetName = string.Empty });

        this.WriteDraft(draft);

        DraftStatus? status = this.Status();
        string draftFolder = DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftStatusReaderTests.SystemKey);

        Assert.Equal(new BoardProblemCounts(0, 3), DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder));
    }

    // The same answer the table gives, and remembered: a draft nobody touched is not read again.
    [Fact]
    public void An_untouched_drafts_problems_are_answered_from_the_cache()
    {
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(3));

        DraftStatus? status = this.Status();
        string draftFolder = DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftStatusReaderTests.SystemKey);
        Assert.Equal(new BoardProblemCounts(0, 3), DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder));

        string workbook = status!.WorkbookPath;
        long length = new FileInfo(workbook).Length;
        DateTime written = File.GetLastWriteTimeUtc(workbook);

        File.WriteAllBytes(workbook, new byte[length]);
        File.SetLastWriteTimeUtc(workbook, written);

        Assert.Equal(new BoardProblemCounts(0, 3), DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder));
    }

    // ###########################################################################################
    // A picture the draft cites arriving on its own - the background image sync finishing, or a file
    // dropped into the draft folder by hand - changes neither stamped file. The row kept "1 error"
    // for the session while the table and Submit found the file (code review, 2026-10-04); a path
    // the check did not find is now looked for again.
    // ###########################################################################################
    [Fact]
    public void A_cited_file_arriving_later_takes_its_error_off_the_drafts_row()
    {
        const string Picture = "Commodore/C64/250407/Images/U0 pinout.png";

        this.WritePublished(DraftStatusReaderTests.BoardWith(1));

        BoardData draft = DraftStatusReaderTests.BoardWith(1);
        draft.ComponentImages.Add(new ComponentImageEntry { BoardLabel = "U0", Name = "Pinout", File = Picture });
        this.WriteDraft(draft);

        DraftStatus? status = this.Status();
        string draftFolder = DraftFolderLayout.GetSystemFolder(this.DraftsRoot, DraftStatusReaderTests.SystemKey);

        BoardProblemCounts before = DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder);
        Assert.True(before.Errors >= 1);

        string arrived = Path.Combine(this.DataRoot, Picture.Replace('/', Path.DirectorySeparatorChar));
        Directory.CreateDirectory(Path.GetDirectoryName(arrived)!);
        File.WriteAllBytes(arrived, [0x89, 0x50, 0x4E, 0x47]);

        Assert.Equal(before.Errors - 1, DraftStatusReader.CountProblemsCached(status, this.DataRoot, draftFolder).Errors);
    }

    [Fact]
    public void A_draft_edited_since_the_last_count_is_counted_afresh()
    {
        // The Excel case: the workbook changes on disk and the tab must not keep showing the old
        // number.
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(3));

        DraftStatus? status = this.Status();
        Assert.Equal(2, DraftStatusReader.CountChangesCached(status));

        this.WriteDraft(DraftStatusReaderTests.BoardWith(5));
        File.SetLastWriteTimeUtc(status!.WorkbookPath, DateTime.UtcNow.AddMinutes(1));

        Assert.Equal(4, DraftStatusReader.CountChangesCached(status));
    }

    [Fact]
    public void A_new_published_board_is_counted_afresh()
    {
        // The sync bringing down a new published board changes the answer without the draft
        // moving at all, so the published side is part of the stamp too.
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(3));

        DraftStatus? status = this.Status();
        Assert.Equal(2, DraftStatusReader.CountChangesCached(status));

        this.WritePublished(DraftStatusReaderTests.BoardWith(3));
        File.SetLastWriteTimeUtc(status!.PublishedWorkbookPath, DateTime.UtcNow.AddMinutes(1));

        Assert.Equal(0, DraftStatusReader.CountChangesCached(status));
    }

    [Fact]
    public void A_changed_SIDECAR_is_counted_afresh()
    {
        // Highlights live in the sidecar, not the workbook, and the count includes them - so a
        // label-editor save that touches only the sidecar must still move the number.
        this.WritePublished(DraftStatusReaderTests.BoardWith(1));
        this.WriteDraft(DraftStatusReaderTests.BoardWith(1));

        DraftStatus? status = this.Status();
        Assert.Equal(0, DraftStatusReader.CountChangesCached(status));

        BoardSidecarWriter.Write(
            status!.WorkbookPath,
            new List<ComponentHighlightEntry>
            {
                new() { SchematicName = "Top", BoardLabel = "U0", X = "1", Y = "2", Width = "3", Height = "4" },
            },
            new List<KiCadCalibrationEntry>());

        Assert.Equal(1, DraftStatusReader.CountChangesCached(status));
    }

    [Fact]
    public void No_status_counts_zero()
    {
        Assert.Equal(0, DraftStatusReader.CountChangesCached(null));
    }

    // ------------------------------------------------------------------ helpers

    private DraftStatus? Status() =>
        DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftStatusReaderTests.SystemKey);

    private static BoardData BoardWith(int components)
    {
        var board = new BoardData { RevisionDate = "2026-09-01" };

        for (int i = 0; i < components; i++)
        {
            board.Components.Add(new ComponentEntry { BoardLabel = $"U{i}", Description = "part" });
        }

        return board;
    }

    private void WritePublished(BoardData board)
    {
        string path = DraftBoardSource.PublishedPathOf(this.DataRoot, DraftStatusReaderTests.SystemKey);

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
        BoardDataReader.ClearCache(path);
    }

    private void WriteDraft(BoardData board)
    {
        string path = DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, DraftStatusReaderTests.SystemKey);

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        BoardWorkbookWriter.Write(path, board);
        BoardDataReader.ClearCache(path);

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftStatusReaderTests.SystemKey),
            new DraftMarker
            {
                SystemKey = DraftStatusReaderTests.SystemKey,
                BaseRevision = "2026-09-01",
                CreatedUtc = "2026-09-25T00:00:00Z",
            });
    }
    // ###########################################################################################
    // ResolveForSystem - a submission receipt's system id back to its draft (2026-09-25).
    // ###########################################################################################

    // Resolve takes a WORKBOOK path. Handed a system id it looks one folder too high and finds
    // nothing - which is exactly what the application did, so no published draft was ever retired.
    [Fact]
    public void Resolve_handed_a_system_id_instead_of_a_workbook_finds_nothing()
    {
        this.WriteDraftMarker();

        Assert.NotNull(DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, DraftStatusReaderTests.SystemKey));
        Assert.Null(DraftStatusReader.Resolve(this.DataRoot, this.DraftsRoot, "Commodore/C64/250407"));
    }

    // The id is built from the workbook path when a draft is submitted; the lookup goes back the
    // same way, so the two cannot disagree about what a system is called.
    [Fact]
    public void A_receipts_system_id_finds_the_draft_of_the_board_it_was_built_from()
    {
        this.WriteDraftMarker();
        string systemId = SystemDescriptorRules.SystemIdFromExcelDataFile(DraftStatusReaderTests.SystemKey);

        DraftStatus? status = DraftStatusReader.ResolveForSystem(
            this.DataRoot, this.DraftsRoot, systemId,
            ["Amstrad/CPC 664/MC0005A/Data CPC 664 MC0005A.xlsx", DraftStatusReaderTests.SystemKey]);

        Assert.NotNull(status);
        Assert.Equal(DraftStatusReaderTests.SystemKey, status!.SystemKey);
    }

    [Fact]
    public void A_system_id_no_known_board_has_resolves_to_nothing()
    {
        this.WriteDraftMarker();

        Assert.Null(DraftStatusReader.ResolveForSystem(this.DataRoot, this.DraftsRoot, "Commodore/C64/250407", []));
        Assert.Null(DraftStatusReader.ResolveForSystem(this.DataRoot, this.DraftsRoot, "Commodore/C64/250407", null));
        Assert.Null(DraftStatusReader.ResolveForSystem(this.DataRoot, this.DraftsRoot, "  ", [DraftStatusReaderTests.SystemKey]));
    }

    private void WriteDraftMarker()
    {
        string marker = DraftFolderLayout.GetMarkerPath(this.DraftsRoot, DraftStatusReaderTests.SystemKey);
        Directory.CreateDirectory(Path.GetDirectoryName(marker)!);
        DraftMarkerStore.Save(marker, new DraftMarker { BaseRevision = "2026-September-25" });
    }
}
