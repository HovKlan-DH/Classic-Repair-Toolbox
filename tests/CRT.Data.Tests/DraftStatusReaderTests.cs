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
}
