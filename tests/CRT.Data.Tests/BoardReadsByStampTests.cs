using System;
using System.IO;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardReadsByStamp - one parse of a workbook per version of it, shared by the Drafts tab's counts
// (code review, 2026-10-13).
//
// Shares "BoardData": the workbooks are read through BoardDataReader.
// ###########################################################################################
[Collection("BoardData")]
public sealed class BoardReadsByStampTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private string WriteWorkbook(string name)
    {
        string path = Path.Combine(this.thisWorkspace.Root, name, "Data C64 250407.xlsx");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        BoardWorkbookWriter.Write(path, new BoardData
        {
            RevisionDate = "2026-09-01",
            Components = [new ComponentEntry { BoardLabel = "U8" }],
        });

        return path;
    }

    [Fact]
    public void A_workbook_read_once_is_remembered()
    {
        string path = this.WriteWorkbook("draft");

        Assert.NotNull(BoardReadsByStamp.Read(path));
        Assert.True(BoardReadsByStamp.HoldsForTests(path));
    }

    // ###########################################################################################
    // *** A WORKBOOK THAT IS GONE IS LET GO (code review, 2026-10-14). *** A discarded or retired
    // draft's parsed board - tens of megabytes for a large board - stayed in the static cache for
    // the rest of the session, one more for every draft made and discarded. The next read of ANY
    // workbook drops the entries whose file no longer exists.
    // ###########################################################################################
    [Fact]
    public void A_workbook_deleted_since_it_was_read_is_dropped_at_the_next_read_of_any_workbook()
    {
        string discarded = this.WriteWorkbook("discarded");
        string kept = this.WriteWorkbook("kept");

        Assert.NotNull(BoardReadsByStamp.Read(discarded));
        Assert.NotNull(BoardReadsByStamp.Read(kept));

        Directory.Delete(Path.GetDirectoryName(discarded)!, recursive: true);

        Assert.NotNull(BoardReadsByStamp.Read(kept));

        Assert.False(BoardReadsByStamp.HoldsForTests(discarded));
        Assert.True(BoardReadsByStamp.HoldsForTests(kept));
    }

    // Asking for the deleted workbook itself lets it go too, and answers nothing.
    [Fact]
    public void Reading_a_deleted_workbook_answers_nothing_and_drops_it()
    {
        string path = this.WriteWorkbook("draft");
        Assert.NotNull(BoardReadsByStamp.Read(path));

        File.Delete(path);

        Assert.Null(BoardReadsByStamp.Read(path));
        Assert.False(BoardReadsByStamp.HoldsForTests(path));
    }
}
