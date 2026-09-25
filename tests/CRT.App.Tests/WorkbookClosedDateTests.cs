using Avalonia;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The workbook's endDate: when the job actually FINISHED, so the top bar above the tabs can report a
// closed workbook by its finish date instead of by the date it was started.
//
// The rule has two halves and both matter. It is stamped on the TRANSITION into Closed, so ordinary
// edits to an already-closed workbook (attaching a photo to a closed worklog, say) do not keep
// pushing the finish date forward to today. And startDate is never written by any of it - a job's
// start is a fact about the past.
//
// Touches WorklogManager's static root, so this joins the "Worklog" collection.
[Collection("Worklog")]
public sealed class WorkbookClosedDateTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // Detach from the temp folder so nothing written later can reach the user's real one.
        this.LoadWorklog();
        this.thisWorkspace.Dispose();
    }

    private string LoadWorklog()
    {
        string root = this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(root);
        return root;
    }

    private const string BoardKey = "Commodore 64|250469";

    private static WorkbookRecord ReadBack(int workbookId) =>
        WorklogManager.GetWorkbooksForBoard(BoardKey).Single(w => w.Id == workbookId);

    private static WorklogEntryRecord AddEntry(int workbookId, string state)
    {
        var entry = WorklogManager.AddEntry(
            workbookId, "Motherboard", new Rect(0, 0, 10, 10), "Bad cap", "", "Issue", state, Array.Empty<string>());

        Assert.NotNull(entry);
        return entry!;
    }

    // A brand-new workbook has never been closed, so it carries no finish date at all - the top bar
    // falls back to the start date for it, which is the right thing for a job in progress.
    [Fact]
    public void A_new_workbook_has_no_end_date()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        Assert.NotNull(workbook);

        Assert.Null(ReadBack(workbook!.Id).EndDate);
    }

    // An open worklog keeps the workbook open, and an open workbook must have no finish date - it has
    // not finished.
    [Fact]
    public void A_workbook_with_an_open_worklog_has_no_end_date()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        AddEntry(workbook!.Id, "Open");

        var reloaded = ReadBack(workbook.Id);
        Assert.False(WorklogManager.IsWorkbookStatusOpen(reloaded.Status) is false);
        Assert.Null(reloaded.EndDate);
    }

    // The transition itself: closing the last outstanding worklog closes the workbook and stamps
    // today as the finish date.
    [Fact]
    public void Closing_the_last_open_worklog_stamps_the_end_date()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        var entry = AddEntry(workbook!.Id, "Open");

        Assert.Null(ReadBack(workbook.Id).EndDate);

        entry.State = "Closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        var closed = ReadBack(workbook.Id);
        Assert.False(WorklogManager.IsWorkbookStatusOpen(closed.Status));
        Assert.Equal(DateTime.Now.Date, closed.EndDate);
    }

    // THE POINT OF "ON THE TRANSITION ONLY". Editing an already-closed workbook recomputes its status
    // (it stays Closed) and must NOT restamp the finish date - otherwise every later touch of a
    // finished job silently rewrites when it was finished, and a workbook closed last year would
    // report today the moment someone opened one of its worklogs to read it.
    //
    // The date is forced to a past value on disk first, because a same-day restamp is invisible: a
    // test that closed and immediately re-saved would write today's date twice and pass either way.
    [Fact]
    public void Editing_an_already_closed_workbook_does_not_move_its_end_date()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        var entry = AddEntry(workbook!.Id, "Open");

        entry.State = "Closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        // Backdate the stamped finish date directly in index.json, then reload so the manager reads
        // it back - this is the "closed a while ago" state a same-session test cannot otherwise
        // produce.
        string indexPath = Path.Combine(root, $"workbook_{workbook.Id}", "index.json");
        var stored = System.Text.Json.JsonSerializer.Deserialize<WorkbookRecord>(File.ReadAllText(indexPath))!;
        var closedOn = new DateTime(2020, 3, 4);
        stored.EndDate = closedOn;
        File.WriteAllText(indexPath, System.Text.Json.JsonSerializer.Serialize(stored));

        Assert.Equal(closedOn, ReadBack(workbook.Id).EndDate);

        // An ordinary edit to the closed worklog - the workbook is recomputed and stays Closed.
        entry.Description = "Also replaced the socket";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        var after = ReadBack(workbook.Id);
        Assert.False(WorklogManager.IsWorkbookStatusOpen(after.Status));
        Assert.Equal(closedOn, after.EndDate);
    }

    // Reopening and re-closing IS a new transition, so the finish date moves to when the job actually
    // finished the second time. The stale value left behind while it was reopened is harmless
    // (nothing reads it while the workbook is Open) but must not survive the re-close.
    [Fact]
    public void Reopening_and_closing_again_stamps_the_new_finish_date()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        var entry = AddEntry(workbook!.Id, "Closed");

        string indexPath = Path.Combine(root, $"workbook_{workbook.Id}", "index.json");
        var stored = System.Text.Json.JsonSerializer.Deserialize<WorkbookRecord>(File.ReadAllText(indexPath))!;
        stored.EndDate = new DateTime(2020, 3, 4);
        File.WriteAllText(indexPath, System.Text.Json.JsonSerializer.Serialize(stored));

        // Reopen: the workbook goes back to Open, and the old finish date is now meaningless.
        entry.State = "Open";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));
        Assert.True(WorklogManager.IsWorkbookStatusOpen(ReadBack(workbook.Id).Status));

        // Close it again - a real transition, so today's date replaces the stale one.
        entry.State = "Closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.Equal(DateTime.Now.Date, ReadBack(workbook.Id).EndDate);
    }

    // The start date is a fact about the past and nothing in this feature may touch it - the top bar
    // still shows it for every open workbook, and falls back to it for a workbook closed before
    // endDate existed.
    [Fact]
    public void Closing_a_workbook_never_changes_its_start_date()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        var startDate = workbook!.StartDate;

        var entry = AddEntry(workbook.Id, "Open");
        entry.State = "Closed";
        Assert.True(WorklogManager.UpdateEntry(workbook.Id, entry));

        Assert.Equal(startDate, ReadBack(workbook.Id).StartDate);
    }

    // Every workbook written before endDate existed deserializes to null, INCLUDING ones already
    // Closed - there is deliberately no migration. Readers must handle that (the top bar falls back
    // to the start date) rather than seeing default(DateTime) and printing "0001-January-01".
    [Fact]
    public void A_workbook_written_before_the_field_existed_reads_back_with_no_end_date()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");

        // index.json as an older build wrote it: a closed workbook with no endDate key at all.
        string indexPath = Path.Combine(root, $"workbook_{workbook!.Id}", "index.json");
        File.WriteAllText(indexPath, """
        {
          "id": 1,
          "boardKey": "Commodore 64|250469",
          "title": "C64 job",
          "note": "",
          "status": "Closed",
          "startDate": "2024-01-15T00:00:00",
          "entryCount": 2
        }
        """);

        var reloaded = ReadBack(workbook.Id);
        Assert.Null(reloaded.EndDate);
        Assert.Equal(new DateTime(2024, 1, 15), reloaded.StartDate);
    }
}
