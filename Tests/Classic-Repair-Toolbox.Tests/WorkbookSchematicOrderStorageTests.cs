using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Persisting the Workbooks board pane's drag-to-reorder result in the workbook's own index.json.
// The ordering RULE itself is WorkbookSchematicOrderTests; this covers only that the chosen order
// survives a round trip and that saving it does not disturb anything else in the record.
//
// Touches WorklogManager's static root, so this joins the "Worklog" collection.
[Collection("Worklog")]
public sealed class WorkbookSchematicOrderStorageTests : IDisposable
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

    // A workbook nobody has rearranged carries an empty order, which the pane treats as "keep the
    // alphabetical grouping" - so this feature is invisible until it is used, with no migration.
    [Fact]
    public void A_new_workbook_has_no_schematic_order()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        Assert.NotNull(workbook);

        Assert.Empty(ReadBack(workbook!.Id).SchematicOrder);
    }

    [Fact]
    public void A_saved_schematic_order_survives_a_reload()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        var order = new List<string> { "Video", "Motherboard", "Power supply" };

        Assert.True(WorklogManager.UpdateWorkbookSchematicOrder(workbook!.Id, order));

        // Through a real reload, not just the in-memory record: the point of this method is that the
        // order is on DISK, so the pane comes back in the user's arrangement next session.
        WorklogManager.LoadFrom(root);

        Assert.Equal(order, ReadBack(workbook.Id).SchematicOrder);
    }

    // The order is saved on its own, separately from the "Edit workbook" dialog's title/note save.
    // It must therefore leave every other field exactly as it found it - a reorder is not an edit of
    // the job itself.
    [Fact]
    public void Saving_the_order_leaves_the_rest_of_the_workbook_alone()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "Mr. Jensens C64", "Dropped off Tuesday");

        Assert.True(WorklogManager.UpdateWorkbookSchematicOrder(workbook!.Id, new List<string> { "Video" }));

        var reloaded = ReadBack(workbook.Id);
        Assert.Equal("Mr. Jensens C64", reloaded.Title);
        Assert.Equal("Dropped off Tuesday", reloaded.Note);
        Assert.Equal(workbook.Status, reloaded.Status);
        Assert.Equal(workbook.StartDate, reloaded.StartDate);
        Assert.Equal(workbook.WorklogCount, reloaded.WorklogCount);
    }

    [Fact]
    public void Re_saving_the_order_replaces_the_previous_one_rather_than_appending()
    {
        this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");

        WorklogManager.UpdateWorkbookSchematicOrder(workbook!.Id, new List<string> { "A", "B", "C" });
        WorklogManager.UpdateWorkbookSchematicOrder(workbook.Id, new List<string> { "C", "A" });

        Assert.Equal(new[] { "C", "A" }, ReadBack(workbook.Id).SchematicOrder);
    }

    // The order lives inside the workbook's own folder, so deleting the workbook takes it with it -
    // there is no separate bookkeeping anywhere that could be left dangling. This is the whole reason
    // it is stored per workbook rather than in UserSettings.
    [Fact]
    public void Deleting_the_workbook_takes_its_order_with_it()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");
        WorklogManager.UpdateWorkbookSchematicOrder(workbook!.Id, new List<string> { "Video" });

        Assert.True(WorklogManager.DeleteWorkbook(workbook.Id));

        Assert.False(Directory.Exists(Path.Combine(root, $"workbook_{workbook.Id}")));
        Assert.Empty(WorklogManager.GetWorkbooksForBoard(BoardKey));
    }

    // A workbook that no longer exists must be reported rather than throwing - a refresh can rebuild
    // the pane mid-drag, and the drop then commits against a workbook another window has deleted.
    [Fact]
    public void Saving_the_order_for_a_missing_workbook_reports_failure()
    {
        this.LoadWorklog();

        Assert.False(WorklogManager.UpdateWorkbookSchematicOrder(999, new List<string> { "Video" }));
    }

    // Every workbook written before this field existed has no schematicOrder key, and must
    // deserialize to an empty list rather than null - the pane and WorkbookSchematicOrder.Apply both
    // read it directly.
    [Fact]
    public void A_workbook_written_before_the_field_existed_reads_back_with_an_empty_order()
    {
        string root = this.LoadWorklog();

        var workbook = WorklogManager.CreateWorkbook(BoardKey, "C64 job", "");

        string indexPath = Path.Combine(root, $"workbook_{workbook!.Id}", "index.json");
        File.WriteAllText(indexPath, """
        {
          "id": 1,
          "boardKey": "Commodore 64|250469",
          "title": "C64 job",
          "note": "",
          "status": "Open",
          "startDate": "2024-01-15T00:00:00",
          "entryCount": 0
        }
        """);

        Assert.NotNull(ReadBack(workbook.Id).SchematicOrder);
        Assert.Empty(ReadBack(workbook.Id).SchematicOrder);
    }
}
