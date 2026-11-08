using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** A WORKBOOK, WORKLOG OR DEVICE FILE THAT CANNOT BE READ IS SAID ONCE (code review, 2026-11-07). *** Each is read again
// at every board switch, worklog action and device pick, and one damaged index.json logged an ERROR at every read - the
// "Something went wrong" card came back as soon as it was closed and counted up for ever. It is an error the first time the
// file is found so, a debug line while it stays as it was, and an error again once it changed (UnreadableFiles).
// ###########################################################################################
[Collection("Worklog")]
public sealed class UnreadableFilesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();
    private readonly string thisRoot;

    public UnreadableFilesTests()
    {
        this.thisRoot = this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N"));
        WorklogManager.LoadFrom(this.thisRoot);
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Detached-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private string Broken(string folder, string content = "{ not json")
    {
        string path = Path.Combine(this.thisRoot, folder, "index.json");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        return path;
    }

    [Fact]
    public void A_damaged_workbook_is_an_error_once_and_again_only_once_it_changed()
    {
        string path = this.Broken("workbook_7");
        using LoggedLines logged = LoggedLines.Watch();

        Assert.Empty(WorklogManager.GetAllWorkbooks());
        Assert.Empty(WorklogManager.GetAllWorkbooks());
        Assert.Empty(WorklogManager.GetAllWorkbooks());

        Assert.Equal(1, logged.Count(LogCategory.CRIT, "Failed to load workbook"));
        Assert.Equal(2, logged.Count(LogCategory.DBUG, "Failed to load workbook"));

        // Broken anew - another state of the file: said again.
        File.WriteAllText(path, "{ \"id\": 7, broken again");
        Assert.Empty(WorklogManager.GetAllWorkbooks());

        Assert.Equal(2, logged.Count(LogCategory.CRIT, "Failed to load workbook"));
    }

    [Fact]
    public void A_damaged_worklog_is_an_error_once()
    {
        WorkbookRecord workbook = WorklogManager.CreateWorkbook("Commodore 64|250407", "No video", string.Empty)!;
        this.Broken(Path.Combine($"workbook_{workbook.Id}", "worklog_3"));
        using LoggedLines logged = LoggedLines.Watch();

        Assert.Empty(WorklogManager.GetEntries(workbook.Id));
        Assert.Empty(WorklogManager.GetEntries(workbook.Id));

        Assert.Equal(1, logged.Count(LogCategory.CRIT, "Failed to load worklog"));
    }

    [Fact]
    public void A_damaged_device_is_an_error_once()
    {
        this.Broken("device_4");
        using LoggedLines logged = LoggedLines.Watch();

        Assert.Empty(OwnedDeviceStore.GetAll());
        Assert.Empty(OwnedDeviceStore.GetAll());

        Assert.Equal(1, logged.Count(LogCategory.CRIT, "Failed to read device"));
    }

    // ###########################################################################################
    // *** A BOARD FILE THAT CANNOT BE READ IS SAID ONCE TOO (code review, 2026-11-08). *** It is read again at every load
    // of the board and every look at its draft, and CRT.Data's reader logged its error each time - the card came back as
    // it was closed. Its warning and the reader's error are each said once, and again once the file changed.
    // ###########################################################################################
    [Fact]
    public void A_damaged_board_file_is_an_error_once_and_again_only_once_it_changed()
    {
        string path = Path.Combine(this.thisRoot, "Board-" + Guid.NewGuid().ToString("N"), DataTreeLayout.BoardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "{ not json");
        using LoggedLines logged = LoggedLines.Watch();

        Assert.Null(BoardDataReader.ReadWorkbookUncached(path));
        Assert.Null(BoardDataReader.ReadWorkbookUncached(path));
        Assert.Null(BoardDataReader.ReadWorkbookUncached(path));

        Assert.Equal(1, logged.Count(LogCategory.CRIT, "Failed to load board file"));
        Assert.Equal(1, logged.Count(LogCategory.WARN, "Could not read the board file"));
        Assert.Equal(2, logged.Count(LogCategory.DBUG, "Failed to load board file"));

        File.WriteAllText(path, "{ \"formatVersion\": broken again");
        Assert.Null(BoardDataReader.ReadWorkbookUncached(path));

        Assert.Equal(2, logged.Count(LogCategory.CRIT, "Failed to load board file"));
    }
}
