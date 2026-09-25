using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The component editor's "File location" drop-down on a system that exists ONLY as a draft.
//
// *** REPORTED 2026-09-24. *** A system created with "Add a new system" gets its own empty
// "Scope baseline" folder (DraftSeeder.CreateNewSystem), but the editor's drop-down was built from
// the data root alone - so that folder was never offered, and a baseline image could not be filed
// where published boards keep theirs.
//
// End to end on purpose: the system is created by the real seeder, and the list is read off a real
// file row the window built, which is exactly what the drop-down's ItemsSource binds to. If the
// seeder and the window ever disagree about where that folder is, this fails.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class ComponentContributionFileLocationTests : IDisposable
{
    private const string Manufacturer = "Test Manu4";
    private const string Hardware = "HW4";
    private const string Board = "Board4";

    private readonly TempWorkspace thisWorkspace = new();

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    public ComponentContributionFileLocationTests()
    {
        // DraftManager is a process-wide static the window reads directly, so it is pointed at
        // the workspace before anything is built - and restored in Dispose, so this class does
        // not leak its root into the next one in the collection.
        ComponentContributionFileLocationTests.PointDraftManagerAt(this.DraftsRoot);
    }

    public void Dispose()
    {
        ComponentContributionFileLocationTests.PointDraftManagerAt(string.Empty);
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void A_NEW_systems_Scope_baseline_folder_is_offered_in_a_file_rows_drop_down()
    {
        // A published board, so the data tree is not empty and the test proves the draft folder
        // is ADDED to it rather than replacing it.
        Directory.CreateDirectory(Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Scope baseline"));

        string excelDataFile = NewSystemIdentity.BuildExcelDataFile(
            ComponentContributionFileLocationTests.Manufacturer,
            ComponentContributionFileLocationTests.Hardware,
            ComponentContributionFileLocationTests.Board);

        DraftSeedResult created = DraftSeeder.CreateNewSystem(this.DraftsRoot, new NewSystemRegistration
        {
            HardwareName = ComponentContributionFileLocationTests.Hardware,
            BoardName = ComponentContributionFileLocationTests.Board,
            ExcelDataFile = excelDataFile,
        });

        Assert.True(created.Created, created.Reason);

        // One component with one image row and no file picked yet - the state the report was made
        // from, where the drop-down is the thing about to be used.
        var board = new BoardData
        {
            Components = { new ComponentEntry { BoardLabel = "C1", Region = "PAL" } },
            ComponentImages = { new ComponentImageEntry { BoardLabel = "C1", Region = "PAL", Pin = "1" } },
        };

        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();

            window.LoadComponent(
                board,
                this.DataRoot,
                ComponentContributionFileLocationTests.Hardware,
                ComponentContributionFileLocationTests.Board,
                "PAL",
                "C1",
                excelDataFile);

            ContributionComponentImageRow row = Assert.Single(
                ComponentContributionFileLocationTests.ImageRowsOf(window));

            Assert.Contains("Test Manu4/HW4/Board4/Scope baseline", row.AvailableFileLocations);
            Assert.Contains("Commodore/C64/250407/Scope baseline", row.AvailableFileLocations);
        });
    }

    private static ObservableCollection<ContributionComponentImageRow> ImageRowsOf(ComponentContributionWindow window)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            "thisComponentImageRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (ObservableCollection<ContributionComponentImageRow>)field!.GetValue(window)!;
    }

    // The same seam ComponentContributionSaveRefreshTests uses - DraftManager's own internal
    // LoadFrom, never Load(), which would resolve the user's real AppData folder.
    private static void PointDraftManagerAt(string draftsRoot)
    {
        var method = typeof(DraftManager).GetMethod(
            "LoadFrom",
            BindingFlags.Static | BindingFlags.NonPublic);

        method!.Invoke(null, new object?[] { draftsRoot });
    }
}
