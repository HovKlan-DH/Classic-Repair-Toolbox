using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "Add a new board"'s last step (Main.NewBoard.cs). The dialog itself cannot be driven headlessly, so
// the step after it is driven on a real Main; Main's StartAsync is never called (MainWindowTests' header).
//
// *** A BOARD THAT COULD NOT BE CREATED SAYS WHY (code review, 2026-11-02). *** After the agreement the
// dialog has closed, and a seed that failed was only logged: nothing appeared, nothing was said, and the
// names typed were gone. A board's first draft says it on the tab's problem line; so does this.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainNewBoardTests : IDisposable
{
    private const string ExcelDataFile = "Comodore/C64/999001/board.json";

    private readonly TempWorkspace thisWorkspace = new();

    private static readonly NewBoardRegistration Registration = new()
    {
        HardwareName = "C64",
        BoardName = "999001",
        ExcelDataFile = MainNewBoardTests.ExcelDataFile,
    };

    public MainNewBoardTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void A_board_that_could_not_be_created_says_why_on_the_tabs_problem_line()
    {
        // The seeder refuses a folder that holds a draft already.
        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, MainNewBoardTests.ExcelDataFile),
            new DraftMarker { BoardKey = MainNewBoardTests.ExcelDataFile });

        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            window.TabDrafts.HardwareBoardsOverrideForTests = [];

            Assert.False(window.TryCreateNewBoard(MainNewBoardTests.Registration));

            Assert.Contains("A draft of this board already exists", window.TabDrafts.DraftProblemForTests, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void A_board_created_says_nothing()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            window.TabDrafts.HardwareBoardsOverrideForTests = [];

            Assert.True(window.TryCreateNewBoard(MainNewBoardTests.Registration));

            Assert.Null(window.TabDrafts.DraftProblemForTests);
            Assert.True(DraftBoardSource.HasDraft(DraftManager.DraftsRoot, MainNewBoardTests.ExcelDataFile));
        });
    }
}
