using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Contribute tab's "Show on board" as Main carries it out (Main.ShowComponentOnBoard - CRT 4.0
// step 2, slice 3c): the Schematics tab with ONLY that component selected - at once, or, when the
// board on screen had to be read again first, once THAT board's load is in (a reload drops the
// selection, and another board's load must not take it). Main's StartAsync is never called
// (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainShowOnBoardTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public MainShowOnBoardTests()
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

    private static List<ComponentListItem> Items() =>
    [
        new ComponentListItem { BoardLabel = "U1", DisplayText = "U1 | VIC", SelectionKey = "U1" },
        new ComponentListItem { BoardLabel = "U8", DisplayText = "U8 | CPU", SelectionKey = "U8" },
    ];

    private static List<string> Selected(CRT.Main window) =>
        window.ComponentFilterListBox.SelectedItems!.Cast<ComponentListItem>().Select(item => item.BoardLabel).ToList();

    [Fact]
    public void Show_on_board_selects_only_that_component_and_shows_the_Schematics_tab()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            window.ComponentFilterListBox.ItemsSource = MainShowOnBoardTests.Items();
            window.ComponentFilterListBox.Selection.Select(0);

            window.ShowComponentOnBoard("U8", readBoardFirst: false);

            Assert.Equal(["U8"], MainShowOnBoardTests.Selected(window));
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // Read again first: selected by that board's load, once - and never by another board's.
    [Fact]
    public void After_reading_the_board_again_it_is_selected_by_that_boards_load_only()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainShowOnBoardTests.BoardKey };

            window.ShowComponentOnBoard("U8", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();
            window.ComponentFilterListBox.ItemsSource = MainShowOnBoardTests.Items();

            // Another board loaded meanwhile: nothing selected, and nothing left waiting.
            window.DoWhatWasAskedAfterLoad("Commodore 64|250425");
            Assert.Empty(MainShowOnBoardTests.Selected(window));
            window.DoWhatWasAskedAfterLoad(MainShowOnBoardTests.BoardKey);
            Assert.Empty(MainShowOnBoardTests.Selected(window));

            window.ShowComponentOnBoard("U8", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();
            window.ComponentFilterListBox.ItemsSource = MainShowOnBoardTests.Items();

            window.DoWhatWasAskedAfterLoad(MainShowOnBoardTests.BoardKey);

            Assert.Equal(["U8"], MainShowOnBoardTests.Selected(window));
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // ###########################################################################################
    // What waited for a load that never reached its end (code review, 2026-10-26) - the board file
    // being replaced as it was read, say - is dropped by the next load: it stayed, and a reload of the
    // same board minutes later replaced the selection out of the blue (or started a picture tool).
    // ###########################################################################################
    [Fact]
    public void What_waited_for_a_load_that_never_finished_is_dropped_by_the_next_load()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainShowOnBoardTests.BoardKey };

            // Its load ends early: no board is selected in a test's lists.
            window.ShowComponentOnBoard("U8", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();

            window.ReloadCurrentBoardFromDisk(string.Empty);
            Dispatcher.UIThread.RunJobs();
            window.ComponentFilterListBox.ItemsSource = MainShowOnBoardTests.Items();

            window.DoWhatWasAskedAfterLoad(MainShowOnBoardTests.BoardKey);

            Assert.Empty(MainShowOnBoardTests.Selected(window));
        });
    }
}
