using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The picture tools and the component menu as Main carries them between the tabs (Main.PictureTools.cs
// - CRT 4.0 step 2, slice 6): a tool asked for on the Contribute tab opens the Schematics tab - on the
// board read again first when asked, the tool started by THAT board's load - and its end goes back to
// the Contribute tab. The component menu's highlight item does what a right-click did by itself
// before; "Edit this component" opens the Contribute tab. Main's StartAsync is never called
// (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainPictureToolsTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public MainPictureToolsTests()
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

    private static void Click(Control owner, string name) =>
        owner.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

    // ###########################################################################################
    // The Schematics tab with the tool asked for - here refused, the board shown having no such
    // picture - and its end back on the Contribute tab.
    // ###########################################################################################
    [Fact]
    public void A_tool_from_the_Contribute_tab_opens_the_Schematics_tab_and_its_end_goes_back()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: false);

            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
            TabSchematics tab = window.TabSchematicsControl;
            Assert.Equal("Components cannot be marked on [Board top]", tab.GetControl<TextBlock>("PictureToolTitleText").Text);

            Click(tab, "PictureToolCancelButton");

            Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);
        });
    }

    // Read again first: started by that board's load, once - and never by another board's.
    [Fact]
    public void After_reading_the_board_again_the_tool_is_started_by_that_boards_load_only()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();

            window.DoWhatWasAskedAfterLoad("Commodore 64|250425");
            Assert.False(tab.GetControl<Border>("PictureToolBarBorder").IsVisible);

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Dispatcher.UIThread.RunJobs();
            window.DoWhatWasAskedAfterLoad(MainPictureToolsTests.BoardKey);

            Assert.True(tab.GetControl<Border>("PictureToolBarBorder").IsVisible);
        });
    }

    // ###########################################################################################
    // A picture tool on while the Contribute tab asks for the board to be read again (code review,
    // 2026-10-26): the reload reset the tab, which ended the label editor and threw its rectangles
    // away without a word. It waits for the tool to end now - and "Mark components" asked for again
    // shows the tool that is on rather than queueing another.
    // ###########################################################################################
    [Fact]
    public void A_reload_asked_for_while_a_tool_is_on_waits_until_the_tool_ends()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainPictureToolsTests.BoardKey };
            TabSchematics tab = window.TabSchematicsControl;
            tab.CurrentBoardDataOverrideForTests = new BoardData
            {
                Schematics = [new BoardSchematicEntry { SchematicName = "Board top" }],
                Components = [new ComponentEntry { BoardLabel = "U1", Category = "IC" }],
                ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Board top", BoardLabel = "U1", X = "10", Y = "10", Width = "40", Height = "20" }],
            };

            var thumbnail = new Tabs.TabSchematics.SchematicThumbnail { Name = "Board top" };
            tab.currentThumbnails.Add(thumbnail);
            tab.SelectedThumbnail = thumbnail;
            tab.BeginLabelEditorModeForTests();
            Assert.Single(tab.LabelEditorWorkingRowsForTests);

            window.ReloadCurrentBoardFromDisk("Board top");
            Dispatcher.UIThread.RunJobs();

            Assert.True(tab.IsLabelEditorModeForTests);
            Assert.Single(tab.LabelEditorWorkingRowsForTests);
            Assert.True(window.IsReloadWaitingForToolForTests);

            window.StartPictureTool(PictureTool.MarkComponents, "Board top", readBoardFirst: true);
            Assert.Same(window.SchematicsTabItem, window.MainTabControl.SelectedItem);
            Assert.True(tab.IsLabelEditorModeForTests);

            // The tool ends: the board is read again then.
            tab.CancelLabelEditorChangesForTests();
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsReloadWaitingForToolForTests);
        });
    }

    // ###########################################################################################
    // The menu's first item: the component's highlight on, then off - what a right-click on a
    // component did by itself before the menu.
    // ###########################################################################################
    [Fact]
    public void The_component_menu_puts_the_highlight_on_and_takes_it_away()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            window.ComponentFilterListBox.ItemsSource = new List<ComponentListItem>
            {
                new() { BoardLabel = "U1", DisplayText = "U1 | VIC", SelectionKey = "U1" },
                new() { BoardLabel = "U8", DisplayText = "U8 | CPU", SelectionKey = "U8" },
            };

            TabSchematics tab = window.TabSchematicsControl;

            tab.ShowComponentMenuForTests("U8");
            Assert.Equal("Highlight [U8]", tab.GetControl<Button>("MenuHighlightButton").Content);
            Click(tab, "MenuHighlightButton");
            Assert.True(tab.IsComponentSelectedForTests("U8"));

            tab.ShowComponentMenuForTests("U8");
            Assert.Equal("Take the highlight away from [U8]", tab.GetControl<Button>("MenuHighlightButton").Content);
            Click(tab, "MenuHighlightButton");
            Assert.False(tab.IsComponentSelectedForTests("U8"));
        });
    }

    // "Edit this component": the Contribute tab.
    [Fact]
    public async Task Edit_this_component_opens_the_Contribute_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();

            await window.EditComponentAsync("U8");

            Assert.Same(window.ContributeTabItem, window.MainTabControl.SelectedItem);
        });
    }
}
