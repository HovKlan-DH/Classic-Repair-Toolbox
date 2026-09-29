using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The Maintainer tab as CRT's window handles it (2026-09-29: the separate CRT Maintainer
// application became a tab - Main.Maintainer.cs): hidden unless "Enable Maintainer tab" is on,
// placed between Drafts and Configuration, and - while selected - the sidebar and the worklog bar
// collapsed so the four screens get the window's width, then put back exactly as they were.
//
// Main is BUILT, never shown, with UserSettings and the workbook folder pointed at temp files -
// MainWindowTests' own setup, and in the "HeadlessUi" collection for the same reason it is (see
// its note on UserSettings).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainMaintainerTabTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public MainMaintainerTabTests()
    {
        this.RedirectToTemp();
    }

    public void Dispose()
    {
        this.RedirectToTemp();
        this.thisWorkspace.Dispose();
    }

    private void RedirectToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
        // An EMPTY file, not a missing one: LoadFrom keeps the settings already in memory when the
        // file does not exist, so a missing one would carry the previous test's choices over.
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile(Guid.NewGuid().ToString("N") + ".json", "{}"));
    }

    // Almost nobody running CRT has a maintainer account, so the tab is not there until asked for.
    [Fact]
    public void The_tab_is_hidden_until_it_is_turned_on()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            Assert.False(window.MaintainerTabItem.IsVisible);

            UserSettings.EnableMaintainerTab = true;
            window.ApplyMaintainerTabVisibility();

            Assert.True(window.MaintainerTabItem.IsVisible);
        });
    }

    // Beside the other contribution work - Contribute, Drafts, Maintainer - and before Configuration.
    [Fact]
    public void The_tab_sits_after_Drafts_and_before_Configuration()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main();
            var items = window.MainTabControl.Items.Cast<object>().ToList();

            int maintainer = items.IndexOf(window.MaintainerTabItem);

            Assert.Equal(items.IndexOf(window.DraftsTabItem) + 1, maintainer);
            Assert.Equal(items.IndexOf(window.ConfigurationTabItem) - 1, maintainer);
            Assert.Equal("Maintainer", window.MaintainerTabItem.Header);
        });
    }

    // ###########################################################################################
    // *** THE SIDEBAR AND THE WORKLOG BAR GO WHILE THE TAB IS SELECTED, AND COME BACK AS THEY
    // WERE. *** Neither has anything to do with reviewing, and the screens need the width. The
    // width put back is the one there before - not the saved setting, not a default - and the
    // saved setting is never overwritten with the collapsed 0.
    // ###########################################################################################
    [Fact]
    public void Selecting_the_tab_collapses_the_sidebar_and_worklog_bar_and_leaving_it_restores_them()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.EnableWorklog = true;
            UserSettings.LeftPanelWidth = 260;

            var window = new CRT.Main();
            window.RootGrid.ColumnDefinitions[0].Width = new GridLength(237);

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;

            Assert.True(window.IsMaintainerLayoutActive);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[0].Width.Value);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[0].MinWidth);
            Assert.Equal(0, window.RootGrid.ColumnDefinitions[1].Width.Value);
            Assert.False(window.LeftPanel.IsVisible);
            Assert.False(window.MainSplitter.IsVisible);
            Assert.False(window.WorklogBar.IsVisible);

            window.MainTabControl.SelectedItem = window.SchematicsTabItem;

            Assert.False(window.IsMaintainerLayoutActive);
            Assert.Equal(new GridLength(237), window.RootGrid.ColumnDefinitions[0].Width);
            Assert.Equal(80, window.RootGrid.ColumnDefinitions[0].MinWidth);
            Assert.Equal(4, window.RootGrid.ColumnDefinitions[1].Width.Value);
            Assert.True(window.LeftPanel.IsVisible);
            Assert.True(window.MainSplitter.IsVisible);
            Assert.True(window.WorklogBar.IsVisible);

            Assert.Equal(260, UserSettings.LeftPanelWidth);
        });
    }

    // The worklog bar comes back only if it was on to begin with.
    [Fact]
    public void Leaving_the_tab_keeps_the_worklog_bar_hidden_when_the_worklog_is_off()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;
            UserSettings.EnableWorklog = false;

            var window = new CRT.Main();

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;
            window.MainTabControl.SelectedItem = window.SchematicsTabItem;

            Assert.False(window.WorklogBar.IsVisible);
            Assert.True(window.LeftPanel.IsVisible);
        });
    }

    // Unticked while the tab is selected: the selection moves off it, and that alone restores the
    // layout - the window is never left with no sidebar and no Maintainer tab either.
    [Fact]
    public void Turning_the_tab_off_while_it_is_selected_moves_off_it_and_restores_the_layout()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();
            window.MainTabControl.SelectedItem = window.MaintainerTabItem;

            UserSettings.EnableMaintainerTab = false;
            window.ApplyMaintainerTabVisibility();

            Assert.False(window.MaintainerTabItem.IsVisible);
            Assert.NotSame(window.MaintainerTabItem, window.MainTabControl.SelectedItem);
            Assert.False(window.IsMaintainerLayoutActive);
            Assert.True(window.LeftPanel.IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE COMPONENT FILTER NEVER TAKES FOCUS BACK OVER THE MAINTAINER TAB. *** The window pulls
    // focus to the component filter after every click (ShouldReturnFocusToComponentSearch). The tab
    // has a password box, a decision comment box and an invitation address, and the filter is
    // hidden with the sidebar - so it backs off whenever the tab is shown, not only with a table
    // open (the Drafts tab's rule).
    // ###########################################################################################
    [Fact]
    public void The_component_filter_does_not_take_focus_over_the_maintainer_tab()
    {
        UiTest.Run(() =>
        {
            UserSettings.EnableMaintainerTab = true;

            var window = new CRT.Main();

            window.MainTabControl.SelectedItem = window.SchematicsTabItem;
            Assert.True(window.ShouldReturnFocusToComponentSearch());

            window.MainTabControl.SelectedItem = window.MaintainerTabItem;
            Assert.False(window.ShouldReturnFocusToComponentSearch());
        });
    }

    // ###########################################################################################
    // *** "SHOW CHANGES ONLY" IS REMEMBERED IN CRT'S SETTINGS (2026-09-29). *** The separate
    // application kept it in a file of its own; Main now hands the tab UserSettings'
    // MaintainerShowChangesOnly, and every later CHOICE is written back.
    // ###########################################################################################
    [Fact]
    public void The_tables_show_changes_only_comes_from_and_goes_back_to_CRTs_settings()
    {
        UiTest.Run(() =>
        {
            UserSettings.MaintainerShowChangesOnly = true;

            var window = new CRT.Main();
            CRT.BoardTableEditor editor = window.TabMaintainer.TableEditorForTests;

            Assert.True(editor.OnlyChangesWanted);

            editor.OnlyChanges = false;

            Assert.False(UserSettings.MaintainerShowChangesOnly);
        });
    }
}
