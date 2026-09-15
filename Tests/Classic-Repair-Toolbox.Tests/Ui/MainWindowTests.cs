using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The main window, constructed headlessly.
//
// Until Main.StartAsync existed, NOTHING here was possible: the constructor ended with
// PopulateHardwareDropDown (DataManager static state), CheckForAppUpdateNowAsync (real HTTP to
// GitHub) and StartBackgroundSyncAsync (network sync), so simply saying "new Main()" in a test
// reached the network. The largest file in the application sat at zero coverage purely because
// of where those three calls lived.
//
// They now live in StartAsync, which App calls right after Show(). These tests construct Main
// and NEVER call StartAsync - that is the whole point of the split, and it is the one rule to
// keep when adding to this file. Calling it here would put the suite back on the network and
// breach the "no test touches the network" rule in .claude/CLAUDE.md.
//
// EVERY assertion runs INSIDE UiTest.Run, including the ones that only read a property: reading
// a control from the test thread throws "the calling thread cannot access this object". Same
// pattern as WorkbooksListTests.
//
// COLLECTION NOTE: "HeadlessUi" rather than "UserSettings" or "DataManager", because these
// construct a Window and so need the shared dispatcher thread - a class can only join one
// collection. They nonetheless drive UserSettings' static state, which is safe only because
// xunit.runner.json sets "parallelizeTestCollections": false; see WorkbooksListTests' note.
// Every test that writes a setting restores it in a finally block.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainWindowTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // Points WorklogManager at a per-test temp folder. ApplyWorklogBarVisibility's ENABLE path
    // calls RefreshWorklogBar, which reads the workbook folder - without this it would read (and
    // the tab could write to) the user's real Workbook directory.
    private void RedirectWorklogToTemp()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-" + Guid.NewGuid().ToString("N")));
    }

    // The same for UserSettings, and just as necessary: tests here WRITE settings (the detached
    // thumbnails pair below), and every write calls Save() against the static _settingsFilePath -
    // which, without this, is whatever the previously-run class left it at, or the user's real
    // AppData settings.json if this class runs first. Restoring the value in a finally block is not
    // enough on its own, since SchematicsThumbnailsWindow's Closing handler also stamps its own
    // window-layout keys into that file with nothing restoring them.
    private void RedirectSettingsToTemp()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
    }

    public void Dispose()
    {
        // Leave both pointed somewhere disposable rather than at the real folder/file.
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        this.thisWorkspace.Dispose();
    }

    // ---------------------------------------------------------------------------------------
    // Construction
    // ---------------------------------------------------------------------------------------

    // The regression guard for the StartAsync split itself. If any outward-facing call ever moves
    // back into the constructor, this test starts hitting the network - and on a machine without
    // one, it fails outright.
    [Fact]
    public void The_main_window_constructs_without_running_its_startup()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            Assert.NotNull(main.MainTabControl);
            Assert.NotNull(main.TabSchematicsControl);
        });
    }

    // The constructor wires the tabs up and hands each its MainWindow. A tab left unwired shows
    // as a null reference the moment anything asks it for board state.
    [Fact]
    public void Constructing_the_window_initializes_its_tabs()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            Assert.NotNull(main.TabOscilloscopeControl);
            Assert.NotNull(main.GetControl<TabControl>("MainTabControl"));
        });
    }

    // With no board selected the board-key accessors must return an empty/absent answer rather
    // than throwing - every worklog surface calls GetCurrentBoardKey on refresh, including before
    // PopulateHardwareDropDown has ever run (which is now the state a freshly built window is in).
    [Fact]
    public void With_no_board_selected_the_board_key_is_empty_and_the_entry_is_null()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            Assert.Equal(string.Empty, main.GetCurrentBoardKey());
            Assert.Null(main.GetCurrentBoardEntry());
        });
    }

    // ---------------------------------------------------------------------------------------
    // Conditional tab visibility
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void The_oscilloscope_tab_follows_its_setting()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.EnableNetworkConnectedOscilloscopeTab;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                var tab = main.GetControl<TabItem>("OscilloscopeTabItem");

                UserSettings.EnableNetworkConnectedOscilloscopeTab = true;
                main.ApplyOscilloscopeTabVisibility();
                Assert.True(tab.IsVisible);

                UserSettings.EnableNetworkConnectedOscilloscopeTab = false;
                main.ApplyOscilloscopeTabVisibility();
                Assert.False(tab.IsVisible);
            });
        }
        finally
        {
            UserSettings.EnableNetworkConnectedOscilloscopeTab = saved;
        }
    }

    // The Workbooks tab and the worklog bar are one feature and must move together - hiding the
    // bar while leaving the tab visible was the shape of an earlier bug.
    [Fact]
    public void The_worklog_bar_and_workbooks_tab_follow_their_setting_together()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.EnableWorklog;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                var tab = main.GetControl<TabItem>("WorkbooksTabItem");
                var bar = main.GetControl<Control>("WorklogBar");

                UserSettings.EnableWorklog = true;
                main.ApplyWorklogBarVisibility();
                Assert.True(bar.IsVisible);
                Assert.True(tab.IsVisible);

                UserSettings.EnableWorklog = false;
                main.ApplyWorklogBarVisibility();
                Assert.False(bar.IsVisible);
                Assert.False(tab.IsVisible);
            });
        }
        finally
        {
            UserSettings.EnableWorklog = saved;
        }
    }

    // Hiding the SELECTED tab must move selection to a still-visible one. Without this the tab
    // control is left showing an empty page - the reason MoveSelectionOffHiddenTab exists.
    [Fact]
    public void Hiding_the_selected_workbooks_tab_moves_selection_to_a_visible_tab()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.EnableWorklog;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                var tabControl = main.GetControl<TabControl>("MainTabControl");
                var workbooksTab = main.GetControl<TabItem>("WorkbooksTabItem");

                UserSettings.EnableWorklog = true;
                main.ApplyWorklogBarVisibility();

                tabControl.SelectedItem = workbooksTab;
                Assert.Same(workbooksTab, tabControl.SelectedItem);

                UserSettings.EnableWorklog = false;
                main.ApplyWorklogBarVisibility();

                Assert.NotSame(workbooksTab, tabControl.SelectedItem);
                Assert.True(((TabItem)tabControl.SelectedItem!).IsVisible);
            });
        }
        finally
        {
            UserSettings.EnableWorklog = saved;
        }
    }

    // The mirror case: hiding a tab that is NOT selected must leave the current selection alone.
    // MoveSelectionOffHiddenTab guards on ReferenceEquals precisely so an unrelated tab being
    // hidden cannot yank the user off the tab they are on.
    [Fact]
    public void Hiding_an_unselected_tab_leaves_the_current_selection_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool savedWorklog = UserSettings.EnableWorklog;
        bool savedScope = UserSettings.EnableNetworkConnectedOscilloscopeTab;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                var tabControl = main.GetControl<TabControl>("MainTabControl");

                UserSettings.EnableWorklog = true;
                main.ApplyWorklogBarVisibility();

                // Park selection on a tab that is always visible, then hide a different one.
                var firstVisible = tabControl.Items.OfType<TabItem>().First(t => t.IsVisible);
                tabControl.SelectedItem = firstVisible;

                UserSettings.EnableWorklog = false;
                main.ApplyWorklogBarVisibility();

                Assert.Same(firstVisible, tabControl.SelectedItem);
            });
        }
        finally
        {
            UserSettings.EnableWorklog = savedWorklog;
            UserSettings.EnableNetworkConnectedOscilloscopeTab = savedScope;
        }
    }

    // ---------------------------------------------------------------------------------------
    // Region toggle
    // ---------------------------------------------------------------------------------------

    // The buttons are driven rather than the handlers called: OnPalRegionClick/OnNtscRegionClick
    // are private, and a click is what a user actually does. The "active" class is the observable
    // state UpdateRegionButtonsState sets, so it is what gets asserted.
    [Fact]
    public void Clicking_a_region_button_switches_the_local_region_and_the_active_class()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        string savedRegion = UserSettings.Region;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                var pal = main.GetControl<Button>("PalRegionButton");
                var ntsc = main.GetControl<Button>("NtscRegionButton");

                RaiseClick(ntsc);
                Assert.Equal("NTSC", main.LocalRegion);
                Assert.Contains("active", ntsc.Classes);
                Assert.DoesNotContain("active", pal.Classes);

                RaiseClick(pal);
                Assert.Equal("PAL", main.LocalRegion);
                Assert.Contains("active", pal.Classes);
                Assert.DoesNotContain("active", ntsc.Classes);
            });
        }
        finally
        {
            UserSettings.Region = savedRegion;
        }
    }

    // The region toggle area hides itself entirely when the current board has no explicit PAL/NTSC
    // components - and a freshly built window has no board data at all, which is that same case.
    [Fact]
    public void The_region_toggle_is_hidden_when_the_board_has_no_region_components()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            Assert.False(main.GetControl<Grid>("RegionButtonsGrid").IsVisible);
        });
    }

    // ---------------------------------------------------------------------------------------
    // Banners
    // ---------------------------------------------------------------------------------------

    // Both banners start hidden on a freshly constructed window. The main-Excel one is raised by
    // StartAsync (never called here) and the update one by the update check, so a banner visible
    // at construction time would mean something ran that should not have.
    [Fact]
    public void Both_update_banners_start_hidden()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            Assert.False(main.GetControl<Border>("UpdateBanner").IsVisible);
            Assert.False(main.GetControl<Border>("MainExcelRequiresAppUpdateBanner").IsVisible);
        });
    }

    // Dismissing one banner must not touch the other - they report different things (a newer app
    // package vs. data that needs a newer app), and the code comments call the independence out
    // explicitly.
    [Fact]
    public void Dismissing_the_main_excel_banner_leaves_the_update_banner_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var excelBanner = main.GetControl<Border>("MainExcelRequiresAppUpdateBanner");
            var updateBanner = main.GetControl<Border>("UpdateBanner");

            excelBanner.IsVisible = true;
            updateBanner.IsVisible = true;

            RaiseClick(main.GetControl<Button>("MainExcelRequiresAppUpdateBannerDismissButton"));

            Assert.False(excelBanner.IsVisible);
            Assert.True(updateBanner.IsVisible);
        });
    }

    [Fact]
    public void Dismissing_the_update_banner_leaves_the_main_excel_banner_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var excelBanner = main.GetControl<Border>("MainExcelRequiresAppUpdateBanner");
            var updateBanner = main.GetControl<Border>("UpdateBanner");

            excelBanner.IsVisible = true;
            updateBanner.IsVisible = true;

            RaiseClick(main.GetControl<Button>("UpdateBannerDismissButton"));

            Assert.True(excelBanner.IsVisible);
            Assert.False(updateBanner.IsVisible);
        });
    }

    // -----------------------------------------------------------------------------------------
    // "Detach thumbnails into their own window" persistence
    // -----------------------------------------------------------------------------------------

    // Reported: the setting never survived to the next launch. The detached window is OWNED by the
    // main window, so quitting the application closes it too and raised its Closed handler - which
    // could not tell that apart from the user dismissing the window, and turned the setting off on
    // every single exit. This test fails against that version.
    [Fact]
    public void Quitting_with_thumbnails_detached_keeps_the_setting_on_for_the_next_launch()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.DetachSchematicsThumbnails;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();

                // Shown because the detached window is opened with Show(owner), which Avalonia
                // refuses against a non-visible owner.
                main.Show();

                try
                {
                    UserSettings.DetachSchematicsThumbnails = true;
                    main.ApplyThumbnailsDetachedState();
                    Assert.True(main.IsThumbnailsWindowOpenForTests);

                    // The application is exiting: OnWindowClosing sets this before Avalonia closes
                    // the owned thumbnails window.
                    main.IsApplicationShuttingDownForTests = true;
                    main.CloseThumbnailsWindowAsUserForTests();

                    // The preference must be untouched, so the next launch comes back detached.
                    Assert.True(UserSettings.DetachSchematicsThumbnails);
                }
                finally
                {
                    main.IsApplicationShuttingDownForTests = true;
                    main.Close();
                }
            });
        }
        finally
        {
            UserSettings.DetachSchematicsThumbnails = saved;
        }
    }

    // The other half: closing the detached window by hand means "I don't want this any more", so
    // the thumbnails EMBED BACK into the Schematics tab and the Configuration checkbox unticks
    // itself - the window is a third way to turn the feature off, alongside the checkbox.
    [Fact]
    public void Closing_the_detached_window_by_hand_re_embeds_the_thumbnails_and_unticks_the_checkbox()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.DetachSchematicsThumbnails;

        try
        {
            UiTest.Run(() =>
            {
                var main = new CRT.Main();
                main.Show();

                try
                {
                    var schematics = main.GetControl<TabSchematics>("TabSchematicsControl");
                    var thumbnailList = schematics.GetControl<ListBox>("SchematicsThumbnailList");
                    var checkBox = main.GetControl<TabConfiguration>("TabConfiguration")
                        .GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

                    UserSettings.DetachSchematicsThumbnails = true;
                    main.ApplyThumbnailsDetachedState();

                    Assert.True(main.IsThumbnailsWindowOpenForTests);
                    Assert.False(thumbnailList.IsVisible);

                    // Not shutting down - the user pressed the window's own close button.
                    main.CloseThumbnailsWindowAsUserForTests();

                    // The thumbnails are back in the tab...
                    Assert.True(thumbnailList.IsVisible);
                    Assert.False(schematics.IsThumbnailsDetached);

                    // ...the window is gone, and both the setting and its checkbox are off.
                    Assert.False(main.IsThumbnailsWindowOpenForTests);
                    Assert.False(UserSettings.DetachSchematicsThumbnails);
                    Assert.False(checkBox.IsChecked);
                }
                finally
                {
                    main.IsApplicationShuttingDownForTests = true;
                    main.Close();
                }
            });
        }
        finally
        {
            UserSettings.DetachSchematicsThumbnails = saved;
        }
    }

    // Raises a Button's Click the way a real press does. The handlers under test are private
    // event handlers wired in markup, so there is nothing else to call.
    private static void RaiseClick(Button button)
    {
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
    }
}
