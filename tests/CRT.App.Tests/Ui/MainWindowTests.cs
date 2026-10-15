using Avalonia;
using Avalonia.Controls;
using Avalonia.LogicalTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

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
    // The Drafts tab appearing after a local edit
    // ---------------------------------------------------------------------------------------

    // ###########################################################################################
    // *** THE REPORTED BUG (2026-09-23): "Save to draft" did not make the Drafts tab appear, and
    // the application had to be restarted before it showed. ***
    //
    // The tab is hidden until a draft exists, and ApplyDraftsTabVisibility is what re-evaluates
    // that. It was called at startup, after "Add a new system", and from inside the Drafts tab
    // after a discard - but by NEITHER of the two paths that create a draft for an existing board:
    // the Contribute window's component save and the label editor's save. Both funnel through
    // Main.ReloadCurrentBoardFromDisk, so that is where the refresh now lives, which is the same
    // "one funnel, not one call per editor" rule Main.RefreshWorklogBar follows for worklogs.
    //
    // Driven through ReloadCurrentBoardFromDisk rather than through either editor: that method IS
    // the shared contract the fix relies on, and a test per editor would pass while the shared
    // step was missing from a third.
    // ###########################################################################################
    private static readonly string DraftBoardKey =
        "Commodore/C64/250407/Data C64 250407.xlsx";

    private void GiveTheDraftsTabSomethingToShow(CRT.Main window)
    {
        // A draft is a folder with a marker in it - a workbook alone is not one, which is the rule
        // DraftBoardSource enforces everywhere else.
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts-" + Guid.NewGuid().ToString("N")));

        DraftMarkerStore.Save(
            DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, MainWindowTests.DraftBoardKey),
            new DraftMarker
            {
                BoardKey = MainWindowTests.DraftBoardKey,
                BaseRevision = "2026-09-01",
                NewBoard = null,
                CreatedUtc = "2026-09-23T00:00:00Z",
            });

        // RefreshDrafts pairs drafts against the known boards, which normally come from
        // DataManager - overridden so this test needs no data tree.
        window.TabDrafts.HardwareBoardsOverrideForTests =
        [
            new HardwareBoardEntry
            {
                HardwareName = "Commodore 64",
                BoardName = "250407",
                ExcelDataFile = MainWindowTests.DraftBoardKey,
            },
        ];
    }

    [Fact]
    public void Reloading_after_a_local_edit_REVEALS_the_drafts_tab()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            // Hidden in the markup, and nothing has happened yet to change that.
            Assert.False(window.DraftsTabItem.IsVisible);

            this.GiveTheDraftsTabSomethingToShow(window);

            // What BOTH save paths call. Before the fix this cleared the board cache and reloaded,
            // and left the tab hidden until the next restart.
            window.ReloadCurrentBoardFromDisk(string.Empty);

            Assert.True(window.DraftsTabItem.IsVisible);
        });
    }

    // The anti-vacuity half: with no draft on disk the reload must NOT reveal the tab, or the
    // "fix" would simply be showing it always and somebody who has never contributed would see a
    // tab they have no use for.
    [Fact]
    public void Reloading_with_NO_draft_leaves_the_drafts_tab_hidden()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var window = new CRT.Main();

            DraftManager.LoadFrom(
                this.thisWorkspace.Path_("Drafts-empty-" + Guid.NewGuid().ToString("N")));

            window.TabDrafts.HardwareBoardsOverrideForTests = [];

            window.ReloadCurrentBoardFromDisk(string.Empty);

            Assert.False(window.DraftsTabItem.IsVisible);
        });
    }

    // ###########################################################################################
    // *** THE REPORTED BUG (2026-09-24): typing into a cell of the Drafts tab's table did nothing.
    // *** Every click in the window hands focus back to the always-on component filter on the
    // left, and a click on a grid cell focuses the GRID - not a TextBox, which was the only thing
    // that stopped the steal. So the keystrokes went into the component filter instead.
    //
    // Asserted on both sides of the table being open: the steal must stop while it is, and must
    // come BACK once it closes - the Drafts tab without a table has nothing to type into, and
    // exempting the whole tab would quietly switch the filter off there.
    // ###########################################################################################
    [Fact]
    public async Task While_the_Drafts_table_is_open_a_click_does_NOT_pull_focus_to_the_component_filter()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main();
            this.GiveTheDraftsTabSomethingToShow(window);

            // The table opens the draft's WORKBOOK, which the marker alone does not provide.
            string workbook = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, MainWindowTests.DraftBoardKey);
            CachedWorkbooks.Write(workbook, new BoardData { Components = [new ComponentEntry { BoardLabel = "C1" }] });

            window.TabDrafts.PublishedBoardOverrideForTests = _ => null;
            window.ApplyDraftsTabVisibility();
            window.MainTabControl.SelectedItem = window.DraftsTabItem;

            // The Drafts tab with no table: the ordinary steal still applies.
            Assert.True(window.ShouldReturnFocusToComponentSearch());

            await window.TabDrafts.OpenTableAsync(window.TabDrafts.HardwareBoardsOverrideForTests![0]);
            Assert.False(window.ShouldReturnFocusToComponentSearch());

            Assert.True(await window.TabDrafts.CloseTableAsync());
            Assert.True(window.ShouldReturnFocusToComponentSearch());
        });
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

    // ###########################################################################################
    // The Drafts tab's title is IndianRed in both themes (owner request, 2026-09-25): a draft
    // is temporary, and the tab says so. Asserted on the header the window really builds, through
    // its theme key, so a header turned back into plain text - which the tab's selected and hover
    // states would recolour - or a key missing from one theme fails here.
    // ###########################################################################################
    [Fact]
    public void The_Drafts_tab_title_is_IndianRed_in_both_themes()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            // The title beside its badge (2026-09-30) - still its own block, in its own colour.
            TextBlock header = main.DraftsTabHeaderText;
            Assert.Equal("Drafts", header.Text);
            Assert.Contains(header, Assert.IsType<StackPanel>(main.DraftsTabItem.Header).Children);

            foreach (Avalonia.Styling.ThemeVariant variant in new[] { Avalonia.Styling.ThemeVariant.Light, Avalonia.Styling.ThemeVariant.Dark })
            {
                Assert.True(Avalonia.Application.Current!.TryGetResource("Main_TabHeader_Drafts_Fg", variant, out object? brush));
                Assert.Equal(Avalonia.Media.Colors.IndianRed, Assert.IsAssignableFrom<Avalonia.Media.ISolidColorBrush>(brush).Color);
            }

            Assert.Equal(Avalonia.Media.Colors.IndianRed, Assert.IsAssignableFrom<Avalonia.Media.ISolidColorBrush>(header.Foreground).Color);
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

    // The notices, by name - the shift test below shows each one in turn.
    private static readonly string[] NoticeNames =
    {
        "SettingsUnreadableBanner",
        "MainExcelRequiresAppUpdateBanner",
        "UpdateBanner",
        "DataCheckBanner",
        "SyncBanner",
        "SourceSwitchBanner",
        "SchematicsLabelEditorModeBanner",
    };

    // ###########################################################################################
    // A notice coming or going must not move the tabs (owner report, 2026-10-11).
    //
    // The notices were rows above the tabs, so every one that appeared or went moved the whole
    // window down or up - the launch data check alone did it twice - and on the Schematics tab
    // the image is fitted to its container, so it was resized as well. They float over the tabs'
    // bottom-right corner now. Each one is shown in turn on a SHOWN window (an unshown one has no
    // layout), and the tabs' place and size compared with nothing showing.
    // ###########################################################################################
    [Fact]
    public void Showing_or_hiding_a_notice_leaves_the_tabs_where_they_are()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            main.Show();

            try
            {
                var tabs = main.GetControl<TabControl>("MainTabControl");
                main.UpdateLayout();
                Rect before = tabs.Bounds;

                foreach (string name in NoticeNames)
                {
                    var notice = main.GetControl<Border>(name);

                    notice.IsVisible = true;
                    main.UpdateLayout();

                    Assert.True(notice.Bounds.Height > 0, $"{name} was not laid out");
                    Assert.Equal(before, tabs.Bounds);

                    // In the tabs' bottom-right corner, not above them.
                    Point corner = notice.TranslatePoint(new Point(notice.Bounds.Width, notice.Bounds.Height), tabs)!.Value;
                    Assert.InRange(tabs.Bounds.Width - corner.X, 0, 24);
                    Assert.InRange(tabs.Bounds.Height - corner.Y, 0, 24);

                    notice.IsVisible = false;
                    main.UpdateLayout();

                    Assert.Equal(before, tabs.Bounds);
                }
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // In a small window a notice stays in the right half of the tabs.
    //
    // Found by rendering CRT at its smallest: a full-width card ran off the thumbnail list and over
    // the Schematics tab's "Global settings" and "Labels visible" panels - and the label editor's
    // card cannot be closed, so they stayed covered for the whole edit. Half the tabs' width keeps
    // it over the thumbnails, at the default split.
    // ###########################################################################################
    [Fact]
    public void In_a_small_window_a_notice_stays_in_the_right_half_of_the_tabs()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main { Width = 820, Height = 520 };
            main.Show();

            try
            {
                var tabs = main.GetControl<TabControl>("MainTabControl");
                var notice = main.GetControl<Border>("SchematicsLabelEditorModeBanner");
                notice.IsVisible = true;
                main.UpdateLayout();

                Point left = notice.TranslatePoint(new Point(0, 0), tabs)!.Value;
                Assert.True(
                    left.X >= tabs.Bounds.Width / 2,
                    $"the notice starts at {left.X:0}, left of the middle of tabs {tabs.Bounds.Width:0} wide");
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // The launch data check says what happened (owner report, 2026-10-11: "I am checking for new
    // data updates at app launch, but I do not see any information about that anywhere").
    //
    // Here the online check could not run at all - no manifest was fetched, as when offline - which
    // used to pass in complete silence. The tests never fetch one, so this is that launch.
    // ###########################################################################################
    [Fact]
    public async Task When_the_launch_data_check_could_not_run_a_notice_says_so()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.CheckDataOnLaunch;

        try
        {
            UserSettings.CheckDataOnLaunch = true;
            CRT.Main? main = null;
            var timers = new NoticeTimers();

            await UiTest.RunAsync(async () =>
            {
                main = new CRT.Main();
                main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;
                await main.StartBackgroundSyncForTests();
            });

            UiTest.Run(() =>
            {
                var shown = main!.GetControl<StackPanel>("NoticeStack").Children
                    .Where(card => card.IsVisible)
                    .SelectMany(card => card.GetLogicalDescendants().OfType<TextBlock>())
                    .Select(NoticeEmphasis.TextOf)
                    .ToList();

                // The source is BETA or stable by whatever the settings say - the words around it count.
                Assert.Contains(shown, text =>
                    text.StartsWith("Could not check the ", StringComparison.Ordinal) &&
                    text.Contains(" source for new data", StringComparison.Ordinal));

                // Nobody asked for this one, and it would greet everybody working offline at every
                // launch - so it goes by itself.
                timers.RunHides();
                Assert.False(main!.GetControl<Border>("DataCheckBanner").IsVisible);
            });
        }
        finally
        {
            UserSettings.CheckDataOnLaunch = saved;
        }
    }

    // ###########################################################################################
    // A launch check that could not run ends nothing it did not start (code review, 2026-10-11).
    // Its finally stopped the spinner and finished the card whether or not it had started them, so
    // a check already running lost its spinner - the icon clickable mid-check, letting a second
    // check start beside it - and its card said it could not check.
    // ###########################################################################################
    [Fact]
    public async Task A_launch_check_that_could_not_run_leaves_a_running_check_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();
        bool saved = UserSettings.CheckDataOnLaunch;

        try
        {
            UserSettings.CheckDataOnLaunch = true;

            await UiTest.RunAsync(async () =>
            {
                var main = new CRT.Main();
                main.ScheduleNoticeTimerOverrideForTests = new NoticeTimers().Schedule;

                main.BeginDataSyncProgress();
                main.StartDataSyncStatusIconSpin();

                await main.StartBackgroundSyncForTests();

                Assert.True(main.IsDataSyncStatusIconSpinning);
                Assert.StartsWith("Checking data from the ", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText")));

                main.StopDataSyncStatusIconSpin();
                Assert.False(main.IsDataSyncStatusIconSpinning);
            });
        }
        finally
        {
            UserSettings.CheckDataOnLaunch = saved;
        }
    }

    // While a check runs, its card says what it is doing - the same lines as the sync icon's tooltip.
    [Fact]
    public void A_running_data_check_says_on_its_card_what_it_is_doing()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var card = main.GetControl<Border>("DataCheckBanner");
            var text = main.GetControl<TextBlock>("DataCheckBannerText");

            main.BeginDataSyncProgress();

            Assert.True(card.IsVisible);
            Assert.StartsWith("Checking data from the ", NoticeEmphasis.TextOf(text));
            Assert.False(main.GetControl<Button>("DataCheckBannerRefreshButton").IsVisible);

            main.ReportDataSyncStatus("Downloading file [2] of [9] from the stable source:");
            main.ReportDataSyncFile("Commodore/C64/250407/pcb.png");

            string[] lines = NoticeEmphasis.TextOf(text).Split('\n');
            Assert.Equal(3, lines.Length);
            Assert.Equal("Downloading file [2] of [9] from the stable source:", lines[1]);
            Assert.Equal("Commodore/C64/250407/pcb.png", lines[2]);

            // The heading's source name, the counts and the line's source in bold (owner request,
            // 2026-10-11) - the heading names whichever source the settings say.
            Assert.Equal([AppConfig.GetOnlineSourceLabel().Split(' ')[0], "2", "9", "stable"], BoldRuns.BoldPiecesOf(text));

            // A closing summary is not progress - the card keeps the last real line.
            main.ReportDataSyncStatus("Sync complete (0 new, 0 updated, 0 failed, 0 invalid)");
            Assert.Equal("Downloading file [2] of [9] from the stable source:", NoticeEmphasis.TextOf(text).Split('\n')[1]);
        });
    }

    // Nothing changed: the card says so and goes by itself - the launch check's usual outcome.
    [Fact]
    public void A_data_check_that_changed_nothing_says_so_and_goes_by_itself()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var timers = new NoticeTimers();
            main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;
            var card = main.GetControl<Border>("DataCheckBanner");

            main.BeginDataSyncProgress();
            main.FinishDataCheck(0, atLaunch: true);

            Assert.True(card.IsVisible);
            Assert.StartsWith("All data files are up to date with the ", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText")));
            Assert.Equal(1, timers.HideCount);

            timers.RunHides();

            Assert.False(card.IsVisible);
        });
    }

    // Files changed: the card offers "Refresh Board" and stays - there is something to do.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void A_data_check_that_changed_files_stays_and_offers_to_refresh_the_board(bool atLaunch)
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var timers = new NoticeTimers();
            main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;

            main.BeginDataSyncProgress();
            main.FinishDataCheck(3, atLaunch);

            Assert.True(main.GetControl<Border>("DataCheckBanner").IsVisible);
            Assert.True(main.GetControl<Button>("DataCheckBannerRefreshButton").IsVisible);
            Assert.EndsWith("- please refresh board", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText")));
            Assert.Equal(0, timers.HideCount);
        });
    }

    // Closed while it ran, the card is not brought back to say "all up to date" - but it is for
    // files to refresh, which the user would otherwise not know about.
    [Fact]
    public void A_data_check_closed_while_it_ran_comes_back_only_with_something_to_do()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var card = main.GetControl<Border>("DataCheckBanner");
            var close = main.GetControl<Button>("DataCheckBannerDismissButton");

            main.BeginDataSyncProgress();
            RaiseClick(close);
            main.FinishDataCheck(0, atLaunch: true);

            Assert.False(card.IsVisible);

            main.BeginDataSyncProgress();
            RaiseClick(close);
            main.FinishDataCheck(2, atLaunch: true);

            Assert.True(card.IsVisible);
        });
    }

    // A new check puts "please wait" back on the card - the earlier "all up to date" must not then
    // take the new check's card down when its time comes.
    [Fact]
    public void A_new_data_check_is_not_taken_down_by_the_last_ones_timer()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var timers = new NoticeTimers();
            main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;
            var card = main.GetControl<Border>("DataCheckBanner");

            main.BeginDataSyncProgress();
            main.FinishDataCheck(0, atLaunch: false);
            main.BeginDataSyncProgress();

            timers.RunHides();

            Assert.True(card.IsVisible);
            Assert.StartsWith("Checking data from the ", NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText")));
        });
    }

    // ###########################################################################################
    // "--demo-notices" plays every notice once (owner request, 2026-10-11). Run here without its
    // waits: every card in the stack must have been shown, so a new card cannot be left out of the
    // demo - and a pretend check must leave no "last check" behind in the tooltip.
    // ###########################################################################################
    [Fact]
    public async Task The_notice_demo_shows_every_notice_and_claims_no_real_check()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            var timers = new NoticeTimers();
            main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;
            main.NoticeDemoDelayOverrideForTests = _ => Task.CompletedTask;

            var shown = new HashSet<string>();
            foreach (Control card in main.GetControl<StackPanel>("NoticeStack").Children)
            {
                card.PropertyChanged += (_, e) =>
                {
                    if (e.Property == Visual.IsVisibleProperty && e.NewValue is true)
                    {
                        shown.Add(card.Name!);
                    }
                };
            }

            await main.PlayNoticeDemoAsync();

            Assert.Equal(NoticeNames.OrderBy(name => name), shown.OrderBy(name => name));

            string tip = ToolTip.GetTip(main.GetControl<Border>("DataSyncStatusIconBorder")) as string ?? string.Empty;
            Assert.DoesNotContain("Last check", tip, StringComparison.Ordinal);

            // The demo ends with "files updated", which stays - and the label editor's card is gone.
            Assert.True(main.GetControl<Border>("DataCheckBanner").IsVisible);
            Assert.False(main.GetControl<Border>("SchematicsLabelEditorModeBanner").IsVisible);
        });
    }

    // ###########################################################################################
    // A real data check's card is left as it is by the demo (code review, 2026-10-11). The demo
    // put its own progress over a real "please refresh board" and then "all up to date", which goes
    // by itself - so the real card's reason to stay was gone.
    // ###########################################################################################
    [Fact]
    public async Task The_notice_demo_leaves_a_real_data_check_card_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            main.ScheduleNoticeTimerOverrideForTests = new NoticeTimers().Schedule;
            main.NoticeDemoDelayOverrideForTests = _ => Task.CompletedTask;

            main.BeginDataSyncProgress();
            main.FinishDataCheck(5, atLaunch: true);
            string real = NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText"));

            await main.PlayNoticeDemoAsync();

            Assert.True(main.GetControl<Border>("DataCheckBanner").IsVisible);
            Assert.Equal(real, NoticeEmphasis.TextOf(main.GetControl<TextBlock>("DataCheckBannerText")));
        });
    }

    // ###########################################################################################
    // Every card's bracketed values and source names are bold (owner request, 2026-10-11:
    // "[<b>3</b>] files updated in background", "Checking data from the <b>stable</b> source", "and
    // do that for others also"). Checked on every card the demo puts words on, against the rule
    // itself - so a card whose words go into Text directly, past NoticeEmphasis.Show, fails here.
    // ###########################################################################################
    [Fact]
    public async Task Every_notice_card_shows_its_counts_and_source_names_in_bold()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        await UiTest.RunAsync(async () =>
        {
            var main = new CRT.Main();
            main.ScheduleNoticeTimerOverrideForTests = new NoticeTimers().Schedule;
            main.NoticeDemoDelayOverrideForTests = _ => Task.CompletedTask;

            await main.PlayNoticeDemoAsync();

            foreach (Control card in main.GetControl<StackPanel>("NoticeStack").Children)
            {
                foreach (TextBlock block in card.GetLogicalDescendants().OfType<TextBlock>())
                {
                    string words = NoticeEmphasis.TextOf(block);

                    Assert.Equal(
                        NoticeEmphasis.Runs(words).Where(run => run.IsBold).Select(run => run.Text).ToList(),
                        BoldRuns.BoldPiecesOf(block));
                }
            }

            // And the demo did put such words on the cards - the check above is not vacuous.
            Assert.Equal(["3"], BoldRuns.BoldPiecesOf(main.GetControl<TextBlock>("DataCheckBannerText")));
            Assert.Equal(["stable", "BETA", "stable"], BoldRuns.BoldPiecesOf(main.GetControl<TextBlock>("SourceSwitchBannerText")));
        });
    }

    // Collects the notices' timers instead of a DispatcherTimer, which never ticks headless. A
    // "goes by itself" timer is told apart from a click guard by its length.
    private sealed class NoticeTimers
    {
        private readonly List<Action> thisHides = new();

        public int HideCount => this.thisHides.Count;

        public IDisposable Schedule(Action elapsed, TimeSpan after)
        {
            if (after == CRT.Main.NoticeGoesAfter)
            {
                this.thisHides.Add(elapsed);
            }

            return new CancelOnDispose();
        }

        public void RunHides()
        {
            foreach (Action hide in this.thisHides.ToList())
            {
                hide();
            }
        }
    }

    // ###########################################################################################
    // A settings file that could not be read is said on a card of its own, with where it was kept
    // (code review, 2026-10-13: it was reset to defaults in silence, and the first setting changed
    // saved the defaults over it). The card STAYS until closed, though it has nothing to press - the
    // one card that does (decided 2026-10-13; Main.SettingsNotice.cs says why).
    // ###########################################################################################
    [Fact]
    public void An_unreadable_settings_file_is_said_on_its_own_card_with_where_it_was_kept_until_closed()
    {
        this.RedirectWorklogToTemp();
        string path = this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json");
        File.WriteAllText(path, "{ \"region\": ");
        UserSettings.LoadFrom(path);

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var timers = new NoticeTimers();
            main.ScheduleNoticeTimerOverrideForTests = timers.Schedule;
            var card = main.GetControl<Border>("SettingsUnreadableBanner");

            Assert.False(card.IsVisible);

            main.ShowSettingsUnreadableNoticeIfNeeded();

            Assert.True(card.IsVisible);
            Assert.Equal(0, timers.HideCount);
            string text = NoticeEmphasis.TextOf(main.GetControl<TextBlock>("SettingsUnreadableBannerText"));
            Assert.StartsWith("Your settings could not be read, so CRT started with default settings.", text, StringComparison.Ordinal);
            Assert.Contains(UserSettings.UnreadableSettings!.KeptAs!, text, StringComparison.Ordinal);

            RaiseClick(main.GetControl<Button>("SettingsUnreadableBannerDismissButton"));

            Assert.False(card.IsVisible);
        });
    }

    // ###########################################################################################
    // The Configuration tab's catalogue tree is built when the tab is first SHOWN, not at launch
    // (code review, 2026-10-13): building it read every board workbook while CRT started, for a tab
    // most launches never open.
    // ###########################################################################################
    [Fact]
    public void The_catalogue_tree_is_built_when_the_Configuration_tab_is_first_shown()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            main.Show();

            try
            {
                var configuration = main.GetControl<TabConfiguration>("TabConfiguration");
                main.UpdateLayout();

                Assert.False(configuration.IsCatalogueTreeRequestedForTests);

                main.GetControl<TabControl>("MainTabControl").SelectedItem = main.GetControl<TabItem>("ConfigurationTabItem");
                main.UpdateLayout();

                Assert.True(configuration.IsCatalogueTreeRequestedForTests);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // The component filter waits for the typing to stop (code review, 2026-10-13): each keystroke
    // rebuilt the list, every highlight and the KiCad overlay. A burst of keystrokes is one pending
    // filter; the label editor's direct call is applied at once and leaves nothing pending.
    // ###########################################################################################
    [Fact]
    public void Typing_in_the_component_filter_waits_for_the_last_keystroke()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var box = main.GetControl<TextBox>("ComponentSearchTextBox");

            box.Text = "U";
            box.Text = "U1";
            box.Text = "U12";

            // TextChanged is raised from the dispatcher's queue.
            Avalonia.Threading.Dispatcher.UIThread.RunJobs();

            Assert.True(main.IsComponentSearchPendingForTests);

            main.FlushComponentSearchForTests();
            Assert.False(main.IsComponentSearchPendingForTests);

            main.OnComponentSearchTextChanged(null, null!);
            Assert.False(main.IsComponentSearchPendingForTests);
        });
    }

    // A settings file read as it should be shows no such card.
    [Fact]
    public void A_readable_settings_file_shows_no_settings_card()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            main.ShowSettingsUnreadableNoticeIfNeeded();

            Assert.False(main.GetControl<Border>("SettingsUnreadableBanner").IsVisible);
        });
    }

    // The stack holds exactly the notices above, so the shift test covers every card - a new one
    // has to be added to NoticeNames - and each starts hidden: one showing at construction would
    // mean something ran that should not have.
    [Fact]
    public void Every_notice_is_in_the_stack_and_starts_hidden()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var stack = main.GetControl<StackPanel>("NoticeStack");

            Assert.Equal(NoticeNames, stack.Children.Select(card => card.Name ?? string.Empty));
            Assert.All(stack.Children, card => Assert.False(card.IsVisible, $"{card.Name} is showing"));

            // No background: the gaps between cards, and an empty stack, let clicks through.
            Assert.Null(stack.Background);
        });
    }

    // Nothing but the worklog bar and the tabs takes a row of the right panel. A notice given a
    // row of its own above the tabs is exactly what moved the window every time it came or went.
    [Fact]
    public void The_right_panel_has_a_row_for_the_worklog_bar_and_one_for_the_tabs_and_nothing_else()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var grid = Assert.IsType<Grid>(main.GetControl<Border>("RightPanel").Child);

            Assert.Equal(2, grid.RowDefinitions.Count);
            Assert.Equal(GridUnitType.Auto, grid.RowDefinitions[0].Height.GridUnitType);
            Assert.Equal(GridUnitType.Star, grid.RowDefinitions[1].Height.GridUnitType);

            Assert.Equal(0, Grid.GetRow(main.GetControl<Border>("WorklogBar")));
            Assert.All(
                grid.Children.Where(child => child.Name != "WorklogBar"),
                child => Assert.Equal(1, Grid.GetRow(child)));
        });
    }

    // ###########################################################################################
    // A notice that has just appeared lets clicks through for a moment.
    //
    // It appears on its own, at a moment nobody chose, possibly right under a pointer about to
    // click a thumbnail - and that click would land on its Install or close button instead. The
    // tooltips taught the same lesson. The release is run by hand here: a headless test's
    // DispatcherTimer never ticks.
    // ###########################################################################################
    [Fact]
    public void A_notice_that_has_just_appeared_lets_clicks_through_until_its_guard_ends()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var releases = new List<(Action Release, TimeSpan After)>();
            main.ScheduleNoticeTimerOverrideForTests = (release, after) =>
            {
                releases.Add((release, after));
                return new CancelOnDispose();
            };

            var update = main.GetControl<Border>("UpdateBanner");
            update.IsVisible = true;

            Assert.False(update.IsHitTestVisible);
            Assert.Equal(CRT.Main.NoticeClickGuard, Assert.Single(releases).After);

            releases[0].Release();

            Assert.True(update.IsHitTestVisible);
        });
    }

    // Hidden and shown again, a card starts its guard over - and the first showing's release,
    // already queued, must not end the second guard early.
    [Fact]
    public void A_notice_shown_again_starts_its_guard_over()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var releases = new List<Action>();
            var cancelled = new List<CancelOnDispose>();
            main.ScheduleNoticeTimerOverrideForTests = (release, _) =>
            {
                releases.Add(release);
                var handle = new CancelOnDispose();
                cancelled.Add(handle);
                return handle;
            };

            var sync = main.GetControl<Border>("SyncBanner");
            sync.IsVisible = true;
            sync.IsVisible = false;
            sync.IsVisible = true;

            Assert.Equal(2, releases.Count);
            Assert.True(cancelled[0].IsDisposed);

            // The first release slipped through anyway - as one the dispatcher had already queued.
            releases[0]();
            Assert.False(sync.IsHitTestVisible);

            releases[1]();
            Assert.True(sync.IsHitTestVisible);
        });
    }

    // Each card has its own guard: one appearing does not lock another that is already usable.
    [Fact]
    public void A_notice_appearing_leaves_the_cards_already_showing_clickable()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var releases = new List<Action>();
            main.ScheduleNoticeTimerOverrideForTests = (release, _) =>
            {
                releases.Add(release);
                return new CancelOnDispose();
            };

            var update = main.GetControl<Border>("UpdateBanner");
            update.IsVisible = true;
            releases[0]();

            main.GetControl<Border>("SourceSwitchBanner").IsVisible = true;

            Assert.True(update.IsHitTestVisible);
            Assert.False(main.GetControl<Border>("SourceSwitchBanner").IsHitTestVisible);
        });
    }

    private sealed class CancelOnDispose : IDisposable
    {
        public bool IsDisposed { get; private set; }

        public void Dispose() => this.IsDisposed = true;
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
                    var thumbnailPanel = schematics.GetControl<SchematicsThumbnailGallery>("InlineThumbnailGallery");
                    var checkBox = main.GetControl<TabConfiguration>("TabConfiguration")
                        .GetControl<CheckBox>("DetachSchematicsThumbnailsCheckBox");

                    UserSettings.DetachSchematicsThumbnails = true;
                    main.ApplyThumbnailsDetachedState();

                    Assert.True(main.IsThumbnailsWindowOpenForTests);
                    Assert.False(thumbnailPanel.IsVisible);

                    // Not shutting down - the user pressed the window's own close button.
                    main.CloseThumbnailsWindowAsUserForTests();

                    // The thumbnails are back in the tab...
                    Assert.True(thumbnailPanel.IsVisible);
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

    // ---------------------------------------------------------------------------------------
    // Update banner
    // ---------------------------------------------------------------------------------------

    // ###########################################################################################
    // A FAILED update install must hand the banner back to the user.
    //
    // A successful install never returns - ApplyUpdatesAndRestart replaces the process - so every
    // path that reaches the code after the await is a failure. The banner was left reading
    // "Downloading update..." with all three buttons disabled: no retry, no dismiss, and nothing
    // saying anything had gone wrong. DownloadAndInstallAsync reports failure by RETURNING FALSE
    // (it catches its own exceptions), so ignoring that return was the whole bug.
    //
    // Driven through the test seam rather than the click handler, because the handler reaches
    // GitHub over the network - which no test may do.
    // ###########################################################################################
    [Fact]
    public void A_failed_update_install_re_enables_the_banner_buttons_and_says_so()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();

            var install = main.GetControl<Button>("UpdateBannerInstallButton");
            var viewNotes = main.GetControl<Button>("UpdateBannerViewNotesButton");
            var dismiss = main.GetControl<Button>("UpdateBannerDismissButton");
            var text = main.GetControl<TextBlock>("UpdateBannerText");

            // The state the click handler leaves behind while the download is in flight.
            install.IsEnabled = false;
            viewNotes.IsEnabled = false;
            dismiss.IsEnabled = false;
            NoticeEmphasis.Show(text, "Downloading update...");

            main.RestoreUpdateBannerAfterFailedInstallForTests();

            // All three actionable again - the dead end is what made this a bug rather than a
            // cosmetic problem.
            Assert.True(install.IsEnabled);
            Assert.True(viewNotes.IsEnabled);
            Assert.True(dismiss.IsEnabled);

            // And it no longer claims to still be downloading.
            Assert.DoesNotContain("Downloading", NoticeEmphasis.TextOf(text));
            Assert.Contains("could not be installed", NoticeEmphasis.TextOf(text));
        });
    }

    // Raises a Button's Click the way a real press does. The handlers under test are private
    // event handlers wired in markup, so there is nothing else to call.
    private static void RaiseClick(Button button)
    {
        button.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
    }

    // ------------------------------------------------------------------ landing on the Drafts tab

    // ###########################################################################################
    // *** CREATING A BOARD LANDS THE USER ON THE DRAFTS TAB (owner request, 2026-09-24). ***
    //
    // Creating a board is never the end of a task - schematic images, KiCad data and component
    // rows all still have to be added, and every one of those actions lives on the Drafts tab.
    // OpenNewBoardWindowAsync used to select the new board and stop there, leaving the user on
    // whichever tab they started from (usually Contribute, which has nothing more to offer) with
    // no sign that a new tab had appeared to hold their work.
    //
    // The dialog itself cannot be driven headlessly, so these cover the navigation step the flow
    // ends with rather than the flow as a whole.
    // ###########################################################################################
    [Fact]
    public void Switching_to_the_drafts_tab_selects_it_when_it_is_visible()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var tabControl = main.GetControl<TabControl>("MainTabControl");
            var draftsTab = main.GetControl<TabItem>("DraftsTabItem");

            // The tab is hidden until a draft exists, so show it the way the real flow does.
            draftsTab.IsVisible = true;

            var firstVisible = tabControl.Items.OfType<TabItem>().First(t => t.IsVisible);
            tabControl.SelectedItem = firstVisible;

            main.SwitchToDraftsTab();

            Assert.Same(draftsTab, tabControl.SelectedItem);
        });
    }

    // ###########################################################################################
    // A HIDDEN tab cannot be selected, and trying would leave the tab control showing an empty
    // page. The caller has just created a draft so the tab will normally be up, but a silent
    // no-op is the right failure here - turning a cosmetic miss into a thrown exception would
    // fail the board creation that already succeeded.
    // ###########################################################################################
    [Fact]
    public void Switching_to_a_HIDDEN_drafts_tab_leaves_the_selection_alone()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var tabControl = main.GetControl<TabControl>("MainTabControl");
            var draftsTab = main.GetControl<TabItem>("DraftsTabItem");

            draftsTab.IsVisible = false;

            var firstVisible = tabControl.Items.OfType<TabItem>().First(t => t.IsVisible);
            tabControl.SelectedItem = firstVisible;

            main.SwitchToDraftsTab();

            Assert.Same(firstVisible, tabControl.SelectedItem);
        });
    }

    // Calling it when already on the Drafts tab is harmless - the guard against reassigning
    // SelectedItem, which would re-raise SelectionChanged for no reason.
    [Fact]
    public void Switching_to_the_drafts_tab_twice_is_harmless()
    {
        this.RedirectWorklogToTemp();
        this.RedirectSettingsToTemp();

        UiTest.Run(() =>
        {
            var main = new CRT.Main();
            var tabControl = main.GetControl<TabControl>("MainTabControl");
            var draftsTab = main.GetControl<TabItem>("DraftsTabItem");

            draftsTab.IsVisible = true;

            main.SwitchToDraftsTab();
            main.SwitchToDraftsTab();

            Assert.Same(draftsTab, tabControl.SelectedItem);
        });
    }
}
