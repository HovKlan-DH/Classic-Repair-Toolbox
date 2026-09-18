using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using Handlers.OnlineHandling;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT
{
    public partial class Main : Window
    {
        // Window placement: tracks the last known normal-state size and position
        private double _restoreWidth;
        private double _restoreHeight;
        private PixelPoint _restorePosition;
        private DispatcherTimer? _windowPlacementSaveTimer;
        private bool _windowPlacementReady = false;

        // Category filter: suppresses saves during programmatic selection changes
        private bool _suppressCategoryFilterSave;
        private bool _suppressComponentSearchRefresh;

        private BoardData? _currentBoardData;
        private bool _suppressComponentHighlightUpdate;
        private ComponentInfoWindow? _singleComponentInfoWindow;
        private readonly Dictionary<string, ComponentInfoWindow> _componentInfoWindowsByKey = new(StringComparer.OrdinalIgnoreCase);
        internal bool isHoveringComponent = false;
        private int _boardSelectionLoadVersion;

        // Blink selected highlights
        private DispatcherTimer? _blinkSelectedTimer;
        private bool _blinkSelectedPhaseVisible = true;
        private bool _blinkSelectedEnabled;
        private bool _isShowingDataSyncDisabledBanner;

        // Font Awesome spinner
        private DispatcherTimer? _dataSyncStatusIconSpinTimer;
        private int _dataSyncStatusIconSpinRequestCount;
        private double _dataSyncStatusIconSpinAngle;
        private bool _isHoveringDataSyncStatusIcon;

        private string _currentSyncFileRelativePath = string.Empty;

        // ###########################################################################################
        // The workbook id the "Show worklogs" checkbox is currently showing entries for (0 when
        // unchecked). RefreshWorklogBar compares this against the board's current active workbook
        // so switching boards (or the active workbook closing) drops the list view instead of
        // silently carrying it over to whatever workbook happens to be active now.
        // ###########################################################################################
        private int _worklogShowEntriesWorkbookId;

        // ###########################################################################################
        // Suppresses the "Show worklogs" preference save while RefreshWorklogBar seeds the checkbox
        // programmatically. Without it, seeding re-enters OnWorklogShowEntriesCheckedChanged and
        // persists the seeded value as if the user had clicked: selecting a board with no workbook
        // forces the box off, which would overwrite a saved "on" preference for every board and
        // every future session. Same pattern as _suppressCategoryFilterSave above.
        // ###########################################################################################
        private bool _suppressWorklogShowEntriesSave;

        // ###########################################################################################
        // Suppresses OnWorklogJobBoxSelectionChanged while RefreshWorklogBar seeds WorklogJobBox's
        // ItemsSource/SelectedItem programmatically. Without it, seeding the combo box for the newly
        // selected board re-enters ActivateWorkbook for whatever workbook happens to land in
        // SelectedItem - including firing it a second time for the very selection a user just made,
        // and firing it at all on every board switch even though nothing was clicked. Same pattern as
        // _suppressWorklogShowEntriesSave immediately above.
        // ###########################################################################################
        private bool _suppressWorklogJobBoxSave;

        // ###########################################################################################
        // The board OnHardwareSelectionChanged must select, instead of that hardware's saved last
        // board, for the one pass right after the worklog picker switches hardware to reach another
        // board's workbook.
        //
        // Without it that switch costs TWO full board loads: setting HardwareComboBox.SelectedItem
        // runs OnHardwareSelectionChanged synchronously, which picks the saved last board for the new
        // hardware and starts a complete OnBoardSelectionChanged for it (Excel, thumbnails, KiCad) -
        // and only then does the picker set the board it actually wanted, starting a second load. The
        // user sees the wrong board flash past, and the intermediate selection writes itself to
        // UserSettings.SetLastBoardForHardware as if they had chosen it.
        //
        // Cleared by OnHardwareSelectionChanged as soon as it has been honoured, so it can never
        // affect a later, unrelated hardware change.
        // ###########################################################################################
        private string? _pendingBoardSelectionOverride;

        // Region toggle: local override, does not affect the global setting
        private string _localRegion = UserSettings.Region;

        public string LocalRegion => this._localRegion;
        public BoardData? CurrentBoardData => this._currentBoardData;
        private bool _suppressRegionToggle;

        // Cascading offset for multiple popups
        private int _popupCascadeOffset = 0;

        // Fullscreen
        private SchematicsFullscreenWindow? _schematicsFullscreenWindow;

        // Detached schematics thumbnails window
        private SchematicsThumbnailsWindow? _schematicsThumbnailsWindow;

        // Set while CloseThumbnailsWindowForReopen is closing the current detached window in order to
        // reopen a fresh one (a board change, or the per-board setting being toggled). Tells the
        // window's own Closed handler this is bookkeeping, not the user dismissing the window - so it
        // must not turn detached mode OFF (via SetThumbnailsDetached) for the board that window
        // belonged to, which is still meant to stay detached the next time it is shown.
        //
        // Held as the WINDOW INSTANCE being closed rather than as a bool, and captured into that
        // window's Closed closure, so the signal cannot depend on Closed firing before the clearing
        // code runs. A bool cleared in a finally block assumed Close() raises Closed synchronously;
        // if a platform ever deferred it, the handler would read a cleared flag and silently turn the
        // user's detach setting off on an ordinary board change.
        private SchematicsThumbnailsWindow? _thumbnailsWindowClosingForReopen;

        // True while the schematics fullscreen window is being closed. While fullscreen is open it
        // OWNS the detached thumbnails window (see ReownSchematicsThumbnailsWindow), and closing an
        // owner closes what it owns - so the thumbnails window's Closed handler must not read that
        // cascade as the user dismissing it and turn the detach setting off. See
        // ReopenThumbnailsWindowAfterFullscreenClosed, which brings it back re-owned by this window.
        private bool _isClosingFullscreenWindow;

        // True from the moment this window starts closing. Lets owned windows tell "the user closed
        // me" apart from "the application is exiting" - see OnWindowClosing.
        private bool _isApplicationShuttingDown;

        // Lets a headless test put the window into the shutting-down state without actually closing
        // it, which would tear the test's own window down. Set by OnWindowClosing in the real app.
        internal bool IsApplicationShuttingDownForTests
        {
            get => this._isApplicationShuttingDown;
            set => this._isApplicationShuttingDown = value;
        }

        // Whether the detached thumbnails window is currently open, for tests that need to assert
        // the open/close half of ApplyThumbnailsDetachedState without reaching into the field.
        internal bool IsThumbnailsWindowOpenForTests => this._schematicsThumbnailsWindow != null;

        // Closes the detached thumbnails window the way the user's own OS close button does, so a
        // test can drive the Closed handler's "did the user dismiss it" branch.
        internal void CloseThumbnailsWindowAsUserForTests() => this._schematicsThumbnailsWindow?.Close();

        // Stands in for the hardware/board dropdown selection, so a headless test can drive the
        // PER-BOARD half of the detached-thumbnails feature. The same override-then-real pattern
        // TabWorkbooks.BoardKeyOverrideForTests uses, and for the same reason: selecting a board for
        // real means populating the dropdowns, which raises OnBoardSelectionChanged and loads the
        // board's Excel data off disk. null (the default) leaves GetCurrentBoardKey reading the real
        // dropdowns exactly as the shipped app does.
        internal string? BoardKeyOverrideForTests { get; set; }

        // Exposes the Oscilloscope tab control so other UI windows can route scope actions through it.
        public TabOscilloscope TabOscilloscopeControl => this.TabOscilloscope;

        private readonly TaskCompletionSource<bool> thisWindowOpenedCompletionSource =
    new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly TaskCompletionSource<bool> thisBackgroundStartupSyncCompletionSource =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private Task? thisBackgroundDataValidationTask;
        private bool thisHasScheduledOrphanAndUnusedFileCleanup;

        public Main()
        {
            InitializeComponent();

            this.TabSchematicsControl.Initialize(this);
            this.TabOverview.Initialize(this);
            this.TabContribute.Initialize(this);
            this.TabWorkbooks.Initialize(this);
            this.TabConfiguration.Initialize(this);

            this.MainTabControl.SelectionChanged += this.OnMainTabControlSelectionChanged;

            this.ApplyOscilloscopeTabVisibility();

            // Refreshes the worklog surfaces itself on the enable side, so no separate
            // RefreshWorklogBar() call belongs here. There used to be one, and by the time
            // TabWorkbooks.Initialize above had handed the tab its MainWindow it was no longer
            // cheap - it reached ReadAllWorkbooks (a directory scan plus a JSON parse per workbook)
            // and GetEntries synchronously, inside the constructor, before first paint.
            //
            // Nothing is lost by not refreshing here: the hardware/board combos are not populated
            // until PopulateHardwareDropDown, which now runs in StartAsync rather than later in this
            // constructor, so GetCurrentBoardKey returns empty at this point and there is no board
            // whose workbooks could be shown. That population raises OnBoardSelectionChanged, which
            // refreshes with a real board key.
            this.ApplyWorklogBarVisibility();

            // Only collapses the inline thumbnail column here if the setting is already on -
            // opening the detached window itself is deferred to OnWindowFirstOpened, since
            // Window.Show requires its owner to already be visible, which this window is not yet
            // at constructor time.
            if (UserSettings.DetachSchematicsThumbnails)
            {
                this.TabSchematicsControl.EnterThumbnailsDetachedMode();
            }

            // Restore left panel width from settings
            this.RootGrid.ColumnDefinitions[0].Width = new GridLength(UserSettings.LeftPanelWidth);
            this.RootGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);

            // The mode hint clears on the first press anywhere. Tunnel, so it runs before the press
            // reaches the control that was clicked and cannot be swallowed by one that marks the
            // event handled - the schematic image does exactly that when an area drag begins,
            // which is the most likely first click after the hint appears.
            this.AddHandler(
                InputElement.PointerPressedEvent,
                this.OnModeHintDismissPointerPressed,
                RoutingStrategies.Tunnel);

            // Subscribe to splitter pointer-release to persist positions when a drag ends.
            // handledEventsToo: true is required because GridSplitter marks the event as handled.
            this.MainSplitter.AddHandler(
                InputElement.PointerReleasedEvent,
                this.OnMainSplitterPointerReleased,
                RoutingStrategies.Bubble,
                handledEventsToo: true);

            // Initialize restore values from settings, then apply window placement before Show()
            // so Normal windows appear at the right place/size with zero flicker.
            // Maximized windows are positioned on the saved screen before being maximized so the
            // OS maximizes them on the correct monitor.
            this._restoreWidth = Math.Max(this.MinWidth, UserSettings.WindowWidth);
            this._restoreHeight = Math.Max(this.MinHeight, UserSettings.WindowHeight);
            this._restorePosition = new PixelPoint(UserSettings.WindowX, UserSettings.WindowY);

            // Wireup "blink" button
            this.BlinkSelectedCheckBox.IsChecked = UserSettings.BlinkSelected;

            if (UserSettings.HasWindowPlacement)
            {
                this.WindowStartupLocation = WindowStartupLocation.Manual;
                this.Width = this._restoreWidth;
                this.Height = this._restoreHeight;

                if (UserSettings.WindowState == nameof(Avalonia.Controls.WindowState.Maximized))
                {
                    // Place anywhere on the saved screen so the OS maximizes it there
                    this.Position = new PixelPoint(UserSettings.WindowScreenX + 100, UserSettings.WindowScreenY + 100);
                    this.WindowState = Avalonia.Controls.WindowState.Maximized;
                }
                else
                {
                    this.Position = this._restorePosition;
                }
            }

            this.Opened += this.OnWindowFirstOpened;
            this.Closing += this.OnWindowClosing;
            this.Closed += this.OnWindowClosed;

            this.UpdateRegionButtonsState();
            this.HardwareComboBox.SelectionChanged += this.OnHardwareSelectionChanged;
            this.BoardComboBox.SelectionChanged += this.OnBoardSelectionChanged;
            this.CategoryFilterListBox.SelectionChanged += this.OnCategoryFilterSelectionChanged;
            this.ComponentFilterListBox.SelectionChanged += this.OnComponentFilterSelectionChanged;

            // Renders each workbook as "#{Id} | {Title} (hardware/board)", e.g. "#3 | Black screen
            // (C64/250469)" - the short folder-derived label from FormatBoardKeyForDisplay, not the
            // full Excel-sheet names ("Commodore 64" / "250469 (short board)"), which read fine as a
            // combo box's own dedicated row but ran too long once several workbooks from different
            // boards sit side by side in one dropdown. The board suffix itself is what the plain-text
            // box never needed - it only ever showed the current board's own workbooks - but the
            // picker now lists EVERY workbook on every board (see RefreshWorklogBar), so two workbooks
            // that happen to share a title are otherwise indistinguishable. Built here in code rather
            // than as an inline XAML DataTemplate bound to a formatted property, since WorkbookRecord
            // is a plain JSON-backed model (see WorklogManager.cs) with no display-formatting concept
            // of its own.
            this.WorklogJobBox.ItemTemplate = new FuncDataTemplate<WorkbookRecord>(
                (workbook, _) => new TextBlock
                {
                    Text = workbook == null
                        ? string.Empty
                        : $"#{workbook.Id} | {workbook.Title} ({FormatBoardKeyForDisplay(workbook.BoardKey)})",
                    FontSize = 12,
                });

            var versionString = AppConfig.AppDisplayVersionString;
            var assembly = Assembly.GetExecutingAssembly();

            this.PopulateAboutTab(assembly, versionString);

            this.Title = versionString != "0.0.0"
                ? $"Classic Repair Toolbox {versionString}"
                : "Classic Repair Toolbox";

            this.AddHandler(
                InputElement.PointerPressedEvent,
                this.OnMainPointerPressedCloseSinglePopup,
                RoutingStrategies.Bubble,
                handledEventsToo: true
            );

            this.AddHandler(
                InputElement.KeyDownEvent,
                this.OnMainKeyDownCloseSinglePopup,
                RoutingStrategies.Tunnel,
                handledEventsToo: true
            );

            this.AddHandler(
                InputElement.PointerReleasedEvent,
                (s, e) =>
            {
                Dispatcher.UIThread.Post(() =>
                {
                    // Abort stealing focus if another window (like the component popup) is currently active
                    if (!this.IsActive)
                    {
                        return;
                    }

                    // Do not steal focus while the schematics label editor is active
                    if (this.TabSchematicsControl.IsLabelEditorActive)
                    {
                        return;
                    }

                    // Do not steal focus if we are on tabs that utilize text inputs
                    var selectedTab = this.MainTabControl?.SelectedItem as TabItem;
                    string? tabHeader = selectedTab?.Header?.ToString();

                    // The Workbooks tab has its own priority focus ("Find a previous repair"),
                    // set on tab entry by OnMainTabControlSelectionChanged - see FocusSearchBox's
                    // comment for why the global steal has to back off here, not just once but on
                    // every pointer release while this tab is showing.
                    if (tabHeader == "Feedback" || tabHeader == "Configuration" || tabHeader == "Workbooks")
                    {
                        return;
                    }

                    // Avoid stealing focus if another TextBox currently holds it naturally
                    var focusedElement = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
                    if (focusedElement is global::Avalonia.Controls.TextBox && focusedElement != this.ComponentSearchTextBox)
                    {
                        return;
                    }

                    if (this.ComponentSearchTextBox != null && !this.ComponentSearchTextBox.IsFocused)
                    {
                        this.ComponentSearchTextBox.Focus();
                    }
                }, DispatcherPriority.Background);
            },
            RoutingStrategies.Bubble,
            handledEventsToo: true
        );

            UserSettings.CheckDataOnLaunchChanged += this.OnCheckDataOnLaunchSettingChanged;
            UserSettings.WorkbooksScopeChanged += this.OnWorkbooksScopeSettingChanged;
            UserSettings.WorklogCurrencyChanged += this.OnWorklogCurrencySettingChanged;
            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Everything the window does that reaches OUTSIDE itself: the hardware/board data it reads
        // from DataManager, the update check (real HTTP) and the background data sync (network).
        //
        // These deliberately do NOT run from the constructor. Constructing Main used to mean hitting
        // the network and DataManager's static state, so no test could build the window at all - the
        // largest file in the app sat at zero coverage purely because of where these three calls
        // lived. Splitting construction from startup is the same shape as HeadlessTestApp, which
        // inherits the real App and skips its OnFrameworkInitializationCompleted for the same reason.
        //
        // App.OnFrameworkInitializationCompleted calls this BEFORE Show(), which is what keeps the
        // running app behaving as before: the constructor used to populate the combos, so the window
        // was fully populated by the time it was first painted. Calling this after Show() instead
        // would paint an empty hardware/board dropdown and an empty component list for the duration
        // of the first board load, and would let OnWindowFirstOpened run before that load rather
        // than after it. Tests construct Main and never call this.
        //
        // Fire-and-forget by design (the caller does not await it): StartBackgroundSyncAsync was
        // already an async void started from the constructor, so awaiting it here would hold up the
        // splash close and change startup timing, which App logs as a StartupTimeline milestone.
        //
        // Which is also why this catches: the caller discards the returned Task, so without a catch
        // an exception thrown before the first await - PopulateHardwareDropDown cascades
        // synchronously into the whole first board load - would fault a Task nobody observes.
        // TaskScheduler.UnobservedTaskException only fires when that Task is finalised, so the
        // report would be arbitrarily late or never arrive at all; the old async void path reached
        // the logger immediately. Logging here restores that.
        // ###########################################################################################
        internal async Task StartAsync()
        {
            try
            {
                if (DataManager.DataUpdateRequiresAppUpdate)
                {
                    this.ShowMainExcelRequiresAppUpdateBanner();
                }

                // Populates the hardware combo box, whose SelectionChanged cascades synchronously into
                // OnBoardSelectionChanged and so triggers the first board load. This ran in the
                // constructor before the WorklogJobBox.ItemTemplate above was assigned; running it here
                // means that template is always in place before RefreshWorklogBar can populate the box.
                this.PopulateHardwareDropDown();

                if (UserSettings.CheckVersionOnLaunch)
                {
                    _ = this.CheckForAppUpdateNowAsync();
                }

                await this.StartBackgroundSyncAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Main.StartAsync failed: {ex}");
            }
        }

        // ###########################################################################################
        // Checks for an available update and refreshes the application update banner immediately.
        // This does not interfere with the separate main-Excel compatibility warning banner.
        // ###########################################################################################
        internal async Task CheckForAppUpdateNowAsync()
        {
            bool? available = await UpdateService.CheckForUpdateAsync();

            if (available == true)
            {
                this.ShowApplicationUpdateAvailableBanner();
                return;
            }

            this.HideApplicationUpdateBanner();
        }

        // ###########################################################################################
        // Shows the dedicated dismissable warning banner explaining that newer main Excel data
        // exists but requires a newer application version before it can be used.
        // ###########################################################################################
        private void ShowMainExcelRequiresAppUpdateBanner()
        {
            this.MainExcelRequiresAppUpdateBannerText.Text =
                "Newer main Excel data file is available, but requires a newer application version, due to breaking changes - please update the application. No more data updates will be given for this application version, and worst-case is that future data update will break UI or functionality. Consider yourself informed 😁";
            this.MainExcelRequiresAppUpdateBannerDismissButton.IsEnabled = true;
            this.MainExcelRequiresAppUpdateBanner.IsVisible = true;
        }

        // ###########################################################################################
        // Hides the dedicated main-Excel compatibility warning banner and clears its persisted
        // visible state for the current session.
        // ###########################################################################################
        private void HideMainExcelRequiresAppUpdateBanner()
        {
            this.MainExcelRequiresAppUpdateBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Shows the normal application update banner when GitHub reports that a newer installed
        // application package is available for download.
        // ###########################################################################################
        private void ShowApplicationUpdateAvailableBanner()
        {
            // The "(simulated)" suffix is the only thing distinguishing a faked update from a real
            // one on screen, and screenshots are what arrive in bug reports.
            string simulatedSuffix = SimulationOptions.Current.SimulateUpdate ? " (simulated)" : string.Empty;
            this.UpdateBannerText.Text = $"Version [{UpdateService.PendingVersion}] is available{simulatedSuffix}";
            this.UpdateBannerInstallButton.IsVisible = true;
            this.UpdateBannerViewNotesButton.IsVisible = true;
            this.UpdateBannerInstallButton.IsEnabled = true;
            this.UpdateBannerViewNotesButton.IsEnabled = true;
            this.UpdateBannerDismissButton.IsEnabled = true;
            this.UpdateBanner.IsVisible = true;
        }

        // ###########################################################################################
        // Hides the normal application update banner without affecting the separate main-Excel
        // compatibility warning banner.
        // ###########################################################################################
        private void HideApplicationUpdateBanner()
        {
            this.UpdateBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the dedicated main-Excel compatibility warning banner.
        // ###########################################################################################
        private void OnMainExcelRequiresAppUpdateBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.HideMainExcelRequiresAppUpdateBanner();
        }

        // ###########################################################################################
        // Dismisses the normal application update banner without cancelling the update.
        // ###########################################################################################
        private void OnUpdateBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.HideApplicationUpdateBanner();
        }

        // ###########################################################################################
        // Performs the manual UI data update check as a single visible sync flow.
        // Startup sync behavior remains unchanged and is handled separately during application launch.
        // keepBannerTextStatic: when true, the banner keeps its initial text during the full refresh.
        // ###########################################################################################
        internal async Task CheckForDataUpdatesNowAsync(bool keepBannerTextStatic = false)
        {
            this._currentSyncFileRelativePath = string.Empty;

            this.SetSyncBannerText($"Checking data from {AppConfig.GetOnlineSourceLabel()} - please wait...");
            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBanner.IsVisible = true;

            this.StartDataSyncStatusIconSpin();

            try
            {
                var syncResult = await DataManager.CheckForDataUpdatesNowAsync(
                    status => Dispatcher.UIThread.Post(() =>
                    {
                        if (keepBannerTextStatic)
                        {
                            return;
                        }

                        if (status.Contains("up to date", StringComparison.OrdinalIgnoreCase) ||
                            status.StartsWith("Update complete", StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }

                        // IMPORTANT: status only on line 1 (file path is line 2, managed separately)
                        this.SetSyncBannerText(status);
                    }),
                    filePath => Dispatcher.UIThread.Post(() =>
                    {
                        // IMPORTANT: file path only on line 2
                        this._currentSyncFileRelativePath = filePath ?? string.Empty;

                        if (keepBannerTextStatic)
                        {
                            return;
                        }

                        // Keep whatever line 1 currently says; only refresh line 2.
                        string currentLine1 = this.SyncBannerText.Text?
                            .Split('\n')[0]
                            .Trim() ?? string.Empty;

                        this.SetSyncBannerText(currentLine1);
                    }));

                if (syncResult.ChangedCount < 0)
                {
                    return;
                }

                if (syncResult.MainExcelChanged)
                {
                    this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();
                }

                // Clear file line once we switch to the final summary.
                this._currentSyncFileRelativePath = string.Empty;

                if (syncResult.ChangedCount > 0)
                {
                    string bannerText = syncResult.ChangedCount == 1
                        ? "[1] file updated - please refresh board"
                        : $"[{syncResult.ChangedCount}] files updated - please refresh board";

                    this.SetSyncBannerText(ComponentListBuilder.BuildSyncBannerText(bannerText, syncResult.ProtectedFilesCount));
                    this.SyncBannerRefreshButton.IsVisible = true;
                    this.SyncBanner.IsVisible = true;
                }
                else if (syncResult.ProtectedFilesCount > 0)
                {
                    this.SetSyncBannerText(ComponentListBuilder.BuildSyncBannerText("All data files are up to date", syncResult.ProtectedFilesCount));
                    this.SyncBannerRefreshButton.IsVisible = false;
                    this.SyncBanner.IsVisible = true;
                }
                else
                {
                    this.SyncBanner.IsVisible = false;
                }

                if (DataManager.DataUpdateRequiresAppUpdate)
                {
                    this.ShowMainExcelRequiresAppUpdateBanner();
                }

                if (UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    _ = DataManager.DeleteOrphanAndUnusedFilesAsync();
                }
            }
            finally
            {
                this._currentSyncFileRelativePath = string.Empty;
                this.StopDataSyncStatusIconSpin();
            }
        }

        // ###########################################################################################
        // Syncs any remaining non-Excel files and returns the number of files that changed.
        // keepBannerTextStatic: when true, intermediate status text is suppressed during the run.
        // ###########################################################################################
        private async Task<int> SyncRemainingFilesAsync(bool keepBannerTextStatic = false)
        {
            if (!DataManager.HasPendingSync)
            {
                return 0;
            }

            return await DataManager.SyncRemainingAsync(
                status => Dispatcher.UIThread.Post(() =>
                {
                    if (keepBannerTextStatic)
                    {
                        return;
                    }

                    if (status.Contains("up to date", StringComparison.OrdinalIgnoreCase) ||
                        status.StartsWith("Sync complete", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    // status only on line 1
                    this.SetSyncBannerText(status);
                }),
                filePath => Dispatcher.UIThread.Post(() =>
                {
                    // file only on line 2
                    this._currentSyncFileRelativePath = filePath ?? string.Empty;

                    if (keepBannerTextStatic)
                    {
                        return;
                    }

                    string currentLine1 = this.SyncBannerText.Text?
                        .Split('\n')[0]
                        .Trim() ?? string.Empty;

                    this.SetSyncBannerText(currentLine1);
                }));
        }

        // ###########################################################################################
        // Starts the remaining background sync without blocking the caller.
        //
        // Returns a Task rather than being async void so StartAsync can compose it. It still swallows
        // its own exceptions in the catch below, so nothing faults out of here either way - the
        // signature change is about being awaitable, not about changing failure behaviour.
        // ###########################################################################################
        private async Task StartBackgroundSyncAsync(bool keepBannerTextStatic = false)
        {
            try
            {
                if (!UserSettings.CheckDataOnLaunch || !DataManager.HasPendingSync)
                {
                    if (!this._isShowingDataSyncDisabledBanner)
                    {
                        this.SyncBanner.IsVisible = false;
                        this.SyncBannerRefreshButton.IsVisible = false;
                    }

                    return;
                }

                this.SyncBannerText.Text = $"Checking data from {AppConfig.GetOnlineSourceLabel()} - please wait...";
                this.SyncBannerRefreshButton.IsVisible = false;
                this.SyncBanner.IsVisible = true;

                this.StartDataSyncStatusIconSpin();

                int changed = await this.SyncRemainingFilesAsync(keepBannerTextStatic);

                DataManager.LoadProtectedContributionStateForCurrentData();
                int protectedFilesCount = DataManager.ProtectedContributionFileCount;

                if (changed > 0)
                {
                    string bannerText = changed == 1
                        ? "[1] file updated in background - please refresh board"
                        : $"[{changed}] files updated in background - please refresh board";

                    this.SyncBannerText.Text = ComponentListBuilder.BuildSyncBannerText(bannerText, protectedFilesCount);
                    this.SyncBannerRefreshButton.IsVisible = true;
                    this.SyncBanner.IsVisible = true;
                }
                else if (protectedFilesCount > 0)
                {
                    this.SyncBannerText.Text = ComponentListBuilder.BuildSyncBannerText("All data files are up to date", protectedFilesCount);
                    this.SyncBannerRefreshButton.IsVisible = false;
                    this.SyncBanner.IsVisible = true;
                }
                else
                {
                    this.SyncBanner.IsVisible = false;
                }

                if (UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    _ = DataManager.DeleteOrphanAndUnusedFilesAsync();
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Background data sync failed - [{ex.Message}]");
            }
            finally
            {
                this.thisBackgroundStartupSyncCompletionSource.TrySetResult(true);
                this.StopDataSyncStatusIconSpin();
            }
        }

        // ###########################################################################################
        // Manually reloads the current board configuration.
        // ###########################################################################################
        private void OnRefreshBoardClick(object? sender, RoutedEventArgs e)
        {
            this.SyncBanner.IsVisible = false;
            this.SyncBannerRefreshButton.IsVisible = false;
            this.ReloadCurrentBoardFromDisk(string.Empty);
        }

        // ###########################################################################################
        // Dismisses the sync banner.
        // ###########################################################################################
        private void OnSyncBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this._isShowingDataSyncDisabledBanner = false;
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the sync banner when clicking anywhere on it.
        // ###########################################################################################
        private void OnSyncBannerPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this._isShowingDataSyncDisabledBanner = false;
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Opens the GitHub release notes page for the pending update version.
        // ###########################################################################################
        private void OnViewReleaseNotesClick(object? sender, RoutedEventArgs e)
        {
            string version = UpdateService.PendingVersion ?? string.Empty;
            string url = string.IsNullOrWhiteSpace(version)
                ? $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}/releases"
                : $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}/releases/tag/{version}";
            OpenUrl(url);
        }

        // ###########################################################################################
        // Downloads and installs the pending update, showing progress in the banner text.
        // ###########################################################################################
        private async void OnInstallUpdateClick(object? sender, RoutedEventArgs e)
        {
            this.UpdateBannerInstallButton.IsEnabled = false;
            this.UpdateBannerViewNotesButton.IsEnabled = false;
            this.UpdateBannerDismissButton.IsEnabled = false;
            this.UpdateBannerText.Text = "Downloading update...";

            await UpdateService.DownloadAndInstallAsync(progress =>
            {
                Dispatcher.UIThread.Post(() => this.UpdateBannerText.Text = $"Downloading update: {progress}%");
            });
        }

        // ###########################################################################################
        // Populates the hardware drop-down with distinct hardware names from loaded data.
        // ###########################################################################################
        private void PopulateHardwareDropDown()
        {
            var hardwareNames = BuildVisibleHardwareNames(UserSettings.CatalogueUncheckedKeysSnapshot);

            this.HardwareComboBox.ItemsSource = hardwareNames;

            if (hardwareNames.Count == 0)
            {
                this.HardwareComboBox.SelectedIndex = -1;
                return;
            }

            var lastHardware = UserSettings.GetLastHardware();
            var savedIndex = hardwareNames.FindIndex(h =>
                string.Equals(h, lastHardware, StringComparison.OrdinalIgnoreCase));

            this.HardwareComboBox.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
        }

        // ###########################################################################################
        // Whether a hardware belongs in the Hardware drop-down: not itself unchecked, AND has at
        // least one board that is not itself unchecked. Schematic-level visibility is deliberately
        // NOT considered here - that would mean loading every board's Excel file just to populate a
        // drop-down, which the board dropdown/thumbnail filters already handle lazily per selection.
        // ###########################################################################################
        private static bool HasAnyVisibleBoard(IReadOnlySet<string> uncheckedKeys, string hardwareName)
        {
            if (!CatalogueVisibility.IsHardwareVisible(uncheckedKeys, hardwareName))
                return false;

            return DataManager.HardwareBoards.Any(entry =>
                string.Equals(entry.HardwareName, hardwareName, StringComparison.OrdinalIgnoreCase) &&
                CatalogueVisibility.IsBoardVisible(uncheckedKeys, hardwareName, entry.BoardName));
        }

        // ###########################################################################################
        // Rebuilds the hardware and board selectors after the main Excel data changed, while trying
        // to preserve the current selection when those entries still exist.
        // ###########################################################################################
        private void RefreshHardwareAndBoardSelectionsAfterMainExcelSync()
        {
            string previousHardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;
            string previousBoard = this.BoardComboBox.SelectedItem as string ?? string.Empty;

            this.PopulateHardwareDropDown();

            var hardwareNames = this.HardwareComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            if (hardwareNames.Count == 0)
            {
                return;
            }

            int hardwareIndex = hardwareNames.FindIndex(h =>
                string.Equals(h, previousHardware, StringComparison.OrdinalIgnoreCase));

            this.HardwareComboBox.SelectedIndex = hardwareIndex >= 0 ? hardwareIndex : 0;

            var boardNames = this.BoardComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            if (boardNames.Count == 0)
            {
                return;
            }

            int boardIndex = boardNames.FindIndex(b =>
                string.Equals(b, previousBoard, StringComparison.OrdinalIgnoreCase));

            this.BoardComboBox.SelectedIndex = boardIndex >= 0 ? boardIndex : 0;
        }

        // ###########################################################################################
        // Filters the board drop-down to only show boards belonging to the selected hardware.
        // ###########################################################################################
        private void OnHardwareSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this._suppressComponentSearchRefresh = true;
            this.ComponentSearchTextBox.Text = string.Empty;
            this._suppressComponentSearchRefresh = false;

            var selectedHardware = this.HardwareComboBox.SelectedItem as string;

            var boards = this.BuildVisibleBoardNamesForSelectedHardware(
                UserSettings.CatalogueUncheckedKeysSnapshot);

            this.BoardComboBox.ItemsSource = boards;

            if (string.IsNullOrWhiteSpace(selectedHardware) || boards.Count == 0)
            {
                this.BoardComboBox.SelectedIndex = -1;
                return;
            }

            UserSettings.SetLastHardware(selectedHardware);

            // Honoured once and cleared, so the caller that set it gets the board it asked for
            // rather than this hardware's saved last board - see the field's own comment for why
            // going through the saved board costs an entire extra board load.
            var pendingBoard = this._pendingBoardSelectionOverride;
            this._pendingBoardSelectionOverride = null;

            var targetBoard = pendingBoard ?? UserSettings.GetLastBoardForHardware(selectedHardware);
            var targetIndex = boards.FindIndex(b =>
                string.Equals(b, targetBoard, StringComparison.OrdinalIgnoreCase));

            this.BoardComboBox.SelectedIndex = targetIndex >= 0 ? targetIndex : 0;
        }

        // ###########################################################################################
        // Handles board selection changes and loads the visible board UI first, then starts heavier
        // schematic/KiCad work in the background so the window can remain responsive immediately.
        // ###########################################################################################
        private async void OnBoardSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            await this.LoadSelectedBoardAsync(clearComponentSearch: ReferenceEquals(sender, this.BoardComboBox));
        }

        // ###########################################################################################
        // The board load itself, separated from the SelectionChanged handler so it can also be
        // invoked where no selection actually changed - see ApplyCatalogueVisibility, which needs a
        // reload when a schematic under the CURRENT board is hidden or shown and re-selecting the
        // same index raises no event.
        //
        // clearComponentSearch carries what the handler used to read off its sender: the component
        // search box is emptied for a board change the user made in the drop-down, but not for a
        // load triggered some other way.
        // ###########################################################################################
        private async Task LoadSelectedBoardAsync(bool clearComponentSearch)
        {
            this.TabSchematicsControl.CancelWorklogEntryMode();

            // The Workbooks tab's search box is deliberately NOT cleared here. It used to be, on
            // the reasoning that a board change is a change of subject - but in "Show all workbooks"
            // scope, clicking a search result on another board IS how the user follows the search,
            // and that click switches the board, so clearing here wiped the query at the exact
            // moment it had done its job. See TabWorkbooks.ClearSearch, which is now the only thing
            // that clears it. (OnHardwareSelectionChanged still clears ComponentSearchTextBox, which
            // is a different box with no cross-board results to follow.)

            this._suppressCategoryFilterSave = true;
            int loadVersion = unchecked(++this._boardSelectionLoadVersion);

            if (clearComponentSearch)
            {
                this._suppressComponentSearchRefresh = true;
                this.ComponentSearchTextBox.Text = string.Empty;
                this._suppressComponentSearchRefresh = false;
            }

            foreach (var thumb in this.TabSchematicsControl.currentThumbnails)
            {
                if (!ReferenceEquals(thumb.ImageSource, thumb.BaseThumbnail))
                {
                    (thumb.ImageSource as IDisposable)?.Dispose();
                }

                (thumb.BaseThumbnail as IDisposable)?.Dispose();
            }

            this.TabSchematicsControl.currentThumbnails.Clear();
            this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList")!.ItemsSource = null;
            this.CategoryFilterListBox.ItemsSource = null;
            this.ComponentFilterListBox.ItemsSource = null;

            this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);
            this.TabSchematicsControl.schematicByName = new(StringComparer.OrdinalIgnoreCase);
            this.SetComponentHighlightRects(new(StringComparer.OrdinalIgnoreCase));

            this._currentBoardData = null;
            this.UpdateRegionButtonsState();
            this.PopulateBoardInfoSection(null, null);
            this.TabSchematicsControl.ResetSchematicsViewer();

            var selectedHardware = this.HardwareComboBox.SelectedItem as string;
            var selectedBoard = this.BoardComboBox.SelectedItem as string;

            if (string.IsNullOrEmpty(selectedHardware) || string.IsNullOrEmpty(selectedBoard))
            {
                // No board selected at all. _currentBoardData was cleared above, so refresh to the
                // empty state rather than leaving the bar and the Workbooks tab showing the board
                // that WAS selected a moment ago - stale worklog surfaces above a blank schematic
                // view read as the previous board still being loaded.
                this.RefreshWorklogBar();
                return;
            }

            UserSettings.SetLastHardware(selectedHardware);
            UserSettings.SetLastBoardForHardware(selectedHardware, selectedBoard);

            string boardKey = this.GetCurrentBoardKey();

            this.ApplySchematicsSplitterRatioForCurrentBoard(boardKey);
            this.ApplyThumbnailsDetachedStateForBoardChange();

            var entry = DataManager.HardwareBoards.FirstOrDefault(ent =>
                string.Equals(ent.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ent.BoardName, selectedBoard, StringComparison.OrdinalIgnoreCase));

            if (entry == null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                // Same reasoning as the no-selection case above: this board has no data file, so
                // there is nothing to load and _currentBoardData stays null. Refresh so the worklog
                // surfaces show THIS board (its workbook list does not need board data) with an
                // empty board pane, rather than the previous board's previews.
                this.RefreshWorklogBar();
                return;
            }

            var boardData = await DataManager.LoadBoardDataAsync(entry);

            // A SUPERSEDED load returns without touching anything: the user has selected another
            // board since, and that newer load owns every surface now - refreshing here would put
            // this board's worklog state under the newer board's header.
            if (loadVersion != this._boardSelectionLoadVersion)
            {
                return;
            }

            if (boardData == null)
            {
                // The board's Excel file is missing or unreadable. Still the currently selected
                // board, so its worklog surfaces must show IT (empty board pane, real workbook
                // list) rather than whatever the previously selected board left on screen.
                this.RefreshWorklogBar();
                return;
            }

            this._currentBoardData = boardData;
            this.UpdateRegionButtonsState();
            this.PopulateBoardInfoSection(boardData.RevisionDate, boardData.Credits);

            // AFTER _currentBoardData is assigned, deliberately - this call used to sit above the
            // await, before the board data existed. Everything board-data-dependent in
            // RefreshWorklogBar was wrong in that first pass, not just the board pane:
            // RefreshSelectedSchematicEntries reset the selected schematic and drew the placeholder,
            // and SetShowWorklogEntriesList re-seeded the Schematics overlay against a
            // just-blanked highlight cache. It was patched with a second, narrow board-pane refresh
            // here; running the whole refresh once, in the right place, fixes the class of bug
            // rather than the one instance, and removes a full board-pane rebuild per board switch.
            //
            // Nothing above the await needs it back: the splitter-ratio restore reads only boardKey.
            this.RefreshWorklogBar();

            var categories = ComponentListBuilder.BuildDistinctCategories(boardData);
            this.CategoryFilterListBox.ItemsSource = categories;

            var savedCategories = UserSettings.GetSelectedCategories(boardKey);
            if (savedCategories == null)
            {
                try
                {
                    this.CategoryFilterListBox.SelectAll();
                }
                catch (OutOfMemoryException ex)
                {
                    Logger.Debug(ex, "Failed to apply default category selection - group was too large to select");
                }
            }
            else
            {
                for (int i = 0; i < categories.Count; i++)
                {
                    if (savedCategories.Contains(categories[i], StringComparer.OrdinalIgnoreCase))
                    {
                        this.CategoryFilterListBox.Selection.Select(i);
                    }
                }
            }

            this._suppressCategoryFilterSave = false;

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(boardData, UserSettings.Region, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try
                {
                    this.ComponentFilterListBox.SelectAll();
                }
                catch
                {
                }
            }

            this._suppressComponentHighlightUpdate = false;

            _ = Task.Run(async () =>
            {
                try
                {
                    var highlightRects = await Task.Run(() =>
                        HighlightRectBuilder.BuildHighlightRects(boardData, UserSettings.Region));

                    var uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;
                    bool IsSchematicVisible(string schematicName) =>
                        selectedHardware != null && selectedBoard != null &&
                        CatalogueVisibility.IsSchematicVisible(uncheckedKeys, selectedHardware, selectedBoard, schematicName);

                    var schematicByName = boardData.Schematics
                        .Where(schematic => !string.IsNullOrWhiteSpace(schematic.SchematicName))
                        .Where(schematic => IsSchematicVisible(schematic.SchematicName))
                        .ToDictionary(
                            schematic => schematic.SchematicName,
                            schematic => schematic,
                            StringComparer.OrdinalIgnoreCase);

                    var loaded = await Task.Run(() =>
                    {
                        var result = new List<(string Name, string FullPath, Bitmap? FullBitmap)>();

                        foreach (var schematic in boardData.Schematics)
                        {
                            if (string.IsNullOrWhiteSpace(schematic.SchematicImageFile))
                            {
                                continue;
                            }

                            if (!IsSchematicVisible(schematic.SchematicName))
                            {
                                continue;
                            }

                            var fullPath = Path.Combine(
                                DataManager.DataRoot,
                                schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar));

                            Bitmap? bitmap = null;

                            if (File.Exists(fullPath))
                            {
                                try
                                {
                                    bitmap = new Bitmap(fullPath);
                                }
                                catch (Exception ex)
                                {
                                    Logger.Warning($"Could not load schematic image [{fullPath}] - [{ex.Message}]");
                                }
                            }

                            result.Add((schematic.SchematicName, fullPath, bitmap));
                        }

                        return result;
                    });

                    var thumbnails = new List<SchematicThumbnail>();

                    foreach (var (name, fullPath, fullBitmap) in loaded)
                    {
                        RenderTargetBitmap? baseThumbnail = null;
                        PixelSize originalPixelSize = default;

                        if (fullBitmap != null)
                        {
                            baseThumbnail = TabSchematics.CreateScaledThumbnail(fullBitmap, AppConfig.ThumbnailMaxWidth);
                            originalPixelSize = fullBitmap.PixelSize;
                            fullBitmap.Dispose();
                        }

                        thumbnails.Add(new SchematicThumbnail
                        {
                            Name = name,
                            ImageFilePath = fullPath,
                            BaseThumbnail = baseThumbnail,
                            OriginalPixelSize = originalPixelSize,
                            ImageSource = baseThumbnail,
                            VisualOpacity = 1.0,
                            IsMatchForSelection = false
                        });
                    }

                    List<string> rawPaths = this.GetCurrentBoardKiCadRawPaths();

                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        if (loadVersion != this._boardSelectionLoadVersion)
                        {
                            foreach (var thumbnail in thumbnails)
                            {
                                if (!ReferenceEquals(thumbnail.ImageSource, thumbnail.BaseThumbnail))
                                {
                                    (thumbnail.ImageSource as IDisposable)?.Dispose();
                                }

                                (thumbnail.BaseThumbnail as IDisposable)?.Dispose();
                            }

                            return;
                        }

                        // The schematic image is about to appear while the KiCad project is still
                        // loading behind it, so flag that wait as soon as the board has KiCad data.
                        this.TabSchematicsControl.SetKiCadInitializingIndicatorVisible(rawPaths.Count > 0);

                        this.SetComponentHighlightRects(highlightRects);
                        this.TabSchematicsControl.schematicByName = schematicByName;
                        this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

                        this.TabSchematicsControl.LoadSortedThumbnails(boardKey, thumbnails);

                        if (this.TabSchematicsControl.currentThumbnails.Count > 0)
                        {
                            string? savedSchematic = UserSettings.GetLastSchematicForBoard(boardKey);
                            var orderedThumbnails = this.TabSchematicsControl.currentThumbnails.ToList();

                            int savedIndex = string.IsNullOrEmpty(savedSchematic)
                                ? -1
                                : orderedThumbnails.FindIndex(thumbnail =>
                                    string.Equals(thumbnail.Name, savedSchematic, StringComparison.OrdinalIgnoreCase));

                            this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList")!.SelectedIndex =
                                savedIndex >= 0 ? savedIndex : 0;
                        }

                        var localFiles = boardData.BoardLocalFiles.Select(file => new ResourceItem(
                            file.Category,
                            file.Name,
                            string.IsNullOrWhiteSpace(file.File)
                                ? string.Empty
                                : Path.Combine(DataManager.DataRoot, file.File.Replace('/', Path.DirectorySeparatorChar))));

                        var webLinks = boardData.BoardLinks.Select(link => new ResourceItem(
                            link.Category,
                            link.Name,
                            link.Url));

                        this.TabResources.LoadData(localFiles, webLinks);
                        this.TabOverview.LoadData(boardData);
                        this.TabContribute.LoadData(boardData, this._localRegion);
                        this.TabOverview.ApplyFilter(this.ComponentSearchTextBox?.Text ?? string.Empty);
                    }, DispatcherPriority.Background);

                    if (rawPaths.Count > 0)
                    {
                        await Dispatcher.UIThread.InvokeAsync(async () =>
                        {
                            if (loadVersion != this._boardSelectionLoadVersion)
                            {
                                return;
                            }

                            await this.TabSchematicsControl.LoadKiCadProjectForCurrentBoardAsync();
                        }, DispatcherPriority.Background);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Background board load failed - [{ex}]");
                }
            });
        }

        // ###########################################################################################
        // Handles component selection changes and drives highlight updates in both the main viewer
        // and all thumbnails.
        // ###########################################################################################
        private void OnComponentFilterSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressComponentHighlightUpdate)
                return;

            var boardLabels = this.ComponentFilterListBox.SelectedItems?
                .Cast<ComponentListItem>()
                .Select(item => item.BoardLabel)
                .Where(l => !string.IsNullOrEmpty(l))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            this.TabSchematicsControl.UpdateHighlightsForComponents(boardLabels);
            this.TabOverview.ApplyFilter(this.ComponentSearchTextBox?.Text ?? string.Empty);
        }

        // ###########################################################################################
        // Saves the selected category list for the current board whenever the user changes it.
        // ###########################################################################################
        private void OnCategoryFilterSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressCategoryFilterSave)
                return;

            var boardKey = this.GetCurrentBoardKey();
            if (string.IsNullOrEmpty(boardKey))
                return;

            var selected = this.CategoryFilterListBox.SelectedItems?
                .Cast<string>()
                .ToList() ?? new List<string>();

            UserSettings.SetSelectedCategories(boardKey, selected);

            if (this._currentBoardData != null)
            {
                var previouslySelectedKeys = new HashSet<string>(
                    this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                        .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                    StringComparer.OrdinalIgnoreCase);

                var categoryFilter = new HashSet<string>(selected, StringComparer.OrdinalIgnoreCase);
                var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
                var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData, this._localRegion, categoryFilter, searchTerm);

                this._suppressComponentHighlightUpdate = true;
                this.ComponentFilterListBox.ItemsSource = componentItems;

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                    // selection model has not yet caught up with the new item count.
                    try { this.ComponentFilterListBox.SelectAll(); } catch { }
                }
                else
                {
                    for (int i = 0; i < componentItems.Count; i++)
                    {
                        if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                            this.ComponentFilterListBox.Selection.Select(i);
                    }
                }

                this._suppressComponentHighlightUpdate = false;

                var survivingLabels = componentItems
                    .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    survivingLabels = componentItems
                        .Select(item => item.BoardLabel)
                        .Where(l => !string.IsNullOrEmpty(l))
                        .ToList();
                }

                this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
                this.TabOverview.ApplyFilter(searchTerm);
            }
        }

        // ###########################################################################################
        // Returns a composite key uniquely identifying the current hardware and board selection.
        // ###########################################################################################
        internal string GetCurrentBoardKey()
        {
            if (this.BoardKeyOverrideForTests != null)
                return this.BoardKeyOverrideForTests;

            var hw = this.HardwareComboBox.SelectedItem as string;
            var board = this.BoardComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(hw) || string.IsNullOrEmpty(board))
            {
                return string.Empty;
            }
            return $"{hw}|{board}";
        }

        // ###########################################################################################
        // Returns the currently selected hardware/board entry, or null if the selection is invalid.
        // ###########################################################################################
        internal HardwareBoardEntry? GetCurrentBoardEntry()
        {
            var selectedHardware = this.HardwareComboBox.SelectedItem as string;
            var selectedBoard = this.BoardComboBox.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(selectedHardware) || string.IsNullOrWhiteSpace(selectedBoard))
            {
                return null;
            }

            return DataManager.HardwareBoards.FirstOrDefault(entry =>
                string.Equals(entry.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(entry.BoardName, selectedBoard, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Finds the hardware/board entry a "{hardware}|{board}" composite key names, by matching
        // against DataManager.HardwareBoards rather than splitting the string - a hardware or board
        // name containing "|" would otherwise split wrong, and this way GetCurrentBoardKey stays the
        // one place that knows the composite's format.
        //
        // THE one reader of that format besides GetCurrentBoardKey itself: both
        // TryResolveHardwareAndBoardForBoardKey and FormatBoardKeyForDisplay go through here rather
        // than each running their own FirstOrDefault, so a change to the key's shape is a change in
        // exactly two places rather than three.
        // ###########################################################################################
        private static HardwareBoardEntry? FindEntryForBoardKey(string boardKey) =>
            DataManager.HardwareBoards.FirstOrDefault(e =>
                string.Equals($"{e.HardwareName}|{e.BoardName}", boardKey, StringComparison.OrdinalIgnoreCase));

        // ###########################################################################################
        // Reverses a board key into the two drop-down names, for the worklog picker jumping to a
        // workbook's board when it differs from the one on screen.
        // ###########################################################################################
        private static bool TryResolveHardwareAndBoardForBoardKey(string boardKey, out string hardwareName, out string boardName)
        {
            var entry = FindEntryForBoardKey(boardKey);

            hardwareName = entry?.HardwareName ?? string.Empty;
            boardName = entry?.BoardName ?? string.Empty;
            return entry != null;
        }

        // ###########################################################################################
        // Short "hardware/board" label for a workbook's BoardKey, for the worklog picker's item
        // labels - see WorklogJobBox.ItemTemplate in the constructor. Deliberately not
        // TryResolveHardwareAndBoardForBoardKey's HardwareName/BoardName (the full names from the
        // main Excel sheet, e.g. "Commodore 64" / "250469 (short board)") - those are too long once
        // several workbooks from different boards sit in one dropdown. HardwareBoardEntry's own
        // ShortHardwareBoardLabel reads the same short names straight from ExcelDataFile's folder
        // structure instead (e.g. "C64/250469") - see its own comment for why.
        //
        // Falls back to the raw key if the board no longer exists in the synced data, e.g. content
        // that was later removed from classic-repair-toolbox.dk.
        // ###########################################################################################
        internal static string FormatBoardKeyForDisplay(string boardKey) =>
            FindEntryForBoardKey(boardKey)?.ShortHardwareBoardLabel is { Length: > 0 } shortLabel
                ? shortLabel
                : boardKey;

        // ###########################################################################################
        // The FULL hardware and board names for a workbook's BoardKey, as the two drop-downs at the
        // top-left show them - "Commodore 128" and "310378 (C128 & C128D)".
        //
        // For the Workbooks tab's cards, which have a line of their own to spend on this and are
        // read while deciding which repair job to open. They used to show
        // FormatBoardKeyForDisplay's short folder-derived label ("C128/310378"), which is a
        // storage-layout detail rather than a name the user chose the board by - the drop-downs
        // never show it, so it had to be mentally mapped back. The short form is still right for
        // the worklog PICKER, where several boards sit in one dropdown row and the full names run
        // too long, so both formatters exist deliberately.
        //
        // The names come from the KEY ITSELF ("HardwareName|BoardName") rather than from the synced
        // data, so a workbook whose board was later removed from classic-repair-toolbox.dk still
        // names its board properly instead of falling back to a raw key with a pipe in it. The
        // resolved entry is preferred when there is one, since that is the authority on the current
        // names; the split is the fallback.
        //
        // Returns the two parts rather than one joined string: the card renders them on separate
        // lines, and joining here would make the caller split them again.
        // ###########################################################################################
        internal static (string Hardware, string Board) FormatBoardKeyAsFullNames(string boardKey)
        {
            var entry = FindEntryForBoardKey(boardKey);

            if (entry != null &&
                !string.IsNullOrWhiteSpace(entry.HardwareName) &&
                !string.IsNullOrWhiteSpace(entry.BoardName))
            {
                return (entry.HardwareName, entry.BoardName);
            }

            // "Hardware|Board" is the key's own shape - see FindEntryForBoardKey, which builds it.
            int separator = boardKey?.IndexOf('|') ?? -1;

            if (separator > 0 && separator < boardKey!.Length - 1)
            {
                return (boardKey[..separator], boardKey[(separator + 1)..]);
            }

            // Not a key of the expected shape at all (blank, or hand-edited): show whatever it is on
            // one line rather than an empty card row.
            return (boardKey ?? string.Empty, string.Empty);
        }

        // ###########################################################################################
        // Resolves the full path to the currently selected board Excel file.
        // ###########################################################################################
        internal string GetCurrentBoardExcelPath()
        {
            var entry = this.GetCurrentBoardEntry();
            if (entry == null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                return string.Empty;
            }

            return Path.Combine(DataManager.DataRoot, entry.ExcelDataFile.Replace('/', Path.DirectorySeparatorChar));
        }

        // ###########################################################################################
        // Resolves full paths to modern raw KiCad files for the currently selected board.
        // Raw files are auto-discovered from the board-local KiCad folder.
        // ###########################################################################################
        internal List<string> GetCurrentBoardKiCadRawPaths()
        {
            var entry = this.GetCurrentBoardEntry();
            if (entry == null)
            {
                return new List<string>();
            }

            var paths = new List<string>();

            string boardExcelPath = this.GetCurrentBoardExcelPath();
            string boardDirectory = Path.GetDirectoryName(boardExcelPath) ?? string.Empty;
            string kiCadDirectory = Path.Combine(boardDirectory, "KiCad data");

            if (Directory.Exists(kiCadDirectory))
            {
                foreach (string path in Directory.EnumerateFiles(kiCadDirectory, "*.*", SearchOption.TopDirectoryOnly)
                             .Where(ComponentListBuilder.IsSupportedKiCadRawFile))
                {
                    paths.Add(path);
                }
            }

            var result = paths
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return result;
        }

        // ###########################################################################################
        // Reloads the currently selected board from disk and restores the given schematic selection.
        // ###########################################################################################
        internal void ReloadCurrentBoardFromDisk(string schematicNameToRestore)
        {
            var boardKey = this.GetCurrentBoardKey();
            if (!string.IsNullOrWhiteSpace(boardKey) && !string.IsNullOrWhiteSpace(schematicNameToRestore))
            {
                UserSettings.SetLastSchematicForBoard(boardKey, schematicNameToRestore);
            }

            var entry = this.GetCurrentBoardEntry();
            if (entry != null && !string.IsNullOrWhiteSpace(entry.ExcelDataFile))
            {
                BoardDataReader.ClearCache(entry.ExcelDataFile);
            }

            // Not a board CHANGE, so the component search box is deliberately left alone - the same
            // thing passing a null sender used to express before this became a typed argument.
            _ = this.LoadSelectedBoardAsync(clearComponentSearch: false);
        }

        // ###########################################################################################
        // Saves the left panel width after the main splitter drag ends.
        // ###########################################################################################
        private void OnMainSplitterPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            Dispatcher.UIThread.Post(() => UserSettings.LeftPanelWidth = this.LeftPanel.Bounds.Width);
        }

        // ###########################################################################################
        // On first open: validates the saved position is on a live screen and focuses the search.
        // ###########################################################################################
        private void OnWindowFirstOpened(object? sender, EventArgs e)
        {
            this.Opened -= this.OnWindowFirstOpened;

            // Deferred from the constructor - see its own comment. The constructor already collapsed
            // the inline column if the GLOBAL setting was on; ApplyThumbnailsDetachedState below
            // re-derives the real answer (global, or per-board once RememberThumbnailWindowSettingsPerBoard
            // is on and PopulateHardwareDropDown has since selected a real board - see
            // ApplyThumbnailsDetachedStateForBoardChange's own skip-before-first-open guard, which is
            // what left this as the one place the STARTUP board's state gets applied) and both opens
            // the window and brings the column in or out of line with it, so it is safe to call
            // unconditionally rather than only for the "on" case.
            //
            // No fullscreen guard: the two modes COMPOSE (see ApplyThumbnailsDetachedState). The
            // gallery binds to the tab's own hosted ListBox, which travels WITH the tab into the
            // fullscreen window and keeps working from there, so a detached window opened during
            // fullscreen is fully usable rather than wired to something unreachable.
            this.ApplyThumbnailsDetachedState();

            if (UserSettings.HasWindowPlacement && this.WindowState == Avalonia.Controls.WindowState.Normal)
            {
                if (UserSettings.WindowState == nameof(Avalonia.Controls.WindowState.Maximized))
                {
                    // Some Linux window managers (X11/Wayland) ignore a Maximized WindowState set
                    // before the window is shown/mapped, since the WM only honors the maximize
                    // request once it can see the window. Re-assert it now that the window is open.
                    this.WindowState = Avalonia.Controls.WindowState.Maximized;
                }
                else
                {
                    double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
                    int centerX = this._restorePosition.X + (int)((this._restoreWidth * scaling) / 2);
                    int centerY = this._restorePosition.Y + (int)((this._restoreHeight * scaling) / 2);

                    bool isOnScreen = this.Screens.All.Any(s =>
                        centerX >= s.Bounds.X &&
                        centerY >= s.Bounds.Y &&
                        centerX < s.Bounds.X + s.Bounds.Width &&
                        centerY < s.Bounds.Y + s.Bounds.Height);

                    if (!isOnScreen)
                    {
                        var primary = this.Screens.Primary;
                        if (primary != null)
                        {
                            this.Position = new PixelPoint(
                                primary.Bounds.X + Math.Max(0, (primary.Bounds.Width - (int)(this.Width * scaling)) / 2),
                                primary.Bounds.Y + Math.Max(0, (primary.Bounds.Height - (int)(this.Height * scaling)) / 2));
                        }
                    }
                }
            }

            this.PropertyChanged += (s, args) =>
            {
                if (!this._windowPlacementReady)
                    return;

                if (args.Property == Window.WindowStateProperty)
                    this.ScheduleWindowPlacementSave();
            };

            this.PositionChanged += this.OnWindowPositionChanged;
            this.SizeChanged += this.OnWindowSizeChanged;

            if (UserSettings.ValidateDataOnLaunch)
            {
                this.thisBackgroundDataValidationTask = StartBackgroundDataValidationAsync();
            }

            this.thisWindowOpenedCompletionSource.TrySetResult(true);
            this.ScheduleOrphanAndUnusedFileCleanupIfEnabled();

            Dispatcher.UIThread.Post(() => this._windowPlacementReady = true, DispatcherPriority.Background);

            Dispatcher.UIThread.Post(() =>
            {
                this.TabOscilloscopeControl.InitializeForMainWindow(this);
            }, DispatcherPriority.Background);

            Dispatcher.UIThread.Post(() =>
            {
                this.ComponentSearchTextBox?.Focus();
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Shows or hides the "Oscilloscope" tab to match the "Enable network connected oscilloscope
        // tab" configuration setting. If the tab is hidden while it is the selected one, selection
        // falls back to the first still-visible tab so the tab control never shows an empty page.
        //
        // The same setting also governs auto-connect and the oscilloscope rows in the component info
        // popups, so the tab is told to stop or resume its background work and any popup that is
        // already open is refreshed here, rather than only when it is next opened.
        // ###########################################################################################
        public void ApplyOscilloscopeTabVisibility()
        {
            if (this.OscilloscopeTabItem == null || this.MainTabControl == null)
                return;

            bool isEnabled = UserSettings.EnableNetworkConnectedOscilloscopeTab;
            this.OscilloscopeTabItem.IsVisible = isEnabled;

            this.TabOscilloscopeControl.ApplyOscilloscopeTabAvailability();

            this.UpdateComponentInfoWindowsOscilloscopeSessionState(
                this.TabOscilloscopeControl.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                this.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

            if (isEnabled)
                return;

            this.MoveSelectionOffHiddenTab(this.OscilloscopeTabItem);
        }

        // ###########################################################################################
        // Moves tab selection to the first still-visible tab, but only if the tab just hidden is the
        // one currently selected - otherwise the tab control is left showing an empty page.
        //
        // One helper rather than a copy in each Apply*Visibility: there are two conditional tabs now
        // (Oscilloscope and Workbooks) and the block was verbatim in both. Each caller keeps only its
        // own feature-specific teardown.
        // ###########################################################################################
        private void MoveSelectionOffHiddenTab(TabItem? hiddenTab)
        {
            if (this.MainTabControl == null || !ReferenceEquals(this.MainTabControl.SelectedItem, hiddenTab))
                return;

            var firstVisibleTab = this.MainTabControl.Items
                .OfType<TabItem>()
                .FirstOrDefault(tab => tab.IsVisible);

            if (firstVisibleTab != null)
                this.MainTabControl.SelectedItem = firstVisibleTab;
        }

        // ###########################################################################################
        // Re-filters the Hardware/Board drop-downs against the Configuration tab's catalogue
        // visibility tree after a checkbox there changes, so an unchecked hardware/board disappears
        // immediately rather than only after the next board load.
        //
        // Reuses RefreshHardwareAndBoardSelectionsAfterMainExcelSync, which already preserves the
        // current selection where it is still valid and falls back to index 0 otherwise - exactly
        // what is needed here too. If the currently selected board's own visibility changed (now
        // hidden), the selection-index change that follows re-fires OnHardwareSelectionChanged/
        // OnBoardSelectionChanged through the normal SelectionChanged event, which rebuilds the
        // schematic thumbnails using the same filter.
        //
        // WHICH of those two happens is decided here, rather than by always doing both. That
        // refresh reassigns HardwareComboBox.ItemsSource, and reassigning it re-fires the whole
        // selection chain even when the list is identical - which means a full board reload: the
        // board Excel re-read, every schematic bitmap re-decoded. Running that on EVERY checkbox
        // toggle made ticking through a tree of schematics, none of which belong to the board on
        // screen, cost a board reload each time.
        //
        // So: a hardware/board key rebuilds the drop-downs, but only when their contents actually
        // changed (unchecking an already-hidden board's schematic changes nothing). A SCHEMATIC key
        // never changes a drop-down at all - it only changes which thumbnails the current board
        // shows, and then only when the key belongs to the board on screen, which is the one case
        // that still needs a reload.
        // ###########################################################################################
        public void ApplyCatalogueVisibility(string changedKey)
        {
            var uncheckedKeys = UserSettings.CatalogueUncheckedKeysSnapshot;

            if (CatalogueVisibility.IsSchematicKey(changedKey))
            {
                // A schematic toggle only matters to the board currently on screen; the filter is
                // applied while a board loads, so any other board picks it up on its next load.
                if (CatalogueVisibility.KeyNamesBoard(changedKey, this.GetCurrentBoardKeyParts()))
                {
                    this.ReloadCurrentBoardForCatalogueVisibility();
                }

                return;
            }

            var previousHardwareNames = this.HardwareComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            var previousBoardNames = this.BoardComboBox.ItemsSource?
                .Cast<string>()
                .ToList() ?? new List<string>();

            var hardwareNames = BuildVisibleHardwareNames(uncheckedKeys);
            var boardNames = this.BuildVisibleBoardNamesForSelectedHardware(uncheckedKeys);

            if (hardwareNames.SequenceEqual(previousHardwareNames, StringComparer.OrdinalIgnoreCase) &&
                boardNames.SequenceEqual(previousBoardNames, StringComparer.OrdinalIgnoreCase))
            {
                return;
            }

            this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();
        }

        // ###########################################################################################
        // Re-runs the current board's load so the schematic-visibility filter inside it is applied
        // afresh. Re-selecting the same index raises no SelectionChanged, so the load is invoked
        // directly. The component search box is left alone: the board has not changed, and the user
        // was ticking a checkbox on another tab, not choosing a different board.
        // ###########################################################################################
        private void ReloadCurrentBoardForCatalogueVisibility()
        {
            if (this.BoardComboBox.SelectedItem is not string)
            {
                return;
            }

            _ = this.LoadSelectedBoardAsync(clearComponentSearch: false);
        }

        // ###########################################################################################
        // The selected hardware and board as the pair CatalogueVisibility compares a key against -
        // empty strings when nothing is selected, which no key can match.
        // ###########################################################################################
        private (string HardwareName, string BoardName) GetCurrentBoardKeyParts() =>
            (this.HardwareComboBox.SelectedItem as string ?? string.Empty,
             this.BoardComboBox.SelectedItem as string ?? string.Empty);

        private static List<string> BuildVisibleHardwareNames(IReadOnlySet<string> uncheckedKeys) =>
            DataManager.HardwareBoards
                .Select(e => e.HardwareName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(hardwareName => HasAnyVisibleBoard(uncheckedKeys, hardwareName))
                .ToList();

        private List<string> BuildVisibleBoardNamesForSelectedHardware(IReadOnlySet<string> uncheckedKeys)
        {
            if (this.HardwareComboBox.SelectedItem is not string selectedHardware)
            {
                return new List<string>();
            }

            return DataManager.HardwareBoards
                .Where(entry => string.Equals(entry.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.BoardName)
                .Where(b => !string.IsNullOrWhiteSpace(b))
                .Where(b => CatalogueVisibility.IsBoardVisible(uncheckedKeys, selectedHardware, b))
                .ToList();
        }

        // ###########################################################################################
        // Shows or hides the permanent worklog bar above the tabs AND the "Workbooks" tab to match
        // the "Enable Worklog" configuration setting - and, when switching the feature off, tears
        // down everything the feature had put on the schematic.
        //
        // Hiding the bar alone was not enough: the entry overlays, "#N" badges and thumbnail pills
        // all stayed drawn and clickable, and any active entry-drawing mode stayed live with its
        // cross cursor - while the only controls that could dismiss them had just been hidden.
        //
        // The Workbooks tab is part of the same feature and follows the same switch, so it is
        // driven from here rather than from a second method that could fall out of step. As with
        // the oscilloscope tab, hiding it while it is the SELECTED tab moves selection to the first
        // still-visible tab, otherwise the tab control would be left showing an empty page.
        // ###########################################################################################
        public void ApplyWorklogBarVisibility()
        {
            if (this.WorklogBar == null)
                return;

            bool isEnabled = UserSettings.EnableWorklog;
            this.WorklogBar.IsVisible = isEnabled;

            if (this.WorkbooksTabItem != null)
                this.WorkbooksTabItem.IsVisible = isEnabled;

            if (isEnabled)
            {
                // Rebuilt BEFORE the early return, matching ApplyOscilloscopeTabVisibility's own
                // enable-side work: the tab was hidden while board changes and worklog edits went on
                // behind it (RefreshWorklogBar rebuilds it, but this method is what makes it visible
                // again), so re-ticking "Enable Worklog" would otherwise reveal whatever was last
                // rendered - which board, and which workbook, is anyone's guess.
                this.RefreshWorklogBar();
                return;
            }

            this.TabSchematicsControl.CancelWorklogEntryMode();
            this.TabSchematicsControl.SetShowWorklogEntriesList(false, 0);
            this._worklogShowEntriesWorkbookId = 0;

            this.MoveSelectionOffHiddenTab(this.WorkbooksTabItem);
        }

        // ###########################################################################################
        // Shows/hides the Schematics tab's inline thumbnail column and opens/closes the detached
        // thumbnails window to match the "Detach thumbnails into their own window" setting.
        //
        // Composes with F11 schematics fullscreen rather than excluding it: fullscreen collapses the
        // same column this mode does, so the two agree about the layout and only have to agree about
        // who restores the strip. EnterThumbnailsDetachedMode sets its flag and leaves the columns
        // alone while fullscreen is active, and ExitFullscreenMode checks that flag before bringing
        // the strip back - so the detached window stays usable beside a fullscreen schematic, and
        // leaving fullscreen lands on whichever state the setting actually describes.
        // ###########################################################################################
        // ###########################################################################################
        // The ONE way to turn "Detach thumbnails into their own window" on or off: persist the setting,
        // apply the layout, and resync the Configuration checkbox to match. Every surface that
        // offers the toggle goes through here - the checkbox itself, the thumbnail panel's
        // right-click menu, and the detached window's own OS close button - so a step added to the
        // feature is added once rather than found and updated at three call sites.
        //
        // Resyncing the checkbox is harmless when the checkbox IS the caller: it re-reads the
        // setting it just wrote, behind the suppress flag that stops it re-entering its own handler.
        // ###########################################################################################
        internal void SetThumbnailsDetached(bool isDetached)
        {
            string boardKey = this.GetCurrentBoardKey();

            // Per board rather than the single global flag: this board's own detach state is
            // what "Detach thumbnails into their own window" now means when it is toggled, matching
            // the checkbox that gates the setting rather than the global default it falls back
            // to for a board that has never had its own choice made yet (see
            // ResolveThumbnailsDetachedForCurrentBoard).
            //
            // With NO board selected there is no per-board record to write: GetCurrentBoardKey
            // documents an empty string for that state, and storing under it would give every
            // no-selection moment one shared "" entry that the next real board could inherit a
            // detach state from. The global flag is the honest place for the choice then, and it is
            // also what ResolveThumbnailsDetachedForCurrentBoard falls back to for the same key.
            if (UserSettings.RememberThumbnailWindowSettingsPerBoard && !string.IsNullOrEmpty(boardKey))
            {
                UserSettings.SetThumbnailWindowDetachedForBoard(boardKey, isDetached);
            }
            else
            {
                UserSettings.DetachSchematicsThumbnails = isDetached;
            }

            this.ApplyThumbnailsDetachedState();
            this.RefreshDetachThumbnailsCheckBox();
        }

        // ###########################################################################################
        // Resyncs the Configuration tab's "Detach thumbnails into their own window" checkbox with the
        // state actually in force for the current board. Hands the resolved value over rather than
        // letting the tab read a setting for itself, so the checkbox cannot end up showing a
        // different tier from the one the layout is following - see
        // ResolveThumbnailsDetachedForCurrentBoard.
        // ###########################################################################################
        internal void RefreshDetachThumbnailsCheckBox()
            => this.TabConfiguration.RefreshDetachSchematicsThumbnailsCheckBoxFromSettings(
                this.ResolveThumbnailsDetachedForCurrentBoard());

        // ###########################################################################################
        // Whether thumbnails should be shown detached for the CURRENT board: the per-board saved
        // choice while "Remember thumbnail window settings per board" is on (embedded, as asked
        // for, when this board has never had one saved), or the single global
        // DetachSchematicsThumbnails flag otherwise. The one place this decision is made, so the
        // checkbox, the board-switch restore and ApplyThumbnailsDetachedState cannot disagree.
        //
        // internal rather than private because the Configuration checkbox has to resync itself
        // against the SAME tier SetThumbnailsDetached writes to. Reading the global flag there while
        // a per-board one was written made the checkbox re-tick itself to the wrong state on every
        // toggle in per-board mode - see RefreshDetachSchematicsThumbnailsCheckBoxFromSettings.
        //
        // An empty board key (no selection - see GetCurrentBoardKey) falls back to the global flag
        // rather than looking up a shared "" record, matching what SetThumbnailsDetached writes in
        // that same state.
        // ###########################################################################################
        internal bool ResolveThumbnailsDetachedForCurrentBoard()
        {
            if (!UserSettings.RememberThumbnailWindowSettingsPerBoard)
                return UserSettings.DetachSchematicsThumbnails;

            string boardKey = this.GetCurrentBoardKey();
            if (string.IsNullOrEmpty(boardKey))
                return UserSettings.DetachSchematicsThumbnails;

            return UserSettings.GetThumbnailWindowSettingsForBoard(boardKey)?.IsDetached == true;
        }

        // ###########################################################################################
        // Re-applies the detached-thumbnails state for the board just switched to, while "Remember
        // thumbnail window settings per board" is on - a no-op otherwise, so board switches behave
        // exactly as before when the setting is off (the single global flag does not vary by board,
        // so there is nothing to re-apply).
        //
        // Always closes and reopens the window on a board change rather than trying to move it: two
        // boards can each want their OWN size/position (a maximized window for one, a small one for
        // another, as asked for), and SchematicsThumbnailsWindow resolves its layout once, in
        // Initialize - so the only way to pick up a different board's saved layout is a fresh
        // window. Closing first also means the CLOSING board's own layout is saved (via its Closing
        // handler) before the new board's window is opened, rather than losing whatever size/position
        // the user just left it at.
        //
        // Skipped entirely before Main itself has ever been shown: PopulateHardwareDropDown (in
        // StartAsync) selects the startup board, and so raises OnBoardSelectionChanged, BEFORE
        // App.axaml.cs calls Main.Show() - Window.Show(owner) throws "Cannot show window with
        // non-visible owner" if attempted this early, which is exactly why the global (non-per-board)
        // open is already deferred to OnWindowFirstOpened (see the constructor's own comment).
        // OnWindowFirstOpened's own detach-open call handles the startup board once the window is
        // actually open, so nothing is lost by skipping here.
        // ###########################################################################################
        private void ApplyThumbnailsDetachedStateForBoardChange()
        {
            if (!UserSettings.RememberThumbnailWindowSettingsPerBoard)
            {
                // An OPEN window still has to be dealt with even in global mode: unticking "Remember
                // thumbnail window settings per board" while one is open used to make this return
                // here forever, so the window kept the thisPerBoardKey it was opened with and wrote
                // the newly selected board's geometry into the PREVIOUS board's saved layout. The
                // window's key is resolved once, in its Initialize, so the only way to correct it is
                // a fresh window.
                //
                // With no window open there is genuinely nothing to re-apply: the global flag does
                // not vary by board, so a board change cannot change the answer.
                if (this._schematicsThumbnailsWindow == null)
                    return;
            }

            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.CloseThumbnailsWindowForReopen();
            this.ApplyThumbnailsDetachedState();
        }

        // ###########################################################################################
        // Re-applies the detached-thumbnails state after "Remember thumbnail window settings per
        // board" has itself been toggled - the tier that decides both "is this board detached" and
        // "where does its window layout get saved" has just changed under an already-open window, so
        // the answer for the current board can flip and the open window's own board key is now stale.
        //
        // Goes through the same close-and-reopen the board change does rather than just calling
        // ApplyThumbnailsDetachedState: an open window resolved its layout tier once in Initialize
        // (see SchematicsThumbnailsWindow), so a window opened in one tier would keep saving into the
        // other one's record.
        // ###########################################################################################
        internal void ApplyThumbnailsDetachedStateForPerBoardModeChange()
        {
            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.CloseThumbnailsWindowForReopen();
            this.ApplyThumbnailsDetachedState();
        }

        // ###########################################################################################
        // Closes the detached thumbnails window as BOOKKEEPING - so its Closed handler leaves the
        // detach setting alone instead of reading the close as the user dismissing the window - and
        // lets the caller reopen a fresh one. A no-op when no window is open.
        //
        // Close() raises Closing (saving the closing window's layout under the key it was opened
        // with) then Closed (nulling the field), so the caller's ApplyThumbnailsDetachedState opens a
        // fresh window resolved against the CURRENT board and tier.
        // ###########################################################################################
        private void CloseThumbnailsWindowForReopen()
        {
            var window = this._schematicsThumbnailsWindow;
            if (window == null)
                return;

            this._thumbnailsWindowClosingForReopen = window;
            try
            {
                window.Close();
            }
            finally
            {
                // Only ever clears the marker for THIS window. The Closed closure compares against
                // the instance it was created for, so clearing here cannot race it either way.
                if (ReferenceEquals(this._thumbnailsWindowClosingForReopen, window))
                    this._thumbnailsWindowClosingForReopen = null;
            }
        }

        public void ApplyThumbnailsDetachedState()
        {
            bool isEnabled = this.ResolveThumbnailsDetachedForCurrentBoard();

            if (isEnabled)
            {
                this.TabSchematicsControl.EnterThumbnailsDetachedMode();
                this.OpenSchematicsThumbnailsWindow();
            }
            else
            {
                this.CloseSchematicsThumbnailsWindow();
                this.TabSchematicsControl.ExitThumbnailsDetachedMode();

                // ExitThumbnailsDetachedMode restores column 0/2 to whatever widths were LIVE on the
                // grid when detach mode was entered - which, on a startup with the setting already
                // ticked, is still the XAML-declared default (detach happens in the constructor,
                // before OnBoardSelectionChanged has ever applied the per-board ratio to the live
                // grid). Re-applying the saved ratio here, the same way a board change does, is what
                // actually lands the splitter where the user last dragged it for this board rather
                // than at the default. A no-op when fullscreen owns the columns instead (the call
                // below applies its own skip).
                this.ApplySchematicsSplitterRatioForCurrentBoard(this.GetCurrentBoardKey());
            }

        }

        // ###########################################################################################
        // Applies the per-board Schematics/thumbnail splitter ratio (or the default) to the live
        // inner grid. Shared by OnBoardSelectionChanged (every board change) and
        // ApplyThumbnailsDetachedState (leaving detached-thumbnails mode) so the two cannot drift.
        //
        // Skipped while thumbnails are detached or the schematic is fullscreen: both modes collapse
        // columns 0/2 down to their own layout (see TabSchematics.ThumbnailsDetach.cs and
        // EnterFullscreenMode) - applying it while either is active would silently re-expand column 2
        // right back out from under it, leaving an empty widened column since the inline list itself
        // stays hidden either way.
        // ###########################################################################################
        private void ApplySchematicsSplitterRatioForCurrentBoard(string boardKey)
        {
            if (this.TabSchematicsControl.IsThumbnailsDetached || this.TabSchematicsControl.IsFullscreenModeActive)
                return;

            var innerGrid = this.TabSchematicsControl.FindControl<Grid>("SchematicsInnerGrid");

            if (innerGrid == null)
                return;

            if (UserSettings.HasSchematicsSplitterRatio(boardKey))
            {
                double ratio = UserSettings.GetSchematicsSplitterRatio(boardKey);
                innerGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
                innerGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);
            }
            else
            {
                innerGrid.ColumnDefinitions[0].Width = new GridLength(1.0, GridUnitType.Star);
                innerGrid.ColumnDefinitions[2].Width = new GridLength(300.0, GridUnitType.Pixel);
            }
        }

        // ###########################################################################################
        // Opens the detached thumbnails window, bound to the same currentThumbnails collection and
        // hosted (now-hidden) SchematicsThumbnailList the inline gallery uses. A no-op if already open.
        // ###########################################################################################
        private void OpenSchematicsThumbnailsWindow()
        {
            if (this._schematicsThumbnailsWindow != null)
                return;

            var hostedThumbnailList = this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList");
            if (hostedThumbnailList == null)
                return;

            // null (the app-wide layout tier) both when per-board mode is off AND when no board is
            // selected - GetCurrentBoardKey returns an empty string for the latter, and handing that
            // to the window would save every no-selection layout into one shared "" record that a
            // real board could then restore from. Matches the same guard in SetThumbnailsDetached
            // and ResolveThumbnailsDetachedForCurrentBoard.
            string currentBoardKey = this.GetCurrentBoardKey();
            string? perBoardKey =
                UserSettings.RememberThumbnailWindowSettingsPerBoard && !string.IsNullOrEmpty(currentBoardKey)
                    ? currentBoardKey
                    : null;

            var thumbnailsWindow = new SchematicsThumbnailsWindow();
            this._schematicsThumbnailsWindow = thumbnailsWindow;
            thumbnailsWindow.Initialize(
                this.TabSchematicsControl.currentThumbnails,
                hostedThumbnailList,
                this.TabSchematicsControl,
                perBoardKey);

            thumbnailsWindow.Closed += (_, _) =>
            {
                this._schematicsThumbnailsWindow = null;

                // Not on the way out: this window is OWNED (see ReownSchematicsThumbnailsWindow) by
                // whichever window was the visible schematics surface, so quitting the application
                // closes it too and lands here. Turning the setting off then meant the preference was
                // wiped on every exit and never survived to the next launch.
                if (this._isApplicationShuttingDown)
                    return;

                // A board switch (or a per-board-mode toggle) closing this window to reopen a fresh
                // one is bookkeeping, not the user dismissing it - the closing board's own detached
                // choice must survive for next time it is shown. Compared by INSTANCE, so this does
                // not depend on the marker still being set when Closed happens to fire - see
                // _thumbnailsWindowClosingForReopen.
                if (ReferenceEquals(this._thumbnailsWindowClosingForReopen, thumbnailsWindow))
                    return;

                // Closing the fullscreen window destroys this one along with it: while fullscreen is
                // open this window is OWNED by it (see ReownSchematicsThumbnailsWindow), and closing
                // an owner closes what it owns. That is the fullscreen window going away, not the
                // user dismissing the thumbnails - so the detach setting must survive, exactly as it
                // does on application shutdown above. ReopenThumbnailsWindowAfterFullscreenClosed
                // (in the fullscreen Closed handler) then brings it back owned by this window.
                if (this._isClosingFullscreenWindow)
                    return;

                // Closed via its own OS close button rather than the checkbox - the window is a
                // third way to turn the feature off, so it goes through the same one entry point
                // the other two do. The field above is already null, so the CloseSchematicsThumbnailsWindow
                // inside it finds nothing to close and this cannot recurse.
                //
                // Uses the board this window was opened FOR (perBoardKey), not whatever board is
                // currently selected - a board switch takes the path above instead of reaching here,
                // so by the time this handler runs for a genuine user-initiated close the two already
                // agree; captured explicitly anyway so this cannot silently turn off the WRONG
                // board's setting if that ever changes.
                bool stillDetachedForOpenedBoard = perBoardKey != null
                    ? UserSettings.GetThumbnailWindowSettingsForBoard(perBoardKey)?.IsDetached == true
                    : UserSettings.DetachSchematicsThumbnails;

                if (stillDetachedForOpenedBoard)
                {
                    this.SetThumbnailsDetached(false);
                }
            };

            // Owned by whichever window is CURRENTLY the visible schematics surface - the fullscreen
            // window if one is already open (e.g. ticking "Detach thumbnails" while in fullscreen),
            // this window otherwise. See ReownSchematicsThumbnailsWindow's header for why this must
            // track fullscreen rather than always being Main.
            this._schematicsThumbnailsWindow.Show(
                (Window?)this._schematicsFullscreenWindow ?? this);
        }

        // ###########################################################################################
        // Closes the detached thumbnails window, if open.
        // ###########################################################################################
        private void CloseSchematicsThumbnailsWindow()
        {
            // The field is NOT nulled here: Close() raises Closed synchronously and that handler
            // owns this field's lifetime. Nulling it again afterwards would mask a close that
            // actually refused (a future Closing handler setting e.Cancel), leaving a visible
            // window Main can no longer find or close - and the next toggle opening a second one
            // on top of it.
            this._schematicsThumbnailsWindow?.Close();
        }

        // ###########################################################################################
        // Resolves the ONE workbook the worklog bar shows, "Show worklogs" draws, and "Add worklog"
        // writes new entries into - the single notion of "the active workbook" every worklog-facing
        // control in Main and TabSchematics shares.
        //
        // Defaults to the board's newest workbook (open or closed - see RefreshWorklogBar's header
        // for why status is not a filter here). Selecting a workbook on the Workbooks tab
        // (TabWorkbooks.SelectWorkbook, via ActivateWorkbook below) overrides that default by saving
        // the choice in UserSettings.ActiveWorkbookIdByBoard, so "activating" an older or closed
        // workbook makes IT the one every other worklog surface acts on until something reactivates
        // the newest one again.
        //
        // The saved id is validated against this board's actual workbooks on every call rather than
        // trusted blindly: a workbook can be deleted from disk by hand, or the saved id can be stale
        // after switching boards, and an unvalidated id would otherwise make the bar quietly show
        // nothing (or another board's workbook, if ids ever collided) instead of falling back.
        //
        // The rule itself lives in WorklogManager.ResolveActiveWorkbook, not here: the Workbooks tab
        // needs the same answer for its highlighted card, and two copies of "saved id if valid, else
        // newest" is exactly how the card and the bar came to disagree. This wrapper only fetches
        // the two inputs.
        // ###########################################################################################
        // internal rather than private so the component popup's "attach capture to worklog" flow can
        // ask the same question this window does, rather than re-deriving "which workbook" a third
        // time - WorklogManager.ResolveActiveWorkbook's own header is explicit that it is the ONE
        // place that rule lives.
        internal static WorkbookRecord? ResolveActiveWorkbookForBoard(string boardKey) =>
            WorklogManager.ResolveActiveWorkbook(
                WorklogManager.GetWorkbooksForBoard(boardKey),
                UserSettings.GetActiveWorkbookId(boardKey));

        // ###########################################################################################
        // Activates a workbook for its board: persists it as the board's active workbook (so it
        // survives a tab switch, a board switch and back, and an app restart) and refreshes every
        // worklog surface that depends on "the active workbook" - the bar, and (via
        // TabSchematicsControl.SetShowWorklogEntriesList inside RefreshWorklogBar) the Schematics
        // tab's "Show worklogs" overlay for whichever schematic is on screen there.
        //
        // Called from TabWorkbooks.SelectWorkbook when the user clicks a card in the Workbooks tab -
        // see that method for why the caller ALSO switches to the Schematics tab, which this method
        // deliberately does not do itself (a board/data refresh must be able to call this without
        // stealing the user's current tab).
        // ###########################################################################################
        // ###########################################################################################
        // Sets the component highlight-rect cache and tells everything that reads it.
        //
        // The cache physically lives on TabSchematics (it is built as a side effect of that tab's
        // board load and most of its readers are there), but MAIN is what actually owns it: all four
        // writes are here - a board switch blanking it, the board load populating it, and both
        // region-filter paths rebuilding it - and TabSchematics never assigns it at all.
        //
        // Routing every write through this one method exists for the SECOND reader. The Workbooks
        // tab needs the same cache for a pill's "Mark components in scope" checklist, and the pane's
        // pills go on screen before the board load's fire-and-forget task has populated it: click one
        // in that window and the lookup missed, the checklist was silently skipped, and the two
        // modals were no longer identical - the exact bug this feature was written to fix, back as an
        // intermittent one. A region switch had the same gap from the other end, leaving the pane
        // stale against a cache that had just been rebuilt with a different region's rects.
        //
        // Refreshing the board pane from here rather than from each write site means a fifth write
        // added later cannot forget to.
        // ###########################################################################################
        private void SetComponentHighlightRects(Dictionary<string, Dictionary<string, List<Rect>>> highlightRects)
        {
            this.TabSchematicsControl.highlightRectsBySchematicAndLabel = highlightRects;
            this.TabWorkbooks?.RefreshBoardPreviewsForCurrentSelection();
        }

        public void ActivateWorkbook(string boardKey, int workbookId)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            // An entry-drawing mode started for the PREVIOUSLY active workbook captured that
            // workbook's id (BeginWorklogEntryMode), and nothing about switching tabs cancels it. So
            // without this, activating another workbook here left the cross cursor live on the
            // Schematics tab and the next drawn entry was written into the workbook the user had just
            // navigated away from, while every visible surface named the new one.
            // ApplyWorklogBarVisibility already performs the same teardown when the feature is
            // switched off; this is the other way "which workbook is being written to" can change.
            this.TabSchematicsControl.CancelWorklogEntryMode();

            UserSettings.SetActiveWorkbookId(boardKey, workbookId);
            this.RefreshWorklogBar();
        }

        // ###########################################################################################
        // Lets the worklog bar's own workbook picker activate ANY workbook directly - including one
        // for a board other than the one currently on screen, since the picker now lists every
        // workbook on every board (see RefreshWorklogBar). Guarded by _suppressWorklogJobBoxSave so
        // RefreshWorklogBar re-seeding the box's ItemsSource/SelectedItem does not loop back in here.
        // ###########################################################################################
        private void OnWorklogJobBoxSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressWorklogJobBoxSave)
                return;

            if (this.WorklogJobBox.SelectedItem is not WorkbookRecord selected)
                return;

            this.ActivateWorkbookAcrossBoards(selected.BoardKey, selected.Id);
        }

        // ###########################################################################################
        // Activates a workbook that may belong to a board other than the one currently on screen -
        // shared by the worklog bar's own picker (OnWorklogJobBoxSelectionChanged) and the Workbooks
        // tab's "Show all workbooks" scope (TabWorkbooks.SelectWorkbook), both of which can now list
        // workbooks for every board rather than just the current one.
        //
        // Same-board case: identical to clicking a card when the tab is scoped to the current board -
        // just calls ActivateWorkbook.
        //
        // Different-board case: the hardware/board selectors are switched FIRST. Persisting the
        // target workbook as its board's active one before that switch (rather than after, via
        // ActivateWorkbook) matters because switching BoardComboBox.SelectedItem starts the async
        // OnBoardSelectionChanged, and that method calls RefreshWorklogBar itself once the board
        // finishes loading (see its own comment on why that call moved after _currentBoardData is
        // assigned) - so by the time it does, ResolveActiveWorkbookForBoard must already see this
        // choice as the saved one, or the freshly loaded board would show its OWN newest workbook
        // instead of the one just picked. Calling ActivateWorkbook here too, after triggering the
        // switch, would be redundant at best and would run against the board being LEFT rather than
        // the one being entered.
        // ###########################################################################################
        public void ActivateWorkbookAcrossBoards(string boardKey, int workbookId)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            if (string.Equals(boardKey, this.GetCurrentBoardKey(), StringComparison.Ordinal))
            {
                this.ActivateWorkbook(boardKey, workbookId);
                return;
            }

            if (!TryResolveHardwareAndBoardForBoardKey(boardKey, out string hardwareName, out string boardName))
            {
                // The board this workbook belongs to is no longer in the synced data (its Excel entry
                // was removed by a later sync), so there is nothing to switch to. Refresh instead of
                // switching, so every surface goes back to naming the workbook that IS active rather
                // than one that cannot be reached.
                this.RefreshWorklogBar();
                return;
            }

            UserSettings.SetActiveWorkbookId(boardKey, workbookId);

            this.TabSchematicsControl.CancelWorklogEntryMode();

            // Switching HardwareComboBox first (if needed) synchronously repopulates BoardComboBox's
            // ItemsSource via OnHardwareSelectionChanged before BoardComboBox.SelectedItem is set -
            // setting the board first would pick an index into the OLD hardware's board list.
            //
            // _pendingBoardSelectionOverride makes that hardware switch land directly on the board we
            // are actually after. Without it OnHardwareSelectionChanged selects the new hardware's
            // SAVED last board, running a whole board load for a board nobody asked for before the
            // real one is set - see the field's own comment.
            var currentHardware = this.HardwareComboBox.SelectedItem as string;
            if (!string.Equals(currentHardware, hardwareName, StringComparison.OrdinalIgnoreCase))
            {
                this._pendingBoardSelectionOverride = boardName;
                this.HardwareComboBox.SelectedItem = hardwareName;

                // Cleared by OnHardwareSelectionChanged, which also selected boardName for us - so
                // the assignment below is a no-op in the normal case and only does real work if that
                // handler could not honour the override (the board vanished between the lookup above
                // and the switch).
            }

            this.BoardComboBox.SelectedItem = boardName;
        }

        // ###########################################################################################
        // Switches the main tab strip to Schematics, if it is not already there. A public wrapper
        // rather than exposing MainTabControl/SchematicsTabItem themselves to other tabs: those
        // fields are Avalonia's own x:Name-generated ones, and TabWorkbooks (the one other caller
        // that needs this, from SelectWorkbook) has no business reaching into Main's tab control
        // directly - the same reasoning OnWorklogAddEntryClick already followed for itself.
        // ###########################################################################################
        public void SwitchToSchematicsTab()
        {
            // Null-guarded like ApplyOscilloscopeTabVisibility and ApplyWorklogBarVisibility, its two
            // siblings that reach for the same control. This is public and callable from another tab
            // now, so it can no longer assume it only runs at a point where the tab control is up.
            if (this.MainTabControl == null || this.SchematicsTabItem == null)
                return;

            if (!ReferenceEquals(this.MainTabControl.SelectedItem, this.SchematicsTabItem))
                this.MainTabControl.SelectedItem = this.SchematicsTabItem;
        }

        // ###########################################################################################
        // Gives the Workbooks tab's own search box priority focus the moment it becomes the selected
        // tab - see TabWorkbooks.FocusSearchBox for how that coexists with the global "steal focus
        // into ComponentSearchTextBox" handler wired in the constructor (which now excludes this tab
        // by header). Every OTHER tab is left alone: this handler only ever hands focus TO the
        // Workbooks box, never takes it away when leaving - the global handler's own per-click logic
        // already covers every tab that wants ComponentSearchTextBox back, including this one once
        // the user clicks something on it.
        //
        // e.Source MUST be checked against MainTabControl itself. SelectionChanged is a bubbling
        // routed event, and several tabs hold their own ListBox/ComboBox controls (the Schematics
        // tab's thumbnail list among them) - without this guard, selecting a thumbnail or any other
        // nested list would bubble up through MainTabControl and re-fire this, stealing focus back
        // to the search box away from whatever the user had just clicked on the Workbooks tab.
        // ###########################################################################################
        private void OnMainTabControlSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!ReferenceEquals(e.Source, this.MainTabControl))
            {
                return;
            }

            if (this.MainTabControl?.SelectedItem is TabItem { Header: "Workbooks" })
            {
                this.TabWorkbooks?.FocusSearchBox();
            }
        }

        // ###########################################################################################
        // Refreshes the worklog bar's content for the currently selected board: either the empty
        // "no jobs recorded" state, or that board's active workbook (see ResolveActiveWorkbookForBoard).
        //
        // The bar deliberately shows the active workbook whether it is Open or Closed. Resolving a
        // workbook's last outstanding entry closes it automatically, and showing only open workbooks
        // meant a finished workbook disappeared from the UI entirely - indistinguishable from having
        // been deleted. Status is presentation only: it picks the status dot's color and appears in
        // the label, and changes nothing about what the bar's buttons offer. "Add worklog" still works
        // on a closed workbook (adding a still-Open entry reopens it through the normal
        // RecomputeWorkbookStatus rule), as does the "Show worklogs" toggle.
        // ###########################################################################################
        public void RefreshWorklogBar()
        {
            if (this.WorklogNoJobText == null)
                return;

            string boardKey = this.GetCurrentBoardKey();

            // ONE disk scan for both lists. GetAllWorkbooks and GetWorkbooksForBoard each go through
            // ReadAllWorkbooks, which enumerates every workbook folder and does File.Exists +
            // ReadAllText + Deserialize per folder, uncached - so calling both would scan the whole
            // Workbook tree twice on every board change, entry save, workbook create/close/delete and
            // card click. The board-scoped list is a SUBSET of the full one, so it is filtered here
            // rather than re-read; the ordering (descending id, newest first) is GetAllWorkbooks' own
            // and is exactly what GetWorkbooksForBoard would have produced.
            var allWorkbooks = WorklogManager.GetAllWorkbooks();
            var workbooks = allWorkbooks
                .Where(w => string.Equals(w.BoardKey, boardKey, StringComparison.Ordinal))
                .ToList();

            var activeWorkbook = WorklogManager.ResolveActiveWorkbook(workbooks, UserSettings.GetActiveWorkbookId(boardKey));
            bool hasWorkbook = activeWorkbook != null;

            // The Workbooks tab's list is rebuilt from here rather than from its own wiring: this
            // method is already the single place worklog state is refreshed from - board changes,
            // entry saves, workbook creation and closure all reach it - so the tab cannot go stale
            // in a case the bar handles and the tab forgot about.
            //
            // Skipped entirely when the worklog feature is switched off. That rebuild decodes every
            // schematic image with an entry in the active workbook (full-resolution PNGs, hundreds of
            // MB of BGRA on a big board) and re-reads entries.json, and it ran on every board change
            // for a user who has the whole feature disabled and the tab hidden. Safe to skip because
            // ApplyWorklogBarVisibility rebuilds the tab on the enable side before revealing it, so
            // it cannot come back showing what was last rendered.
            //
            // Deliberately NOT also gated on "is the Workbooks tab currently selected". Clicking a
            // card in that tab goes ActivateWorkbook -> RefreshWorklogBar -> RefreshWorkbooks, so an
            // early-out on tab selection would make card clicks silently do nothing.
            if (UserSettings.EnableWorklog)
            {
                // The tab's own list can show every board's workbooks ("Show all workbooks", the
                // Configuration tab's radio group below "Enable Worklog") rather than just this
                // board's - allWorkbooks is already the unfiltered read from above, so passing it
                // through costs nothing extra; board is still used to resolve which workbook is
                // active and to filter the fallback when no explicit choice was made (see
                // TabWorkbooks.RefreshWorkbooks).
                bool showAllBoards = string.Equals(UserSettings.WorkbooksScope, "AllBoards", StringComparison.Ordinal);
                this.TabWorkbooks?.RefreshWorkbooks(showAllBoards ? allWorkbooks : workbooks);
            }

            this.WorklogNoJobText.IsVisible = !hasWorkbook;

            // The picker stays visible whenever ANY workbook exists anywhere, not just when THIS
            // board has one. It lists every workbook on every board and is the only way to reach
            // another board's workbook from the bar, so hiding it on a board with none of its own
            // hid it exactly where it is most useful - a board you have never worked on is precisely
            // where you want to jump to the job you were doing elsewhere. (The visibility rule was
            // written for the old read-only text box, which genuinely had nothing to show.)
            this.WorklogJobBox.IsVisible = allWorkbooks.Count > 0;

            // These three still follow the ACTIVE workbook: there is no status to show, no entries to
            // draw and nothing to add an entry to when this board has no workbook.
            this.WorklogJobStatusPanel.IsVisible = hasWorkbook;
            this.WorklogShowEntriesPanel.IsVisible = hasWorkbook;
            this.WorklogAddEntryButton.IsVisible = hasWorkbook;

            if (!hasWorkbook || activeWorkbook!.Id != this._worklogShowEntriesWorkbookId)
            {
                // A different (or no) workbook is now shown - re-seed the checkbox from the user's
                // saved preference and apply it to this workbook directly, rather than relying on
                // IsCheckedChanged (which will not fire below when the new value matches what the
                // checkbox already showed for the previous workbook).
                //
                // The write is suppressed because it is this code seeding the checkbox, not the
                // user clicking it: showByDefault is forced false whenever the board has no
                // workbook, and letting that reach the handler would persist it as the user's
                // preference. See _suppressWorklogShowEntriesSave.
                bool showByDefault = hasWorkbook && UserSettings.WorklogShowEntriesChecked;

                this._suppressWorklogShowEntriesSave = true;
                try
                {
                    this.WorklogShowEntriesCheckBox.IsChecked = showByDefault;
                }
                finally
                {
                    this._suppressWorklogShowEntriesSave = false;
                }

                int workbookId = showByDefault ? activeWorkbook!.Id : 0;
                this._worklogShowEntriesWorkbookId = workbookId;
                this.TabSchematicsControl.SetShowWorklogEntriesList(showByDefault, workbookId);
            }
            else
            {
                // Same workbook as before, but its ENTRIES may have just changed - an area redrawn,
                // "Show marked area" ticked, an entry added or deleted. The branch above only fires
                // when the shown WORKBOOK changes, so without this the schematic overlay and the
                // thumbnail pills kept drawing the entry as it was before the edit until something
                // else happened to rebuild them. Reported as the marker not updating.
                this.TabSchematicsControl.RefreshWorklogEntriesListForCurrentWorkbook();
            }

            // The picker lists EVERY workbook on every board, not just this board's own workbooks in
            // "workbooks" above - selecting one for a different board is how the bar can jump there
            // (see OnWorklogJobBoxSelectionChanged). "allWorkbooks" is the single disk read taken at
            // the top of this method; "workbooks" is its board-scoped subset.
            //
            // SelectedItem is looked up BY ID inside this same list rather than set to "activeWorkbook"
            // directly: WorkbookRecord has no Equals/GetHashCode override, so ComboBox's SelectedItem
            // only shows as selected when it is REFERENCE-equal to an item actually in ItemsSource -
            // and while activeWorkbook now comes from a filtered view of this very list (so reference
            // equality would in fact hold today), the by-id lookup keeps that from being a silent
            // dependency on how "workbooks" happens to be derived.
            //
            // Seeded whether or not a workbook is active: with no active workbook the box shows no
            // selection while still listing every OTHER board's workbooks, which is what makes it
            // usable as a navigator on a board that has none of its own. Suppressed the same way
            // OnWorklogShowEntriesCheckedChanged's seed is: this is RefreshWorklogBar re-populating
            // the list for the board on screen, not a user picking a workbook, and letting it reach
            // OnWorklogJobBoxSelectionChanged would call ActivateWorkbook -> RefreshWorklogBar again
            // for a selection nobody made.
            this._suppressWorklogJobBoxSave = true;
            try
            {
                this.WorklogJobBox.ItemsSource = allWorkbooks;
                this.WorklogJobBox.SelectedItem = hasWorkbook
                    ? allWorkbooks.FirstOrDefault(w => w.Id == activeWorkbook!.Id)
                    : null;
            }
            finally
            {
                this._suppressWorklogJobBoxSave = false;
            }

            if (activeWorkbook == null)
                return;

            // Border, padlock, label and the padlock's overshoot padding all applied by the ONE
            // shared informational styling - see Handlers/Theme/WorklogInfoPillBuilder.cs. This
            // pill is declared in Main.axaml (long-lived, only its text changes), so it is
            // restyled in place rather than rebuilt. It used to be styled here by hand at 2px,
            // which is exactly the drift that class now prevents.
            //
            // The status WORD lives in the pill, so WorklogJobStatusText below carries only the
            // counts and the start date - otherwise "Open" would read twice, once in each.
            Handlers.Theming.WorklogInfoPillBuilder.ApplyStatePillVisual(
                this.WorklogJobStatusPill,
                this.WorklogJobStatusDot,
                this.WorklogJobStatusPillText,
                activeWorkbook.Status);

            // A CLOSED workbook reports when it finished; an open one reports when it started -
            // the whole rule, word and date together, lives in the shared
            // WorklogEntryScope.FormatWorkbookDateLine, which the Workbooks tab's own cards use
            // too. This bar had the rule and the card had a hardcoded "started" against StartDate,
            // so the same workbook read "ended 2026-September-6" here and "started 2026-September-6"
            // on its card a row below. See that method for why "ended" rather than "closed", and
            // for the no-EndDate fallback.
            string dateLine = WorklogEntryScope.FormatWorkbookDateLine(
                activeWorkbook.Status,
                activeWorkbook.StartDate,
                activeWorkbook.EndDate);

            if (activeWorkbook.WorklogCount == 0)
            {
                this.WorklogJobStatusText.Text = $"No worklogs yet · {dateLine}";
                return;
            }

            string worklogWord = activeWorkbook.WorklogCount == 1 ? "worklog" : "worklogs";
            this.WorklogJobStatusText.Text =
                $"{activeWorkbook.WorklogCount} {worklogWord} · {dateLine}";
        }

        // ###########################################################################################
        // Opens the "Create new workbook" dialog for the currently selected board and ACTIVATES the
        // workbook it creates.
        //
        // ActivateWorkbook, not a bare RefreshWorklogBar: "which workbook is active" used to be
        // "the board's newest", which a just-created workbook always was, so a plain refresh was
        // enough. It is now ResolveActiveWorkbookForBoard, which prefers a previously-saved
        // ActiveWorkbookIdByBoard entry whenever that id still names a real workbook - so after the
        // user has ever clicked a card in the Workbooks tab, a refresh alone would leave the bar,
        // "Show worklogs" and "Add worklog" pointing at the OLD workbook and write the next drawn
        // entry into it rather than into the one just created and named.
        // ###########################################################################################
        private async void OnWorklogCreateWorkbookClick(object? sender, RoutedEventArgs e)
        {
            string boardKey = this.GetCurrentBoardKey();
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            var dialog = new CreateWorkbookWindow();
            dialog.Initialize(boardKey);

            var record = await dialog.ShowDialog<WorkbookRecord?>(this);
            if (record == null)
                return;

            this.ActivateWorkbook(boardKey, record.Id);
        }

        // ###########################################################################################
        // Switches to the Schematics tab and enters worklog entry-drawing mode for the ACTIVE
        // workbook - the one ResolveActiveWorkbookForBoard resolves, which the bar is already
        // showing. Not the Open-only lookup, so this keeps working on a closed workbook - the entry
        // it adds is Open, which reopens the workbook anyway.
        // ###########################################################################################
        private void OnWorklogAddEntryClick(object? sender, RoutedEventArgs e)
        {
            var activeWorkbook = ResolveActiveWorkbookForBoard(this.GetCurrentBoardKey());
            if (activeWorkbook == null)
                return;

            this.SwitchToSchematicsTab();

            if (!this.TabSchematicsControl.BeginWorklogEntryMode(activeWorkbook.Id))
                return;

            this.WorklogAddEntryButton.IsEnabled = false;
            this.WorklogCancelEntryButton.IsVisible = true;
        }

        // ###########################################################################################
        // Cancels the in-progress worklog entry-drawing mode. TabSchematics calls back into
        // ResetWorklogEntryModeButtons() once it has actually torn the mode down, so the buttons
        // stay in sync whether cancellation came from here, Escape, or the entry editor closing.
        // ###########################################################################################
        private void OnWorklogCancelEntryClick(object? sender, RoutedEventArgs e)
        {
            this.TabSchematicsControl.CancelWorklogEntryMode();
        }

        // ###########################################################################################
        // Toggles the "Show worklogs" list view for the workbook the bar is showing, and saves the
        // checked state as the user's default for next time. TabSchematics scopes what it actually
        // draws to the schematic currently on screen - see TabSchematics.Worklog.cs's
        // RefreshWorklogEntriesListOverlay for why. Uses the same latest-workbook lookup the bar
        // does, so a closed workbook's entries stay viewable.
        //
        // Only a real user toggle is saved: RefreshWorklogBar seeds the checkbox programmatically
        // and suppresses the save, because the value it seeds is derived from the board on screen
        // rather than from the user's intent. It applies the overlay itself, so returning early
        // here loses nothing.
        // ###########################################################################################
        private void OnWorklogShowEntriesCheckedChanged(object? sender, RoutedEventArgs e)
        {
            if (this._suppressWorklogShowEntriesSave)
                return;

            bool isChecked = this.WorklogShowEntriesCheckBox.IsChecked == true;
            UserSettings.WorklogShowEntriesChecked = isChecked;

            var activeWorkbook = ResolveActiveWorkbookForBoard(this.GetCurrentBoardKey());
            int workbookId = isChecked && activeWorkbook != null ? activeWorkbook.Id : 0;

            this._worklogShowEntriesWorkbookId = workbookId;
            this.TabSchematicsControl.SetShowWorklogEntriesList(isChecked && activeWorkbook != null, workbookId);
        }

        // ###########################################################################################
        // Makes the "Show worklogs" text act like part of the checkbox it labels, since a bare
        // TextBlock does not react to clicks on its own. Toggling the checkbox itself fires
        // IsCheckedChanged, which does the actual work and persists the preference.
        // ###########################################################################################
        private void OnWorklogShowEntriesLabelPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.WorklogShowEntriesCheckBox.IsChecked = !(this.WorklogShowEntriesCheckBox.IsChecked == true);
        }

        // ###########################################################################################
        // Restores the worklog bar's "Add worklog" / "Cancel" buttons to their idle state.
        // Called by TabSchematics whenever worklog entry-drawing mode ends, regardless of trigger.
        // ###########################################################################################
        public void ResetWorklogEntryModeButtons()
        {
            this.WorklogAddEntryButton.IsEnabled = true;
            this.WorklogCancelEntryButton.IsVisible = false;
        }

        // ###########################################################################################
        // Tracks the window's position in Normal state and schedules a debounced save.
        // ###########################################################################################
        private void OnWindowPositionChanged(object? sender, PixelPointEventArgs e)
        {
            if (!this._windowPlacementReady)
                return;

            if (this.WindowState == Avalonia.Controls.WindowState.Normal)
            {
                this._restorePosition = e.Point;
                this.ScheduleWindowPlacementSave();
            }
        }

        // ###########################################################################################
        // Tracks the window's size in Normal state and schedules a debounced save.
        // ###########################################################################################
        private void OnWindowSizeChanged(object? sender, SizeChangedEventArgs e)
        {
            if (!this._windowPlacementReady)
                return;

            if (this.WindowState == Avalonia.Controls.WindowState.Normal)
            {
                this._restoreWidth = e.NewSize.Width;
                this._restoreHeight = e.NewSize.Height;
                this.ScheduleWindowPlacementSave();
            }
        }

        // ###########################################################################################
        // Resets and starts a 500 ms debounce timer;
        // ###########################################################################################
        private void ScheduleWindowPlacementSave()
        {
            if (this._windowPlacementSaveTimer == null)
            {
                this._windowPlacementSaveTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                this._windowPlacementSaveTimer.Tick += (s, e) =>
                {
                    this._windowPlacementSaveTimer.Stop();
                    this.CommitWindowPlacement();
                };
            }

            this._windowPlacementSaveTimer.Stop();
            this._windowPlacementSaveTimer.Start();
        }

        // ###########################################################################################
        // Captures the current window state and screen, then persists to settings.
        // ###########################################################################################
        private void CommitWindowPlacement()
        {
            var state = this.WindowState == Avalonia.Controls.WindowState.Minimized
                ? Avalonia.Controls.WindowState.Normal
                : this.WindowState;

            double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
            double w = this.Bounds.Width > 0 ? this.Bounds.Width : this._restoreWidth;
            double h = this.Bounds.Height > 0 ? this.Bounds.Height : this._restoreHeight;

            int centerX = this.Position.X + (int)((w * scaling) / 2);
            int centerY = this.Position.Y + (int)((h * scaling) / 2);

            var screen = this.Screens.All.FirstOrDefault(s =>
                centerX >= s.Bounds.X &&
                centerY >= s.Bounds.Y &&
                centerX < s.Bounds.X + s.Bounds.Width &&
                centerY < s.Bounds.Y + s.Bounds.Height)
                ?? this.Screens.Primary;

            UserSettings.SaveWindowPlacement(
                state.ToString(),
                this._restoreWidth,
                this._restoreHeight,
                this._restorePosition.X,
                this._restorePosition.Y,
                screen?.Bounds.X ?? 0,
                screen?.Bounds.Y ?? 0,
                screen?.Bounds.Width ?? 1920,
                screen?.Bounds.Height ?? 1080,
                screen?.Scaling ?? 1.0);
        }

        // ###########################################################################################
        // Stops any pending debounce timer and does a final synchronous save on close.
        // ###########################################################################################
        private void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            // Set before anything closes, but only once the close is known to be going ahead: the
            // detached thumbnails window is owned by this one, so Avalonia closes it as part of
            // this shutdown and its Closed handler would otherwise read that as the user dismissing
            // it and turn the setting off - losing the preference on every single exit.
            //
            // The Cancel check matters because this flag is never cleared again. Nothing cancels a
            // close today, but a later confirm-on-exit prompt would leave the flag stuck true for
            // the rest of the session, and the detached window's own OS close button would then
            // silently skip unticking the setting and re-embedding the strip - thumbnails gone from
            // both the tab and the window, with the checkbox still ticked.
            if (e.Cancel)
                return;

            this._isApplicationShuttingDown = true;

            if (this._schematicsFullscreenWindow != null)
            {
                this._schematicsFullscreenWindow.Close();
            }

            this._blinkSelectedTimer?.Stop();
            this._windowPlacementSaveTimer?.Stop();
            this.CommitWindowPlacement();
        }

        // ###########################################################################################
        // Forces the entire application (and all its sub-windows) to shut down once the main window
        // has successfully completed its closing sequence.
        // ###########################################################################################
        private void OnWindowClosed(object? sender, EventArgs e)
        {
            UserSettings.CheckDataOnLaunchChanged -= this.OnCheckDataOnLaunchSettingChanged;
            UserSettings.WorkbooksScopeChanged -= this.OnWorkbooksScopeSettingChanged;
            UserSettings.WorklogCurrencyChanged -= this.OnWorklogCurrencySettingChanged;

            if (Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            {
                Dispatcher.UIThread.Post(() =>
                {
                    desktop.Shutdown();
                });
            }
        }

        // ###########################################################################################
        // Opens the persistent AppData folder that contains the log and settings files.
        // ###########################################################################################
        private void OnOpenAppDataFolderClick(object? sender, RoutedEventArgs e)
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = Path.Combine(appData, AppConfig.AppFolderName);

            try
            {
                Directory.CreateDirectory(directory);

                if (OperatingSystem.IsWindows())
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", $"\"{directory}\"")
                    {
                        UseShellExecute = true
                    });
                }
                else if (OperatingSystem.IsMacOS())
                {
                    Process.Start("open", directory);
                }
                else
                {
                    Process.Start("xdg-open", directory);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open app data folder - [{directory}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Populates About tab fields and loads changelog content from embedded assets.
        // ###########################################################################################
        private void PopulateAboutTab(Assembly assembly, string? versionString)
        {
            this.TabAbout.InitializeAbout(assembly, versionString);
        }

        // ###########################################################################################
        // Opens a validated external target through the shared launcher.
        //
        // Routed through ExternalTargetLauncher rather than calling Process.Start here, so that every
        // outward link in the app passes the same scheme check - ShellExecute runs whatever it is
        // handed, and a single unguarded call is all it takes for that to matter later. The launcher
        // already logs both a refusal and a failed start, so there is nothing to catch at this level.
        // ###########################################################################################
        private static void OpenUrl(string url)
        {
            if (!ExternalTargetLauncher.TryOpen(url))
            {
                Logger.Warning($"Rejected external target from main window: [{url}]");
            }
        }

        // ###########################################################################################
        // Clears the component search text, removes all selected items in the Component filter list,
        // and resets schematic highlights back to an empty selection state.
        // ###########################################################################################
        private void OnClearComponentsClick(object? sender, RoutedEventArgs e)
        {
            this._suppressComponentSearchRefresh = true;
            this.ComponentSearchTextBox.Text = string.Empty;
            this._suppressComponentSearchRefresh = false;

            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.TabSchematicsControl.ApplyLabelEditorSearchFilter(string.Empty);
            }
            else
            {
                this.ApplyNormalComponentSearchFilter(string.Empty);
            }

            this.TabSchematicsControl.ClearSchematicsOnlySelectedComponents();
            this.ComponentFilterListBox.SelectedItems?.Clear();
            this.TabSchematicsControl.UpdateHighlightsForComponents(new List<string>());
        }

        // ###########################################################################################
        // Selects all available items currently populated within the Component filter list box.
        // ###########################################################################################
        private void OnMarkAllComponentsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.ComponentFilterListBox.SelectAll();
            }
            catch (OutOfMemoryException ex)
            {
                Logger.Debug(ex, "Failed to mark all components. The selection was too large to process");
            }
        }

        // ###########################################################################################
        // Switches the local region to PAL and reloads images.
        // ###########################################################################################
        private void OnPalRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this._localRegion = "PAL";
            UserSettings.Region = "PAL";

            this.UpdateRegionButtonsState();
            this.RefreshImages();
            this.TabSchematicsControl.UpdateOverlayLabels();
        }

        // ###########################################################################################
        // Switches the local region to NTSC and reloads images.
        // ###########################################################################################
        private void OnNtscRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;

            this._localRegion = "NTSC";
            UserSettings.Region = "NTSC";

            this.UpdateRegionButtonsState();
            this.RefreshImages();
            this.TabSchematicsControl.UpdateOverlayLabels();
        }

        // ###########################################################################################
        // Updates the region toggle and button states to match the current local region.
        // Hides the entire region toggle area when the current board has no explicit PAL/NTSC components.
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;
            bool isNtsc = string.Equals(this._localRegion, "NTSC", StringComparison.OrdinalIgnoreCase);
            bool hasExplicitRegionComponents = ComponentListBuilder.HasExplicitRegionComponents(this._currentBoardData);

            this.RegionButtonsGrid.IsVisible = hasExplicitRegionComponents;

            this.NtscRegionButton.Classes.Set("active", isNtsc);
            this.PalRegionButton.Classes.Set("active", !isNtsc);

            this._suppressRegionToggle = false;
        }

        // ###########################################################################################
        // Positions a new popup on the same screen as the main window with a slight staggered offset.
        // ###########################################################################################
        private void PositionPopupOnSameScreen(Window popup)
        {
            popup.WindowStartupLocation = WindowStartupLocation.Manual;

            double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
            double w = this.Bounds.Width > 0 ? this.Bounds.Width : this._restoreWidth;
            double h = this.Bounds.Height > 0 ? this.Bounds.Height : this._restoreHeight;

            int centerX = this.Position.X + (int)((w * scaling) / 2);
            int centerY = this.Position.Y + (int)((h * scaling) / 2);

            var screen = this.Screens.All.FirstOrDefault(s =>
                centerX >= s.Bounds.X &&
                centerY >= s.Bounds.Y &&
                centerX < s.Bounds.X + s.Bounds.Width &&
                centerY < s.Bounds.Y + s.Bounds.Height)
                ?? this.Screens.Primary;

            if (screen != null)
            {
                int cascadeStep = (int)(32 * scaling);
                int maxCascade = (int)(256 * scaling);

                int offsetX = this._popupCascadeOffset * cascadeStep;
                int offsetY = this._popupCascadeOffset * cascadeStep;

                if (offsetX > maxCascade)
                {
                    this._popupCascadeOffset = 0;
                    offsetX = 0;
                    offsetY = 0;
                }

                // Base it off the owner window's position slightly indented
                int px = this.Position.X + (int)(40 * scaling) + offsetX;
                int py = this.Position.Y + (int)(40 * scaling) + offsetY;

                // Adjust slightly if it forces itself off the edges of this target screen
                if (px + (popup.Width * scaling) > screen.Bounds.Right)
                    px = screen.Bounds.X + offsetX;
                if (py + (popup.Height * scaling) > screen.Bounds.Bottom)
                    py = screen.Bounds.Y + offsetY;

                popup.Position = new PixelPoint(px, py);
                this._popupCascadeOffset++;
            }
        }

        // ###########################################################################################
        // Opens a component info popup according to user settings.
        // ###########################################################################################
        internal void OpenComponentInfoPopup(string boardLabel, string displayText)
        {
            string componentKey = $"{boardLabel}\u001F{displayText}";
            var boardData = this._currentBoardData;
            bool hasExplicitRegionComponents = ComponentListBuilder.HasExplicitRegionComponents(boardData);
            var images = boardData?.ComponentImages ?? new List<ComponentImageEntry>();
            var localFiles = boardData?.ComponentLocalFiles ?? new List<ComponentLocalFileEntry>();
            var links = boardData?.ComponentLinks ?? new List<ComponentLinkEntry>();
            var componentEntries = boardData?.Components
                .Where(c => string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                .ToList() ?? new List<ComponentEntry>();

            if (UserSettings.MultipleInstancesForComponentPopup)
            {
                if (!this._componentInfoWindowsByKey.TryGetValue(componentKey, out var popup) || !popup.IsVisible)
                {
                    popup = new ComponentInfoWindow();
                    this._componentInfoWindowsByKey[componentKey] = popup;

                    popup.Closed += (_, _) =>
                    {
                        if (this._componentInfoWindowsByKey.TryGetValue(componentKey, out var existing) && ReferenceEquals(existing, popup))
                            this._componentInfoWindowsByKey.Remove(componentKey);
                    };
                }

                popup.SetComponent(
                    boardLabel,
                    displayText,
                    componentEntries,
                    images,
                    localFiles,
                    links,
                    UserSettings.Region,
                    DataManager.DataRoot,
                    hasExplicitRegionComponents);

                popup.UpdateOscilloscopeSessionTitleState(
                    this.TabOscilloscopeControl.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                    this.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

                if (!popup.IsVisible)
                {
                    this.PositionPopupOnSameScreen(popup);
                    popup.Show(this);
                    popup.Focus();
                }
                else
                {
                    popup.Activate();
                    popup.Focus();
                }

                return;
            }

            if (this._singleComponentInfoWindow == null)
            {
                this._singleComponentInfoWindow = new ComponentInfoWindow();
                this._singleComponentInfoWindow.Closed += (_, _) => this._singleComponentInfoWindow = null;
            }

            this._singleComponentInfoWindow.CloseOnDeactivate = false;
            this._singleComponentInfoWindow.SetComponent(
                boardLabel,
                displayText,
                componentEntries,
                images,
                localFiles,
                links,
                UserSettings.Region,
                DataManager.DataRoot,
                hasExplicitRegionComponents);

            this._singleComponentInfoWindow.UpdateOscilloscopeSessionTitleState(
                this.TabOscilloscopeControl.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                this.TabOscilloscopeControl.HasActiveEstablishedOscilloscopeSessionForTitleState());

            if (!this._singleComponentInfoWindow.IsVisible)
            {
                // Single-instance popups reuse the last on-screen position (when not maximized)
                // instead of re-cascading from the main window every time they are reopened.
                if (UserSettings.HasComponentInfoWindowLayout &&
                    this._singleComponentInfoWindow.WindowState != Avalonia.Controls.WindowState.Maximized)
                {
                    this._singleComponentInfoWindow.Position =
                        new PixelPoint(UserSettings.ComponentInfoWindowX, UserSettings.ComponentInfoWindowY);
                }
                else
                {
                    this.PositionPopupOnSameScreen(this._singleComponentInfoWindow);
                }

                this._singleComponentInfoWindow.Show(this);
                this._singleComponentInfoWindow.Focus();
            }
            else
            {
                this._singleComponentInfoWindow.Activate();
                this._singleComponentInfoWindow.Focus();
            }
        }

        // ###########################################################################################
        // Handles "Blink selected" checkbox changes and refreshes highlight visuals immediately.
        // ###########################################################################################
        private void OnBlinkSelectedChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.BlinkSelected = BlinkSelectedCheckBox.IsChecked ?? false;

            this._blinkSelectedEnabled = this.BlinkSelectedCheckBox.IsChecked == true;

            bool hasBlinkEligibleSelection = this.TabSchematicsControl.HasBlinkEligibleSelection();
            bool hasComponentSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;

            if (this._blinkSelectedEnabled && hasBlinkEligibleSelection)
            {
                this._blinkSelectedPhaseVisible = false;
                this.TabSchematicsControl.ApplyHighlightVisuals(
                    hasComponentSelection,
                    this.GetCurrentBlinkFactor(true));
                this.UpdateBlinkTimer(true);
                return;
            }

            this._blinkSelectedPhaseVisible = true;
            this.UpdateBlinkTimer(hasBlinkEligibleSelection);
            this.TabSchematicsControl.ApplyHighlightVisuals(
                hasComponentSelection,
                this.GetCurrentBlinkFactor(hasBlinkEligibleSelection));
        }

        // ###########################################################################################
        // Starts or stops the blink timer depending on current checkbox state and selection state.
        // ###########################################################################################
        internal void UpdateBlinkTimer(bool hasSelection)
        {
            bool shouldBlink = this._blinkSelectedEnabled && hasSelection;

            if (!shouldBlink)
            {
                this._blinkSelectedTimer?.Stop();
                this._blinkSelectedPhaseVisible = true;
                return;
            }

            if (this._blinkSelectedTimer == null)
            {
                this._blinkSelectedTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(450)
                };
                this._blinkSelectedTimer.Tick += this.OnBlinkSelectedTimerTick;
            }

            if (!this._blinkSelectedTimer.IsEnabled)
                this._blinkSelectedTimer.Start();
        }

        // ###########################################################################################
        // Advances blink phase and re-applies highlight visuals while selection exists.
        // ###########################################################################################
        private void OnBlinkSelectedTimerTick(object? sender, EventArgs e)
        {
            bool hasBlinkEligibleSelection = this.TabSchematicsControl.HasBlinkEligibleSelection();
            bool hasComponentSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;

            if (!hasBlinkEligibleSelection)
            {
                this.UpdateBlinkTimer(false);
                this.TabSchematicsControl.ApplyHighlightVisuals(hasComponentSelection, 1.0);
                return;
            }

            this._blinkSelectedPhaseVisible = !this._blinkSelectedPhaseVisible;
            this.TabSchematicsControl.ApplyHighlightVisuals(
                hasComponentSelection,
                this.GetCurrentBlinkFactor(true));
        }

        // ###########################################################################################
        // Computes effective blink multiplier for current frame.
        // ###########################################################################################
        internal double GetCurrentBlinkFactor(bool hasSelection)
        {
            if (!hasSelection || !this._blinkSelectedEnabled)
                return 1.0;

            return this._blinkSelectedPhaseVisible ? 1.0 : 0.0;
        }

        // ###########################################################################################
        // Closes single popup when clicking the main window outside a component hit target.
        // ###########################################################################################
        private void OnMainPointerPressedCloseSinglePopup(object? sender, PointerPressedEventArgs e)
        {
            if (UserSettings.MultipleInstancesForComponentPopup)
                return;

            var popup = this._singleComponentInfoWindow;
            if (popup == null || !popup.IsVisible)
                return;

            if (this.isHoveringComponent)
                return;

            popup.Close();
        }

        // ###########################################################################################
        // Cancels "Add worklog" area-marking mode on Escape, closes the single component popup on
        // Escape otherwise, and toggles the schematics fullscreen window on F11 (ToggleSchematicsFullscreenWindow -
        // opens it if closed, closes it if already open, so repeatedly pressing F11 here behaves the
        // same as pressing Escape while it is open).
        //
        // WHY THE WORKLOG BRANCH IS HERE AND NOT ONLY IN TabSchematics: the mode is started by the
        // "Add worklog" button in this bar, which KEEPS keyboard focus afterwards, so the tab's own
        // OnSchematicsKeyDown - which also handles Escape for this mode - never receives the press
        // unless the user first clicks the schematic. Escape did nothing in exactly the moment it
        // is most wanted: right after clicking the button, before committing to a drag. This
        // handler is registered on the WINDOW at the Tunnel phase, so it sees the key wherever
        // focus happens to be.
        //
        // The tab's own handler is deliberately KEPT rather than replaced: it is the one that runs
        // once the user has clicked into the schematic (which takes focus), and it also covers the
        // fullscreen schematics window, which is a different window and never routes through here.
        // CancelWorklogEntryMode is safe to call twice, so the overlap costs nothing.
        //
        // The DETACHED THUMBNAILS window is a THIRD window and also never routes through here - see
        // SchematicsThumbnailsWindow's own F11 handler, which calls this same
        // ToggleSchematicsFullscreenWindow via TabSchematics.MainWindow.
        //
        // Ordered ABOVE the popup close, and gated on the mode actually being active so Escape
        // still reaches the popup at every other time: while an area is being marked the mode is
        // what Escape means, and cancelling both at once on one press would be two undos for one
        // keystroke.
        // ###########################################################################################
        private void OnMainKeyDownCloseSinglePopup(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.F11)
            {
                this.ToggleSchematicsFullscreenWindow();
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Escape)
                return;

            if (this.TabSchematicsControl.IsWorklogEntryModeActive)
            {
                // Resets this bar's own buttons too, via the ResetWorklogEntryModeButtons callback
                // the tab makes on the way out - so cancelling by keyboard leaves the bar in
                // exactly the state the Cancel button would.
                this.TabSchematicsControl.CancelWorklogEntryMode();
                e.Handled = true;
                return;
            }

            if (UserSettings.MultipleInstancesForComponentPopup)
                return;

            var popup = this._singleComponentInfoWindow;
            if (popup == null || !popup.IsVisible)
                return;

            popup.Close();
            e.Handled = true;
        }

        // ###########################################################################################
        // Updates the UI with info specific to the current board's revision date and credits.
        // ###########################################################################################
        private void PopulateBoardInfoSection(string? revisionDate, List<CreditEntry>? credits)
        {
            this.TabAbout.SetBoardInfo(revisionDate, credits);
        }

        // ###########################################################################################
        // Refreshes the component list and highlight data for the current local region.
        // ###########################################################################################
        private void RefreshImages()
        {
            _ = this.ApplyRegionFilterAsync();
        }

        // ###########################################################################################
        // Refresh the component list according to the active region while recovering any matching
        // existing selection, similar to category filter switching.
        // ###########################################################################################
        private async Task ApplyRegionFilterAsync()
        {
            if (this._currentBoardData == null)
                return;

            this.SetComponentHighlightRects(await Task.Run(() =>
                HighlightRectBuilder.BuildHighlightRects(this._currentBoardData, this._localRegion)));

            var previouslySelectedKeys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                    .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try { this.ComponentFilterListBox.SelectAll(); } catch { }
            }
            else
            {
                for (int i = 0; i < componentItems.Count; i++)
                {
                    if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                        this.ComponentFilterListBox.Selection.Select(i);
                }
            }
            this._suppressComponentHighlightUpdate = false;

            var survivingLabels = componentItems
                .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                .Select(item => item.BoardLabel)
                .Where(l => !string.IsNullOrEmpty(l))
                .ToList();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                survivingLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();
            }

            this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
            this.TabContribute.LoadData(this._currentBoardData, this._localRegion);
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // ###########################################################################################
        // Refreshes the component filter list based on search text, or switches to label-editor
        // search behavior while the component label editor is active.
        // ###########################################################################################
        public void OnComponentSearchTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            if (this._suppressComponentSearchRefresh || this._currentBoardData == null || this._suppressCategoryFilterSave)
                return;

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                this.TabSchematicsControl.ApplyLabelEditorSearchFilter(searchTerm);
                return;
            }

            this.ApplyNormalComponentSearchFilter(searchTerm);
        }

        // ###########################################################################################
        // Applies the normal component search behavior used outside the label editor.
        // ###########################################################################################
        private void ApplyNormalComponentSearchFilter(string searchTerm)
        {
            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData!, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            var highlightLabels = new List<string>();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try { this.ComponentFilterListBox.SelectAll(); } catch { }

                highlightLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(label => !string.IsNullOrEmpty(label))
                    .ToList();
            }

            this._suppressComponentHighlightUpdate = false;

            this.TabSchematicsControl.UpdateHighlightsForComponents(highlightLabels);

            // Forward the search term to filter the Overview tab's list
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // ###########################################################################################
        // Updates the component search box text hint so its current behavior is obvious.
        // ###########################################################################################
        internal void UpdateComponentSearchTextBoxMode()
        {
            if (this.ComponentSearchTextBox == null)
            {
                return;
            }

            this.ComponentSearchTextBox.PlaceholderText = this.TabSchematicsControl.IsLabelEditorActive
                ? "Find component label or category"
                : "Filter components";
        }

        // ###########################################################################################
        // Returns true when the current board has at least one component explicitly tagged as PAL or NTSC.
        // ###########################################################################################
        internal bool CurrentBoardHasExplicitRegionComponents()
        {
            return ComponentListBuilder.HasExplicitRegionComponents(this._currentBoardData);
        }

        // ###########################################################################################
        // Creates placeholder content for the Schematics tab while fullscreen mode is active.
        // Keeps the thumbnail list available so another schematic can be selected without closing
        // the fullscreen window first.
        // ###########################################################################################
        private SchematicsFullscreenPlaceholder CreateSchematicsFullscreenPlaceholder()
        {
            double ratio = 0.70;
            var boardKey = this.GetCurrentBoardKey();
            if (!string.IsNullOrWhiteSpace(boardKey))
            {
                ratio = Math.Clamp(UserSettings.GetSchematicsSplitterRatio(boardKey), 0.1, 0.9);
            }

            var hostedThumbnailList = this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList");

            var placeholder = new SchematicsFullscreenPlaceholder();
            placeholder.Initialize(this.TabSchematicsControl.currentThumbnails, hostedThumbnailList, ratio);
            return placeholder;
        }

        // ###########################################################################################
        // Opens the existing schematics viewer in a separate maximized window. internal (not
        // private) so SchematicsThumbnailsWindow's own F11 handler can reach it via TabSchematics.
        // MainWindow - see ToggleSchematicsFullscreenWindow's header for why F11 is a toggle from
        // both windows rather than an open-only action here.
        // ###########################################################################################
        internal void OpenSchematicsFullscreenWindow()
        {
            if (this._schematicsFullscreenWindow != null)
            {
                if (this._schematicsFullscreenWindow.WindowState == Avalonia.Controls.WindowState.Minimized)
                    this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Normal;

                this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Maximized;
                this._schematicsFullscreenWindow.Activate();
                this._schematicsFullscreenWindow.Focus();
                return;
            }

            this.TabSchematicsControl.EnterFullscreenMode();
            this.SchematicsTabItem.Content = this.CreateSchematicsFullscreenPlaceholder();

            this._schematicsFullscreenWindow = new SchematicsFullscreenWindow(
                this.TabSchematicsControl,
                this.RestoreSchematicsTabContent);

            // Set on Closing rather than around the Close() call sites, because this window closes
            // itself too: Escape and F11 are handled inside SchematicsFullscreenWindow and never pass
            // through Main. Closing runs before the owned thumbnails window is torn down, which is
            // what the thumbnails Closed handler needs in order to tell that cascade apart from the
            // user dismissing it.
            this._schematicsFullscreenWindow.Closing += (_, _) =>
            {
                this._isClosingFullscreenWindow = true;
            };

            this._schematicsFullscreenWindow.Closed += (_, _) =>
            {
                this._isClosingFullscreenWindow = false;
                this._schematicsFullscreenWindow = null;

                // Hand the detached thumbnails window back to THIS window now that the fullscreen
                // window it was following is gone - see ReownSchematicsThumbnailsWindow's header for
                // why it must always own the on-screen schematics surface.
                //
                // Re-owning ALONE is not enough here: closing the fullscreen window closes the
                // windows it owns, so by the time this runs the thumbnails window is usually already
                // gone and the field already null, leaving nothing to re-own. So this reopens it when
                // the setting still says detached, and re-owns it when it somehow survived.
                this.ReopenThumbnailsWindowAfterFullscreenClosed();
            };

            this.PositionFullscreenWindowOnSameScreen(this._schematicsFullscreenWindow);
            this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Maximized;
            this._schematicsFullscreenWindow.Show();

            // The detached thumbnails window (if open) is normally owned by THIS window, which on
            // Windows means clicking into it raises its owner - this window - above the fullscreen
            // window it should stay behind. Re-own it to the fullscreen window instead.
            this.ReownSchematicsThumbnailsWindow();

            this.TabSchematicsControl.RefreshAfterHostChanged();
            this._schematicsFullscreenWindow.Focus();
        }

        // ###########################################################################################
        // Reported: with the thumbnails window detached to its own monitor, clicking into it while
        // schematics fullscreen was active brought THIS window (Main) to the front instead of the
        // fullscreen window, because the thumbnails window's Owner was still Main. Window ownership on
        // Windows is real Win32 owner/HWND parenting (WS_POPUP + GWL_HWNDPARENT), and activating an
        // owned window raises its owner along with it - so the thumbnails window must be owned by
        // whichever window is CURRENTLY the visible schematics surface: the fullscreen window while
        // one is open, this window otherwise. Avalonia's Owner property is live and reassignable after
        // Show() (no close/reopen needed), so this just flips it to match.
        //
        // A no-op when the thumbnails window is not open - there is nothing to re-own, and the next
        // OpenSchematicsThumbnailsWindow call already owns it correctly (Show(this) while fullscreen
        // is inactive, or the ApplyThumbnailsDetachedState call sequence below when it is).
        // ###########################################################################################
        // ###########################################################################################
        // Restores the detached thumbnails window after the schematics fullscreen window has closed.
        //
        // Reported shape of the bug this fixes: while fullscreen is open the thumbnails window is
        // OWNED by it, and closing an owner closes what it owns - so leaving fullscreen destroyed the
        // thumbnails window, and its Closed handler read that as the user dismissing it and silently
        // turned detached mode off. The compensating ReownSchematicsThumbnailsWindow() ran from the
        // fullscreen Closed handler, i.e. after the thumbnails field was already null, so it found
        // nothing to re-own and never prevented anything.
        //
        // The Closed handler now leaves the setting alone during that cascade (_isClosingFullscreenWindow),
        // which makes reopening here the whole of the restore: the setting still says detached, so
        // ApplyThumbnailsDetachedState opens a fresh window - owned by this window, since
        // _schematicsFullscreenWindow is already null by now. A window that somehow survived the
        // cascade is only re-owned, not reopened.
        // ###########################################################################################
        private void ReopenThumbnailsWindowAfterFullscreenClosed()
        {
            if (this._isApplicationShuttingDown)
                return;

            if (this._schematicsThumbnailsWindow != null)
            {
                this.ReownSchematicsThumbnailsWindow();
                return;
            }

            if (!this.thisWindowOpenedCompletionSource.Task.IsCompleted)
                return;

            this.ApplyThumbnailsDetachedState();
        }

        private void ReownSchematicsThumbnailsWindow()
        {
            if (this._schematicsThumbnailsWindow == null)
                return;

            Window desiredOwner = this._schematicsFullscreenWindow ?? (Window)this;

            if (!ReferenceEquals(this._schematicsThumbnailsWindow.Owner, desiredOwner))
            {
                this._schematicsThumbnailsWindow.SetOwnerWindow(desiredOwner);
            }
        }

        // ###########################################################################################
        // The ONE place F11 is interpreted, from EITHER window it can be pressed in (this window's
        // own OnMainKeyDownCloseSinglePopup, and SchematicsThumbnailsWindow's matching handler when
        // that detached window is the active one). Toggles: closes the fullscreen window if one is
        // already open, opens it otherwise - so F11 behaves the same as Escape-to-close while also
        // working as a single repeatable key, which is what was asked for. Escape's own handling in
        // SchematicsFullscreenWindow is untouched.
        // ###########################################################################################
        internal void ToggleSchematicsFullscreenWindow()
        {
            if (this._schematicsFullscreenWindow != null)
            {
                this._schematicsFullscreenWindow.Close();
                return;
            }

            this.OpenSchematicsFullscreenWindow();
        }

        // ###########################################################################################
        // Opens schematics fullscreen from the left-side button.
        // ###########################################################################################
        private void OnFullscreenSchematicsClick(object? sender, RoutedEventArgs e)
        {
            this.OpenSchematicsFullscreenWindow();
        }

        // ###########################################################################################
        // Restores the schematics control back into the normal tab after fullscreen closes.
        // ###########################################################################################
        private void RestoreSchematicsTabContent(Control hostedContent)
        {
            if (!ReferenceEquals(hostedContent, this.TabSchematicsControl))
                return;

            if (!ReferenceEquals(this.SchematicsTabItem.Content, hostedContent))
                this.SchematicsTabItem.Content = hostedContent;

            this.TabSchematicsControl.ExitFullscreenMode();
            this.TabSchematicsControl.RefreshAfterHostChanged();
        }

        // ###########################################################################################
        // Places the fullscreen window on the same screen as the main window before maximizing it.
        // ###########################################################################################
        private void PositionFullscreenWindowOnSameScreen(Window window)
        {
            window.WindowStartupLocation = WindowStartupLocation.Manual;

            double scaling = this.RenderScaling > 0 ? this.RenderScaling : 1.0;
            double w = this.Bounds.Width > 0 ? this.Bounds.Width : this._restoreWidth;
            double h = this.Bounds.Height > 0 ? this.Bounds.Height : this._restoreHeight;

            int centerX = this.Position.X + (int)((w * scaling) / 2);
            int centerY = this.Position.Y + (int)((h * scaling) / 2);

            var screen = this.Screens.All.FirstOrDefault(s =>
                centerX >= s.Bounds.X &&
                centerY >= s.Bounds.Y &&
                centerX < s.Bounds.X + s.Bounds.Width &&
                centerY < s.Bounds.Y + s.Bounds.Height)
                ?? this.Screens.Primary;

            if (screen != null)
            {
                window.Position = new PixelPoint(
                    screen.Bounds.X + 100,
                    screen.Bounds.Y + 100);
            }
        }

        // ###########################################################################################
        // Opens a maximized contribution editor window for the selected component.
        // ###########################################################################################
        internal void OpenComponentContributionWindow(string boardLabel)
        {
            if (this._currentBoardData == null || string.IsNullOrWhiteSpace(boardLabel))
            {
                return;
            }

            this.ShowComponentContributionWindow(window => window.LoadComponent(
                this._currentBoardData,
                DataManager.DataRoot,
                this.HardwareComboBox.SelectedItem as string ?? string.Empty,
                this.BoardComboBox.SelectedItem as string ?? string.Empty,
                this._localRegion,
                boardLabel,
                this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty));
        }

        // ###########################################################################################
        // Opens the contribution editor on a component that does not exist in the board data yet,
        // so a missing component can be suggested from scratch.
        // ###########################################################################################
        internal void OpenNewComponentContributionWindow()
        {
            if (this._currentBoardData == null)
            {
                return;
            }

            this.ShowComponentContributionWindow(window => window.LoadNewComponent(
                this._currentBoardData,
                DataManager.DataRoot,
                this.HardwareComboBox.SelectedItem as string ?? string.Empty,
                this.BoardComboBox.SelectedItem as string ?? string.Empty,
                this._localRegion,
                this.GetCurrentBoardEntry()?.ExcelDataFile ?? string.Empty));
        }

        // ###########################################################################################
        // Creates the contribution editor, lets the caller load it, and shows it maximized on the
        // screen the main window is on.
        // ###########################################################################################
        private void ShowComponentContributionWindow(Action<ComponentContributionWindow> loadContent)
        {
            var window = new ComponentContributionWindow();
            loadContent(window);

            this.PositionFullscreenWindowOnSameScreen(window);
            window.WindowState = Avalonia.Controls.WindowState.Maximized;
            window.Show(this);
            window.Focus();
        }

        // ###########################################################################################
        // Pushes the current oscilloscope session title state into any open component info popup
        // windows so their title suffix stays aligned with the oscilloscope tab.
        // ###########################################################################################
        internal void UpdateComponentInfoWindowsOscilloscopeSessionState(
            bool hasSeenOscilloscopeSession,
            bool hasActiveOscilloscopeSession)
        {
            this._singleComponentInfoWindow?.UpdateOscilloscopeSessionTitleState(
                hasSeenOscilloscopeSession,
                hasActiveOscilloscopeSession);

            foreach (var popup in this._componentInfoWindowsByKey.Values)
            {
                popup.UpdateOscilloscopeSessionTitleState(
                    hasSeenOscilloscopeSession,
                    hasActiveOscilloscopeSession);
            }
        }

        // ###########################################################################################
        // Rebuilds runtime component lists and highlight caches after the schematic label editor
        // has modified the in-memory board data for the current session.
        // ###########################################################################################
        internal void RefreshRuntimeBoardStateAfterLabelEditorApply()
        {
            if (this._currentBoardData == null)
            {
                return;
            }

            var previouslySelectedKeys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                    .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            string searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

            this.SetComponentHighlightRects(
                HighlightRectBuilder.BuildHighlightRects(this._currentBoardData, this._localRegion));

            var componentItems = ComponentListBuilder.BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                // ItemsSource was just reassigned above; SelectAll can throw if Avalonia's
                // selection model has not yet caught up with the new item count.
                try
                {
                    this.ComponentFilterListBox.SelectAll();
                }
                catch
                {
                }
            }
            else
            {
                for (int i = 0; i < componentItems.Count; i++)
                {
                    if (previouslySelectedKeys.Contains(componentItems[i].SelectionKey))
                    {
                        this.ComponentFilterListBox.Selection.Select(i);
                    }
                }
            }

            this._suppressComponentHighlightUpdate = false;

            var survivingLabels = componentItems
                .Where(item => previouslySelectedKeys.Contains(item.SelectionKey))
                .Select(item => item.BoardLabel)
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                survivingLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(label => !string.IsNullOrWhiteSpace(label))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            this.TabSchematicsControl.UpdateHighlightsForComponents(survivingLabels);
            this.TabSchematicsControl.UpdateComponentLabels();
            this.TabOverview.LoadData(this._currentBoardData);
            this.TabContribute.LoadData(this._currentBoardData, this._localRegion);
        }

        // ###########################################################################################
        // Reacts to "Check for new or updated data at application launch" changes.
        // Re-enabling hides the local-edit warning banner, but disabling from Configuration must
        // not show it. That banner is only shown explicitly after label-editor apply.
        // ###########################################################################################
        private void OnCheckDataOnLaunchSettingChanged(bool isEnabled)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.UpdateDataSyncStatusIcon();

                if (isEnabled)
                {
                    this.HideDataSyncDisabledBanner();
                }
            });
        }

        // ###########################################################################################
        // Reacts to the Configuration tab's "Show all workbooks" / "Show this board's workbooks"
        // radio group being flipped, by rebuilding the Workbooks tab through the one funnel every
        // other worklog change already uses.
        //
        // Needed because the scope decides WHICH workbooks RefreshWorklogBar passes the tab, so the
        // tab cannot re-derive it on its own without a refresh. Before this the event was raised and
        // subscribed by nothing: the setting only appeared to work because switching to the
        // Workbooks tab happens to refresh it on attach - incidental, and no help at all to a
        // Workbooks tab that is already on screen when the setting changes.
        //
        // Posted rather than run inline, matching OnCheckDataOnLaunchSettingChanged: the event is
        // raised from a UserSettings setter, so this keeps the rebuild off the setter's own stack.
        // ###########################################################################################
        private void OnWorkbooksScopeSettingChanged()
        {
            Dispatcher.UIThread.Post(() => this.RefreshWorklogBar());
        }

        // ###########################################################################################
        // Reprints every cost on screen when the Configuration tab's currency drop-down changes,
        // through the same funnel and for the same reason as the scope handler above.
        //
        // The Workbooks tab's summary strip, its entry cards and the worklog editor all build their
        // figures once per refresh, so a currency changed while that tab is on screen would leave
        // the old code beside every number until something else happened to rebuild them - and the
        // most likely moment to change this setting is right after noticing the wrong currency on
        // exactly that screen.
        //
        // Posted rather than run inline, so the rebuild stays off the UserSettings setter's stack.
        // ###########################################################################################
        private void OnWorklogCurrencySettingChanged()
        {
            Dispatcher.UIThread.Post(() => this.RefreshWorklogBar());
        }

        // ###########################################################################################
        // Shows a dismissable main-window banner explaining that launch-time data synchronization
        // is disabled and must be re-enabled manually in Configuration when appropriate.
        // ###########################################################################################
        private void ShowDataSyncDisabledBanner()
        {
            this.SyncBannerText.Text =
                "Data synchronization has been disabled because the board Excel data was edited locally. Re-enable it in the \"Configuration\" tab when safe.";
            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBanner.IsVisible = true;
            this._isShowingDataSyncDisabledBanner = true;
        }

        // ###########################################################################################
        // Hides the synchronization-disabled banner without affecting other sync banner flows.
        // ###########################################################################################
        private void HideDataSyncDisabledBanner()
        {
            if (!this._isShowingDataSyncDisabledBanner)
            {
                return;
            }

            this.SyncBanner.IsVisible = false;
            this.SyncBannerRefreshButton.IsVisible = false;
            this._isShowingDataSyncDisabledBanner = false;
        }

        // ###########################################################################################
        // Disables launch-time data synchronization after a local board Excel edit, updates the
        // Configuration tab checkbox, and shows a warning dialog to explain the safety change.
        // ###########################################################################################
        internal async Task DisableLaunchDataSyncAfterLocalBoardEditAsync()
        {
            if (!UserSettings.CheckDataOnLaunch)
            {
                this.ShowDataSyncDisabledBanner();
                return;
            }

            this.TabConfiguration.SetCheckDataOnLaunchCheckBoxValue(false);
            UserSettings.CheckDataOnLaunch = false;

            await this.ShowDataSyncDisabledAfterLocalBoardEditDialogAsync();
        }

        // ###########################################################################################
        // Shows a modal warning dialog explaining why launch-time data synchronization was turned
        // off after the component label editor saved changes to the board Excel file.
        // ###########################################################################################
        private async Task ShowDataSyncDisabledAfterLocalBoardEditDialogAsync()
        {
            var closeButton = new Button
            {
                Content = "OK",
                MinWidth = 110,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center
            };

            var dialog = new Window
            {
                Title = "Data synchronization disabled",
                Width = 520,
                MinWidth = 420,
                CanResize = false,
                ShowInTaskbar = false,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            closeButton.Click += (_, _) => dialog.Close();

            dialog.Content = new Border
            {
                Padding = new Thickness(18),
                Child = new StackPanel
                {
                    Spacing = 14,
                    Children =
                    {
                        new TextBlock
                        {
                            Text =
                                "To avoid potential data loss, \"Check for new or updated data at application launch\" has been disabled because the board Excel data was changed by the component label editor.",
                            TextWrapping = TextWrapping.Wrap
                        },
                        new TextBlock
                        {
                            Text =
                                "A banner is now shown in the main window to indicate that synchronization is disabled.",
                            TextWrapping = TextWrapping.Wrap
                        },
                        closeButton
                    }
                }
            };

            await dialog.ShowDialog(this);
        }

        // ###########################################################################################
        // Enables or disables hardware/board navigation while the schematics label editor owns the board state.
        // ###########################################################################################
        internal void SetSchematicsEditorNavigationEnabled(bool isEnabled)
        {
            this.HardwareComboBox.IsEnabled = isEnabled;
            this.BoardComboBox.IsEnabled = isEnabled;
        }

        // ###########################################################################################
        // Shows or hides the main-window banner indicating that the schematics label editor owns
        // the current board state and navigation is temporarily locked.
        // ###########################################################################################
        internal void SetSchematicsLabelEditorModeBannerVisible(bool isVisible)
        {
            this.SchematicsLabelEditorModeBanner.IsVisible = isVisible;
        }

        // ###########################################################################################
        // Updates the global launch-time data-sync status icon, clickability and tooltip in the main window.
        // When launch-time sync is disabled, hovering the icon temporarily shows the manual refresh icon.
        // ###########################################################################################
        private void UpdateDataSyncStatusIcon()
        {
            bool isEnabled = UserSettings.CheckDataOnLaunch;
            bool isCheckingOnline = this._dataSyncStatusIconSpinRequestCount > 0;
            bool allowManualRefreshWhileDisabled = !isEnabled && this._isHoveringDataSyncStatusIcon;
            bool isClickable = !isCheckingOnline;

            this.DataSyncStatusIconTextBlock.IsVisible = !isCheckingOnline;
            this.DataSyncStatusSpinnerCanvas.IsVisible = isCheckingOnline;

            this.DataSyncStatusIconTextBlock.Text = isEnabled || allowManualRefreshWhileDisabled
                ? "\uf021"
                : "\uf05e";

            if (this.TryFindResource(isEnabled || allowManualRefreshWhileDisabled ? "Text_Success_Fg" : "Text_Fail_Fg", out var brushResource) &&
                brushResource is IBrush brush)
            {
                this.DataSyncStatusIconTextBlock.Foreground = brush;
                this.DataSyncStatusSpinnerEllipse.Stroke = brush;
            }
            else
            {
                IBrush fallbackBrush = isEnabled || allowManualRefreshWhileDisabled
                    ? Brushes.ForestGreen
                    : Brushes.IndianRed;

                this.DataSyncStatusIconTextBlock.Foreground = fallbackBrush;
                this.DataSyncStatusSpinnerEllipse.Stroke = fallbackBrush;
            }

            this.DataSyncStatusIconBorder.Cursor = new Cursor(
                isClickable
                    ? StandardCursorType.Hand
                    : StandardCursorType.Arrow);

            ToolTip.SetTip(
                this.DataSyncStatusIconBorder,
                isCheckingOnline
                    ? $"Checking data from {AppConfig.GetOnlineSourceLabel()}..."
                    : isEnabled
                        ? "Data update is enabled. Click to refresh data now"
                        : this._isHoveringDataSyncStatusIcon
                            ? "Data update at launch is disabled. Click to run a manual refresh now"
                            : "Data update at launch is disabled. Hover to show manual refresh");
        }

        // ###########################################################################################
        // Handles clicks on the top-right data-sync status icon.
        // Clicking always allows a manual refresh unless a sync is already in progress.
        // ###########################################################################################
        private async void OnDataSyncStatusIconPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                return;
            }

            e.Handled = true;

            // Allow the banner to update with progress + current file (2-line banner).
            await this.CheckForDataUpdatesNowAsync(keepBannerTextStatic: false);
        }

        // ###########################################################################################
        // Shows the manual refresh icon when hovering the disabled data-sync status icon.
        // ###########################################################################################
        private void OnDataSyncStatusIconPointerEntered(object? sender, PointerEventArgs e)
        {
            this._isHoveringDataSyncStatusIcon = true;
            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Restores the normal disabled icon when the pointer leaves the data-sync status icon.
        // ###########################################################################################
        private void OnDataSyncStatusIconPointerExited(object? sender, PointerEventArgs e)
        {
            this._isHoveringDataSyncStatusIcon = false;
            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Starts animating the top-right data-sync status spinner while online data checks are running.
        // Nested calls are reference-counted so overlapping sync flows do not stop the spinner early.
        // ###########################################################################################
        private void StartDataSyncStatusIconSpin()
        {
            if (!UserSettings.CheckDataOnLaunch)
            {
                return;
            }

            this._dataSyncStatusIconSpinRequestCount++;

            if (this._dataSyncStatusIconSpinTimer == null)
            {
                this._dataSyncStatusIconSpinTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(25)
                };

                this._dataSyncStatusIconSpinTimer.Tick += (_, _) =>
                {
                    this._dataSyncStatusIconSpinAngle = (this._dataSyncStatusIconSpinAngle + 1.5) % 52.0;
                    this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = -this._dataSyncStatusIconSpinAngle;
                };
            }

            if (!this._dataSyncStatusIconSpinTimer.IsEnabled)
            {
                this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = -this._dataSyncStatusIconSpinAngle;
                this._dataSyncStatusIconSpinTimer.Start();
            }

            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Stops animating the top-right data-sync status spinner after online data checks finish.
        // ###########################################################################################
        private void StopDataSyncStatusIconSpin()
        {
            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                this._dataSyncStatusIconSpinRequestCount--;
            }

            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                return;
            }

            this._dataSyncStatusIconSpinAngle = 0.0;
            this._dataSyncStatusIconSpinTimer?.Stop();
            this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = 0.0;

            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Starts the background data validation task and converts failures into log entries only.
        // ###########################################################################################
        private static Task StartBackgroundDataValidationAsync()
        {
            return Task.Run(async () =>
            {
                try
                {
                    await DataValidator.ValidateAllDataAsync();
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Background data validation failed - [{ex.Message}]");
                }
            });
        }

        // ###########################################################################################
        // Schedules orphan/non-used file cleanup once per application session when both launch-time
        // sync and orphan cleanup are enabled.
        // ###########################################################################################
        internal void ScheduleOrphanAndUnusedFileCleanupIfEnabled()
        {
            if (!UserSettings.CheckDataOnLaunch ||
                !UserSettings.AllowDeletionOfOrphanAndNonUsedFiles ||
                this.thisHasScheduledOrphanAndUnusedFileCleanup)
            {
                return;
            }

            this.thisHasScheduledOrphanAndUnusedFileCleanup = true;
            _ = this.RunOrphanAndUnusedFileCleanupAfterStartupAsync();
        }

        // ###########################################################################################
        // Waits until startup UI work, background sync, and background validation are all settled before
        // running the orphan/non-used file cleanup as a low-priority background task.
        // ###########################################################################################
        private async Task RunOrphanAndUnusedFileCleanupAfterStartupAsync()
        {
            try
            {
                await this.thisWindowOpenedCompletionSource.Task;
                await this.thisBackgroundStartupSyncCompletionSource.Task;

                if (this.thisBackgroundDataValidationTask != null)
                {
                    await this.thisBackgroundDataValidationTask;
                }

                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);

                await Task.Delay(TimeSpan.FromSeconds(15));

                if (!UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    return;
                }

                await DataManager.DeleteOrphanAndUnusedFilesAsync();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Background orphan/non-used file cleanup failed - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Rebuilds the sync banner text as two lines:
        // Line 1: status/progress only
        // Line 2: current relative path being transferred (single-line, optional)
        // ###########################################################################################
        private void SetSyncBannerText(string statusLine)
        {
            string cleanStatus = (statusLine ?? string.Empty)
                .Replace("\r", string.Empty, StringComparison.Ordinal)
                .Replace("\n", " ", StringComparison.Ordinal)
                .Trim();

            string cleanFile = (this._currentSyncFileRelativePath ?? string.Empty)
                .Replace("\r", string.Empty, StringComparison.Ordinal)
                .Replace("\n", string.Empty, StringComparison.Ordinal)
                .Trim();

            this.SyncBannerText.Text = string.IsNullOrWhiteSpace(cleanFile)
                ? cleanStatus
                : $"{cleanStatus}\n{cleanFile}";
        }

        // ###########################################################################################
    }
}