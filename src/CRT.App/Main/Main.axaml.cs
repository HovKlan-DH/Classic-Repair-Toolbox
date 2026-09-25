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
using Handlers.Online;
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
        // ###########################################################################################
        // Main is split across partial-class files by area; this file owns construction, window
        // lifecycle (position/size/placement/closing), the tab-visibility switches, and the small
        // shared helpers several other files call into. The others:
        //
        //   Main.Updates.cs           - app update banner, install, data-sync banner and flow
        //   Main.BoardSelection.cs    - hardware/board drop-downs, board loading, catalogue visibility
        //   Main.SchematicsWindows.cs - detached thumbnails window, fullscreen schematics window
        //   Main.Worklog.cs           - worklog bar, workbook activation, worklog entry mode
        //   Main.ComponentPopup.cs    - component filter/search, component info popup, blink, region
        //   Main.DataSyncStatus.cs    - sync settings reactions, sync status icon, background cleanup
        //   Main.ModeHint.cs          - the mode-hint label shown while a mode awaits user action
        //   Main.NewSystem.cs         - "Add a new system": the create dialog, draft creation, navigation
        //   Main.DraftDrift.cs        - the "official data moved under your draft" banner and report
        //   Main.DraftBadges.cs       - the "Draft" chip on Hardware and Board drop-down entries
        // ###########################################################################################

        // Window placement: tracks the last known normal-state size and position
        private double _restoreWidth;
        private double _restoreHeight;
        private PixelPoint _restorePosition;
        private DispatcherTimer? _windowPlacementSaveTimer;
        private bool _windowPlacementReady = false;

        // Category filter: suppresses saves during programmatic selection changes
        private bool _suppressCategoryFilterSave;
        private bool _suppressComponentSearchRefresh;

        // ###########################################################################################
        // Suppresses OnHardwareSelectionChanged/OnBoardSelectionChanged while ApplyCatalogueVisibility
        // re-populates the drop-downs' ItemsSource for a hardware/board that is NOT the one currently
        // selected. Reassigning ItemsSource momentarily clears SelectedItem even when the caller then
        // restores the same selected value, which still raises SelectionChanged - and without this
        // flag that cascades into a full board reload (Excel re-read, every schematic bitmap
        // re-decoded, the detached thumbnails window closed and reopened) for a checkbox toggle on
        // hardware/board that has nothing to do with what is on screen. See ApplyCatalogueVisibility.
        // ###########################################################################################
        private bool _suppressBoardHardwareSelectionReload;

        private BoardData? _currentBoardData;

        // Snapshotted alongside _currentBoardData in LoadSelectedBoardAsync, from
        // DataManager.LastLoadedDraftSummary right after that same board's LoadBoardDataAsync call
        // - so every component-list rebuild (there are several call sites: category filter change,
        // search box, "mark all", the component popup) marks drafted rows against the draft that
        // is ACTUALLY behind the board on screen, not whatever DataManager's static happens to hold
        // at the moment of a later rebuild (which could differ if another board load raced in
        // between). BoardDraftSummary.Empty until a board has ever loaded.
        private BoardDraftSummary _currentBoardDraftSummary = BoardDraftSummary.Empty;

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

        // Set once the contributor has answered the Drafts tab's "unsaved table edits" prompt on
        // the way out (saved them, or chose to discard them), so the Close() that follows is not
        // stopped by the same prompt again. See OnWindowClosing.
        private bool _unsavedTableEditsSettledForExit;

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
            this.TabDrafts.Initialize(this);
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

                    if (this.ShouldReturnFocusToComponentSearch())
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

                // Drafts (session 2b) are listed across every board, not just the one just
                // selected, so this only needs DataManager.HardwareBoards to exist - which
                // PopulateHardwareDropDown just guaranteed - rather than a board-specific load.
                // Calling it in the constructor would always find an empty HardwareBoards list
                // (see this method's own header comment on why that population moved here).
                this.ApplyDraftsTabVisibility();

                if (UserSettings.CheckVersionOnLaunch)
                {
                    _ = this.CheckForAppUpdateNowAsync();
                }

                await this.StartBackgroundSyncAsync();

                // ###########################################################################################
                // ASK THE SERVER WHETHER A REVIEWER HAS REPLIED (maintainer report, 2026-09-22).
                //
                // *** THIS ONLY EVER RAN FROM A BUTTON BEFORE. *** Receipts were loaded from disk at
                // startup, but nothing asked the server until the contributor opened "My submissions"
                // and pressed Refresh - so a request for changes sat unseen behind a button they had
                // no reason to press. Launching the app is exactly when somebody expects to be told.
                //
                // NOT AWAITED, on purpose. It is a network call whose answer changes nothing the
                // user is currently looking at, and startup timing is a logged milestone
                // (StartupTimeline) that a server round trip must not enter into. It costs nothing
                // at all for a user who has never submitted: RefreshQuietlyAsync returns without a
                // request when there is nothing still open to ask about.
                //
                // Every failure inside is swallowed and logged there - see its header for why an
                // unobserved faulted Task would be the worst possible shape for this.
                //
                // ###########################################################################################
                // RETIRE ANY DRAFT THE PUBLISHED DATA HAS NOW ABSORBED (maintainer request,
                // 2026-09-23).
                //
                // *** HERE, RATHER THAN AT THE MOMENT OF PUBLISHING, BECAUSE THE TEST NEEDS THE SYNCED
                // BOARD. *** The rule is not "the server says published", it is "the draft says
                // nothing the published board does not" - see DraftRetirement. That comparison can
                // only be made once the sync above has brought the new board down, which is why this
                // sits after the awaited StartBackgroundSyncAsync.
                //
                // *** onFinished, NOT onChanged (code review, 2026-09-25). *** It used to run only when
                // a receipt MOVED during this launch, and a receipt that had already reached its last
                // state by an earlier launch never moves again - so a draft whose synced board
                // arrived one launch late was never retired. Retirement depends on a state, so it
                // runs after every check, whatever it found - it costs nothing when no receipt is
                // published.
                // ###########################################################################################
                _ = SubmissionStatusRefresh.RefreshQuietlyAsync(
                    new SubmissionClient().GetStatusAsync,
                    DateTimeOffset.UtcNow,
                    onFinished: changedCount => Dispatcher.UIThread.Post(() => _ = this.RetirePublishedDraftsAsync()));
            }
            catch (Exception ex)
            {
                Logger.Critical($"Main.StartAsync failed: {ex}");
            }
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
        // Re-runs the current board's load so "View boards as officially published" (session 2b)
        // takes effect immediately rather than on the next board switch. Delegates to the same
        // reload ApplyCatalogueVisibility uses - see that method's private counterpart for why one
        // shared reload exists rather than a second near-identical one.
        // ###########################################################################################
        public void ApplyViewOfficialPublishedOnly()
        {
            this.ReloadCurrentBoardForCatalogueVisibility();
        }

        // ###########################################################################################
        // Shows or hides the "Drafts" tab (session 2b, task 7) to match whether ANY system has a
        // local draft - unlike EnableWorklog/EnableNetworkConnectedOscilloscopeTab, this is not a
        // user preference; it is a plain fact about the Drafts/ folder's current contents. Called
        // from StartAsync (once DataManager.HardwareBoards exists) and from TabDrafts itself after
        // a discard, so the tab disappears the moment the last draft is gone rather than staying
        // visible-but-empty until the app restarts.
        // ###########################################################################################
        public void ApplyDraftsTabVisibility()
        {
            if (this.DraftsTabItem == null)
                return;

            this.TabDrafts.RefreshDrafts();

            // The Hardware/Board drop-downs' "Draft" chips follow the list just rebuilt - see
            // Main.DraftBadges.cs.
            this.ApplyDraftBadges();

            // ###########################################################################################
            // *** UNREAD REVIEWER FEEDBACK KEEPS THIS TAB VISIBLE EVEN WITH NO DRAFTS. ***
            //
            // Without this the badge would be unreachable in exactly the case it exists for. A
            // contributor submits a draft and then discards it locally - the ordinary thing to do
            // once it has been sent - which takes their draft count to zero and hides the tab. The
            // reviewer then asks for changes, and the only place that feedback is shown is behind a
            // button on a tab that is no longer on screen. They would never find out.
            //
            // The reverse is deliberately NOT true: no feedback and no drafts still hides the tab,
            // so somebody who has never contributed sees no trace of the feature.
            bool hasUnreadFeedback = SubmissionReceiptStore.UnreadCommentCount() > 0;
            bool hasAnyDrafts = this.TabDrafts.Drafts.Count > 0;

            bool shouldShow = hasAnyDrafts || hasUnreadFeedback;
            this.DraftsTabItem.IsVisible = shouldShow;

            if (!shouldShow)
            {
                this.MoveSelectionOffHiddenTab(this.DraftsTabItem);
            }
        }

        // ###########################################################################################
        // Deletes any local draft whose work is now in the published data (maintainer request,
        // 2026-09-23).
        //
        // *** THE DECISION IS DraftRetirement'S, NOT THIS METHOD'S. *** Everything about when a
        // draft may safely go - published state, equality with the synced board, never a
        // draft-only system - lives in that pure class so it is unit tested rather than trusted.
        // What happens here is only the deletion and the logging, which is the same division
        // TabDrafts' own Discard button keeps.
        //
        // *** THIS DELETES A FOLDER WITHOUT ASKING, so it is worth saying plainly why that is
        // acceptable here and nowhere else in the drafts flow. *** The manual Discard button
        // confirms first because it destroys work that exists in only one place. This path runs
        // only when the draft has been proved to contain NOTHING the published board does not
        // already carry - so the bytes it removes are a duplicate of data the sync will hand back
        // on demand. There is nothing to lose and therefore nothing to confirm; the maintainer
        // asked for it to be automatic precisely so a contributor does not have to remember.
        //
        // Failures are logged and swallowed. This runs unattended on the launch path, and a
        // locked folder must cost a stale row rather than an error dialog on somebody's way in.
        //
        // *** THE SEQUENCING IS PublishedDraftRetirer'S (code review, 2026-09-25). *** The check
        // runs OFF the UI thread (it parses two workbooks and reads every file per draft - it used
        // to freeze the window inside a Post), each draft is re-stamped just before it is deleted,
        // a draft whose table holds unsaved edits is kept, and the board cache is cleared for every
        // folder touched. If the board ON SCREEN was one of them it is reloaded, the same way the
        // manual Discard button refreshes - it used to keep showing rows and "drafted" tints read
        // from a folder that no longer existed.
        //
        // Called after the launch status check and after "My submissions" closes (its Refresh can
        // mark a submission published), and always ends by refreshing the Drafts tab and chips.
        // ###########################################################################################
        internal async Task RetirePublishedDraftsAsync()
        {
            bool refreshed = false;

            try
            {
                List<SubmissionReceipt> receipts = SubmissionReceiptStore.All.ToList();
                string dataRoot = DataManager.DataRoot;
                string draftsRoot = DraftManager.DraftsRoot;

                // Copied here, on the UI thread: the search runs on the pool. A receipt names its
                // system by id, and these are what that id is matched against - the PUBLISHED boards
                // CRT lists, not draft-only systems: a draft is only retired against a published
                // board the contributor can open instead (see DraftRetirement's header).
                List<string> excelDataFiles = DataManager.HardwareBoards
                    .Where(entry => !entry.IsDraftOnly)
                    .Select(entry => entry.ExcelDataFile)
                    .Where(file => !string.IsNullOrWhiteSpace(file))
                    .ToList();

                // Which drafts are draft-only systems: retiring one removes a whole entry from the
                // hardware and board lists, not just a draft.
                var draftOnly = new HashSet<string>(
                    DataManager.HardwareBoards.Where(entry => entry.IsDraftOnly).Select(entry => entry.ExcelDataFile),
                    StringComparer.OrdinalIgnoreCase);

                IReadOnlyList<RetirableDraft> candidates = await PublishedDraftRetirer.FindAsync(
                    receipts, dataRoot, draftsRoot, excelDataFiles);

                if (candidates.Count > 0)
                {
                    DraftRetirementOutcome outcome = PublishedDraftRetirer.Retire(
                        candidates,
                        // Each draft is named by its WORKBOOK path (RetirableDraft.ExcelDataFile),
                        // never by the receipt's system id - passing the id here is the mix-up that
                        // kept every published draft on screen until 2026-09-25.
                        isInUse: excelDataFile => this.TabDrafts.HasUnsavedTableEditsFor(excelDataFile),
                        discard: excelDataFile =>
                        {
                            // A table open on this draft with nothing unsaved just closes - it would
                            // otherwise sit on a folder that no longer exists.
                            this.TabDrafts.CloseTableIfOpenFor(excelDataFile);

                            return DraftManager.DiscardDraft(excelDataFile);
                        },
                        afterDiscard: DataManager.ClearBoardCache);

                    // A NEW system's draft retired (2026-09-25): its draft-only entry is gone, and the
                    // published board the master lists takes its place in the lists - the same
                    // refresh a manual discard does (TabDrafts.DiscardConfirmed).
                    if (outcome.Touched.Any(draftOnly.Contains))
                    {
                        DataManager.RefreshDraftOnlySystems();
                        this.RefreshHardwareAndBoardSelectionsAfterDraftChange();
                    }

                    HardwareBoardEntry? current = this.GetCurrentBoardEntry();

                    if (current is not null
                        && outcome.Touched.Contains(current.ExcelDataFile, StringComparer.OrdinalIgnoreCase))
                    {
                        // Also refreshes the Drafts tab - see ReloadCurrentBoardFromDisk.
                        this.ReloadCurrentBoardFromDisk(string.Empty);
                        refreshed = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not retire published drafts: [{ex.Message}]");
            }

            if (refreshed)
            {
                return;
            }

            // Guarded like the block above, and for the same reason: on the launch path this Task
            // is started and discarded, so anything escaping here would fault it unobserved. The
            // refresh runs whether or not retirement worked - it is also what shows the badge for a
            // reviewer's reply the status check just brought down.
            try
            {
                this.ApplyDraftsTabVisibility();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not refresh the Drafts tab after checking for retirable drafts: [{ex.Message}]");
            }
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
        // Whether a click should send keyboard focus back to the always-on component filter on the
        // left. That filter is meant to catch typing wherever the user happens to be, so every
        // pointer release pulls focus back to it - EXCEPT where a tab needs the keyboard itself.
        //
        // Extracted from the pointer-release handler (2026-09-24) so the rule can be tested; the
        // conditions are the handler's own, unchanged, plus the Drafts table below.
        //
        // *** THE DRAFTS TAB'S TABLE EDITOR OWNS THE KEYBOARD WHILE IT IS OPEN. *** A click on a
        // grid cell focuses the GRID, not a text box, so the "another TextBox holds focus" check
        // below does not cover it: focus went straight back to the filter, and typing into the
        // selected cell landed in the component filter instead (reported by the maintainer).
        // ###########################################################################################
        internal bool ShouldReturnFocusToComponentSearch()
        {
            // Do not steal focus while the schematics label editor is active
            if (this.TabSchematicsControl.IsLabelEditorActive)
            {
                return false;
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
                return false;
            }

            if (ReferenceEquals(selectedTab, this.DraftsTabItem) && this.TabDrafts.IsTableOpen)
            {
                return false;
            }

            // Avoid stealing focus if another TextBox currently holds it naturally
            var focusedElement = TopLevel.GetTopLevel(this)?.FocusManager?.GetFocusedElement();
            if (focusedElement is global::Avalonia.Controls.TextBox && focusedElement != this.ComponentSearchTextBox)
            {
                return false;
            }

            return this.ComponentSearchTextBox != null && !this.ComponentSearchTextBox.IsFocused;
        }

        // ###########################################################################################
        // Stops any pending debounce timer and does a final synchronous save on close.
        // ###########################################################################################
        private void OnWindowClosing(object? sender, WindowClosingEventArgs e)
        {
            // ###########################################################################################
            // *** UNSAVED EDITS IN THE DRAFTS TAB'S TABLE STOP THE EXIT UNTIL ANSWERED (2026-09-24). ***
            //
            // They live only in memory until "Save changes", so quitting would lose them without a
            // word. The close is cancelled, the Drafts tab is brought forward (so a refused save's
            // reason is on screen, not behind another tab), and the prompt is asked from there; a
            // Save that goes through or a Discard then closes for real.
            //
            // It sits FIRST, before the e.Cancel guard below, because it is the thing that cancels -
            // and the shutdown flag after that guard must not be set for a close that did not
            // happen (see the comment on it).
            // ###########################################################################################
            if (!this._unsavedTableEditsSettledForExit && this.TabDrafts.HasUnsavedTableEdits)
            {
                e.Cancel = true;

                Dispatcher.UIThread.Post(async () =>
                {
                    if (this.DraftsTabItem != null && this.MainTabControl != null)
                    {
                        this.MainTabControl.SelectedItem = this.DraftsTabItem;
                    }

                    if (await this.TabDrafts.ConfirmLeavingTableAsync(this))
                    {
                        this._unsavedTableEditsSettledForExit = true;
                        this.Close();
                    }
                });

                return;
            }

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
    }
}
