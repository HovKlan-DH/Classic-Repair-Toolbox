using Avalonia;
using Avalonia.Controls.Templates;
using Avalonia.Data;
using Avalonia.Media;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.OnlineHandling;
using Handlers.DataHandling;

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

        private BoardData? _currentBoardData;
        private bool _suppressComponentHighlightUpdate;
        private ComponentInfoWindow? _singleComponentInfoWindow;
        private readonly Dictionary<string, ComponentInfoWindow> _componentInfoWindowsByKey = new(StringComparer.OrdinalIgnoreCase);
        internal bool isHoveringComponent = false;

        // Blink selected highlights
        private DispatcherTimer? _blinkSelectedTimer;
        private bool _blinkSelectedPhaseVisible = true;
        private bool _blinkSelectedEnabled;

        // Region toggle: local override, does not affect the global setting
        private string _localRegion = UserSettings.Region;
        public string LocalRegion => this._localRegion;
        private bool _suppressRegionToggle;

        // Cascading offset for multiple popups
        private int _popupCascadeOffset = 0;

        // Fullscreen
        private SchematicsFullscreenWindow? _schematicsFullscreenWindow;

        // Exposes the Oscilloscope tab control so other UI windows can route scope actions through it.
        public TabOscilloscope TabOscilloscopeControl => this.TabOscilloscope;

        public Main()
        {
            InitializeComponent();

            this.TabSchematicsControl.Initialize(this);
            this.TabOverview.Initialize(this);
            this.TabContribute.Initialize(this);

            // Restore left panel width from settings
            this.RootGrid.ColumnDefinitions[0].Width = new GridLength(UserSettings.LeftPanelWidth);
            this.RootGrid.ColumnDefinitions[2].Width = new GridLength(1, GridUnitType.Star);

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
            this.PopulateHardwareDropDown();

            var versionString = AppConfig.AppVersionString;
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

                        // Do not steal focus if we are on tabs that utilize text inputs
                        var selectedTab = this.MainTabControl?.SelectedItem as TabItem;
                        string? tabHeader = selectedTab?.Header?.ToString();

                        if (tabHeader == "Feedback" || tabHeader == "Configuration")
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

            if (UserSettings.CheckVersionOnLaunch)
            {
                this.CheckForAppUpdate();
            }
            else if (DataManager.DataUpdateRequiresAppUpdate)
            {
                // Notify if they aren't checking for app updates but missing critical data updates
                this.UpdateBannerText.Text = "Newer main Excel data file is available, but requires a newer application version. No more data updates will be given for this version";
                this.UpdateBannerInstallButton.IsVisible = false;
                this.UpdateBannerViewNotesButton.IsVisible = false;
                this.UpdateBanner.IsVisible = true;
            }

            this.StartBackgroundSyncAsync();
        }


        // ###########################################################################################
        // Checks for an available update on startup and shows the banner if one is found.
        // ###########################################################################################
        private async void CheckForAppUpdate()
        {
            bool? available = await UpdateService.CheckForUpdateAsync();

            if (available == true)
            {
                this.UpdateBannerText.Text = $"Version {UpdateService.PendingVersion} is available";
                this.UpdateBanner.IsVisible = true;
            }
            else if (DataManager.DataUpdateRequiresAppUpdate)
            {
                // App Velopack doesn't see an update, but manifest demands one.
                this.UpdateBannerText.Text = "Newer main Excel data file is available, but requires a newer application version. No more data updates will be given for this version";
                this.UpdateBannerInstallButton.IsVisible = false;
                this.UpdateBannerViewNotesButton.IsVisible = false;
                this.UpdateBanner.IsVisible = true;
            }
        }

        // ###########################################################################################
        // Shows the sync banner during background sync, then hides it automatically if nothing
        // changed, or keeps it visible with an update summary and a refresh button.
        // ###########################################################################################
        private async void StartBackgroundSyncAsync()
        {
            if (!DataManager.HasPendingSync)
                return;

            this.SyncBannerText.Text = "Checking data from online source - please wait...";
            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBanner.IsVisible = true;

            int changed = await DataManager.SyncRemainingAsync(status =>
                Dispatcher.UIThread.Post(() => this.SyncBannerText.Text = status));

            if (changed > 0)
            {
                this.SyncBannerText.Text = changed == 1
                    ? "1 file updated in the background - please refresh board"
                    : $"{changed} files updated in the background - please refresh board";

                this.SyncBannerRefreshButton.IsVisible = true;
            }
            else
            {
                this.SyncBanner.IsVisible = false;
            }
        }

        // ###########################################################################################
        // Manually reloads the current board configuration.
        // ###########################################################################################
        private void OnRefreshBoardClick(object? sender, RoutedEventArgs e)
        {
            this.SyncBanner.IsVisible = false;
            this.OnBoardSelectionChanged(null, null!);
        }

        // ###########################################################################################
        // Dismisses the sync banner.
        // ###########################################################################################
        private void OnSyncBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the sync banner when clicking anywhere on it.
        // ###########################################################################################
        private void OnSyncBannerPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the update banner without cancelling the update.
        // ###########################################################################################
        private void OnUpdateBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.UpdateBanner.IsVisible = false;
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
            this.OpenUrl(url);
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
            var hardwareNames = DataManager.HardwareBoards
                .Select(e => e.HardwareName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

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
        // Filters the board drop-down to only show boards belonging to the selected hardware.
        // ###########################################################################################
        private void OnHardwareSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            var selectedHardware = this.HardwareComboBox.SelectedItem as string;

            var boards = DataManager.HardwareBoards
                .Where(entry => string.Equals(entry.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.BoardName)
                .Where(b => !string.IsNullOrWhiteSpace(b))
                .ToList();

            this.BoardComboBox.ItemsSource = boards;

            if (string.IsNullOrWhiteSpace(selectedHardware) || boards.Count == 0)
            {
                this.BoardComboBox.SelectedIndex = -1;
                return;
            }

            UserSettings.SetLastHardware(selectedHardware);

            var lastBoard = UserSettings.GetLastBoardForHardware(selectedHardware);
            var savedIndex = boards.FindIndex(b =>
                string.Equals(b, lastBoard, StringComparison.OrdinalIgnoreCase));

            this.BoardComboBox.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
        }

        // ###########################################################################################
        // Handles board selection changes - loads board data and builds the thumbnail gallery.
        // ###########################################################################################
        private async void OnBoardSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this._suppressCategoryFilterSave = true;

            foreach (var thumb in this.TabSchematicsControl.currentThumbnails)
            {
                if (!ReferenceEquals(thumb.ImageSource, thumb.BaseThumbnail))
                    (thumb.ImageSource as IDisposable)?.Dispose();
                (thumb.BaseThumbnail as IDisposable)?.Dispose();
            }
            this.TabSchematicsControl.currentThumbnails.Clear();
            this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList")!.ItemsSource = null;
            this.CategoryFilterListBox.ItemsSource = null;
            this.ComponentFilterListBox.ItemsSource = null;

            this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);
            this.TabSchematicsControl.schematicByName = new(StringComparer.OrdinalIgnoreCase);
            this.TabSchematicsControl.highlightRectsBySchematicAndLabel = new(StringComparer.OrdinalIgnoreCase);
            this._currentBoardData = null;
            this.UpdateRegionButtonsState();

            this.PopulateBoardInfoSection(null, null);

            this.TabSchematicsControl.ResetSchematicsViewer();

            var selectedHardware = this.HardwareComboBox.SelectedItem as string;
            var selectedBoard = this.BoardComboBox.SelectedItem as string;

            if (string.IsNullOrEmpty(selectedHardware) || string.IsNullOrEmpty(selectedBoard))
                return;

            UserSettings.SetLastHardware(selectedHardware);
            UserSettings.SetLastBoardForHardware(selectedHardware, selectedBoard);

            var entry = DataManager.HardwareBoards.FirstOrDefault(ent =>
                string.Equals(ent.HardwareName, selectedHardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(ent.BoardName, selectedBoard, StringComparison.OrdinalIgnoreCase));

            if (entry == null || string.IsNullOrWhiteSpace(entry.ExcelDataFile))
                return;

            var boardData = await DataManager.LoadBoardDataAsync(entry);
            if (boardData == null)
                return;

            this._currentBoardData = boardData;
            this.UpdateRegionButtonsState();

            this.PopulateBoardInfoSection(boardData.RevisionDate, boardData.Credits);

            // Populate category filter in insertion order
            var categories = BuildDistinctCategories(boardData);
            var boardKey = this.GetCurrentBoardKey();

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
                        this.CategoryFilterListBox.Selection.Select(i);
                }
            }
            this._suppressCategoryFilterSave = false;

            // Populate component filter for this board
            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = BuildComponentItems(boardData, UserSettings.Region, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                try { this.ComponentFilterListBox.SelectAll(); } catch { }
            }
            this._suppressComponentHighlightUpdate = false;

            this.TabSchematicsControl.highlightRectsBySchematicAndLabel = await Task.Run(() => TabSchematics.BuildHighlightRects(boardData, UserSettings.Region));
            this.TabSchematicsControl.schematicByName = boardData.Schematics
                .Where(s => !string.IsNullOrWhiteSpace(s.SchematicName))
                .ToDictionary(s => s.SchematicName, s => s, StringComparer.OrdinalIgnoreCase);
            this.TabSchematicsControl.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

            var loaded = await Task.Run(() =>
            {
                var result = new List<(string Name, string FullPath, Bitmap? FullBitmap)>();

                foreach (var schematic in boardData.Schematics)
                {
                    if (string.IsNullOrWhiteSpace(schematic.SchematicImageFile))
                        continue;

                    var fullPath = Path.Combine(DataManager.DataRoot,
                        schematic.SchematicImageFile.Replace('/', Path.DirectorySeparatorChar));

                    Bitmap? bitmap = null;
                    if (File.Exists(fullPath))
                    {
                        try { bitmap = new Bitmap(fullPath); }
                        catch (Exception ex) { Logger.Warning($"Could not load schematic image [{fullPath}] - [{ex.Message}]"); }
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

            this.TabSchematicsControl.LoadSortedThumbnails(boardKey, thumbnails);

            if (this.TabSchematicsControl.currentThumbnails.Count > 0)
            {
                var savedSchematic = UserSettings.GetLastSchematicForBoard(boardKey);
                var orderedThumbnails = this.TabSchematicsControl.currentThumbnails.ToList();
                var savedIndex = string.IsNullOrEmpty(savedSchematic) ? -1 : orderedThumbnails.FindIndex(t =>
                    string.Equals(t.Name, savedSchematic, StringComparison.OrdinalIgnoreCase));

                this.TabSchematicsControl.FindControl<ListBox>("SchematicsThumbnailList")!.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
            }

            var ratio = UserSettings.GetSchematicsSplitterRatio(boardKey);
            var innerGrid = this.TabSchematicsControl.FindControl<Grid>("SchematicsInnerGrid");
            if (innerGrid != null)
            {
                innerGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
                innerGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);
            }

            // Populate the Resources tab
            var localFiles = boardData.BoardLocalFiles.Select(f => new ResourceItem(
                f.Category,
                f.Name,
                string.IsNullOrWhiteSpace(f.File) ? string.Empty : Path.Combine(DataManager.DataRoot, f.File.Replace('/', Path.DirectorySeparatorChar))
            ));

            var webLinks = boardData.BoardLinks.Select(l => new ResourceItem(
                l.Category,
                l.Name,
                l.Url
            ));

            this.TabResources.LoadData(localFiles, webLinks);
            this.TabOverview.LoadData(boardData);
            this.TabContribute.LoadData(boardData, this._localRegion);

            // Sync any existing search filter right away if applied
            this.TabOverview.ApplyFilter(this.ComponentSearchTextBox?.Text ?? string.Empty);
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
                var componentItems = BuildComponentItems(this._currentBoardData, this._localRegion, categoryFilter, searchTerm);

                this._suppressComponentHighlightUpdate = true;
                this.ComponentFilterListBox.ItemsSource = componentItems;

                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
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
            }
        }

        // ###########################################################################################
        // Returns a composite key uniquely identifying the current hardware and board selection.
        // ###########################################################################################
        internal string GetCurrentBoardKey()
        {
            var hw = this.HardwareComboBox.SelectedItem as string;
            var board = this.BoardComboBox.SelectedItem as string;
            if (string.IsNullOrEmpty(hw) || string.IsNullOrEmpty(board))
            {
                return string.Empty;
            }
            return $"{hw}|{board}";
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

            if (UserSettings.HasWindowPlacement && this.WindowState == Avalonia.Controls.WindowState.Normal)
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
                _ = Task.Run(DataValidator.ValidateAllDataAsync);
            }

            Dispatcher.UIThread.Post(() => this._windowPlacementReady = true, DispatcherPriority.Background);

            Dispatcher.UIThread.Post(() =>
            {
                this.ComponentSearchTextBox?.Focus();
            }, DispatcherPriority.Background);
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
            if (Application.Current?.ApplicationLifetime is Avalonia.Controls.ApplicationLifetimes.IClassicDesktopStyleApplicationLifetime desktop)
            {
                Dispatcher.UIThread.Post(() =>
                {
                    desktop.Shutdown();
                });
            }
        }

        // ###########################################################################################
        // Builds a distinct list of component categories in the order they first appear.
        // ###########################################################################################
        private static List<string> BuildDistinctCategories(BoardData boardData)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var categories = new List<string>();

            foreach (var component in boardData.Components)
            {
                if (!string.IsNullOrWhiteSpace(component.Category) && seen.Add(component.Category))
                    categories.Add(component.Category);
            }

            return categories;
        }

        // ###########################################################################################
        // Lightweight view model for a component list item.
        // ###########################################################################################
        internal sealed class ComponentListItem
        {
            public string DisplayText { get; init; } = string.Empty;
            public string BoardLabel { get; init; } = string.Empty;
            public string SelectionKey { get; init; } = string.Empty;
            public override string ToString() => this.DisplayText;
        }

        // ###########################################################################################
        // Builds component list items filtered by the given region and search string.
        // ###########################################################################################
        private static List<ComponentListItem> BuildComponentItems(BoardData boardData, string region, HashSet<string>? categoryFilter = null, string searchTerm = "")
        {
            var items = new List<ComponentListItem>();

            var searchTerms = string.IsNullOrWhiteSpace(searchTerm)
                ? Array.Empty<string>()
                : searchTerm.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var component in boardData.Components)
            {
                var componentRegion = component.Region?.Trim() ?? string.Empty;

                if (!string.IsNullOrEmpty(componentRegion) &&
                    !string.Equals(componentRegion, region, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (categoryFilter != null && !categoryFilter.Contains(component.Category ?? string.Empty))
                    continue;

                var parts = new List<string>(3);
                if (!string.IsNullOrWhiteSpace(component.BoardLabel))
                    parts.Add(component.BoardLabel.Trim());
                if (!string.IsNullOrWhiteSpace(component.FriendlyName))
                    parts.Add(component.FriendlyName.Trim());
                if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
                    parts.Add(component.TechnicalNameOrValue.Trim());

                if (parts.Count == 0)
                    continue;

                string displayString = string.Join(" | ", parts);

                if (searchTerms.Length > 0)
                {
                    bool matches = true;
                    foreach (var term in searchTerms)
                    {
                        if (displayString.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            matches = false;
                            break;
                        }
                    }

                    if (!matches)
                        continue;
                }

                items.Add(new ComponentListItem
                {
                    BoardLabel = component.BoardLabel?.Trim() ?? string.Empty,
                    DisplayText = displayString,
                    SelectionKey = string.Join("\u001F",
                        component.BoardLabel?.Trim() ?? string.Empty,
                        component.FriendlyName?.Trim() ?? string.Empty,
                        component.TechnicalNameOrValue?.Trim() ?? string.Empty,
                        component.Region?.Trim() ?? string.Empty)
                });
            }

            return items;
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
        // Opens the configured URL in the system default browser.
        // ###########################################################################################
        private void OpenUrl(string url)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open URL - [{url}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Clears all currently selected items in the Component filter list box.
        // ###########################################################################################
        private void OnClearComponentsClick(object? sender, RoutedEventArgs e)
        {
            this.ComponentFilterListBox.SelectedItems?.Clear();
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
            bool hasExplicitRegionComponents = HasExplicitRegionComponents(this._currentBoardData);

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
            bool hasExplicitRegionComponents = HasExplicitRegionComponents(boardData);
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

            if (!this._singleComponentInfoWindow.IsVisible)
            {
                this.PositionPopupOnSameScreen(this._singleComponentInfoWindow);
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

            bool hasSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;

            if (this._blinkSelectedEnabled && hasSelection)
            {
                this._blinkSelectedPhaseVisible = false;
                this.TabSchematicsControl.ApplyHighlightVisuals(true, this.GetCurrentBlinkFactor(true));
                this.UpdateBlinkTimer(true);
                return;
            }

            this._blinkSelectedPhaseVisible = true;
            this.UpdateBlinkTimer(hasSelection);
            this.TabSchematicsControl.ApplyHighlightVisuals(hasSelection, this.GetCurrentBlinkFactor(hasSelection));
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
            bool hasSelection = this.TabSchematicsControl.highlightIndexBySchematic.Count > 0;
            if (!hasSelection)
            {
                this.UpdateBlinkTimer(false);
                this.TabSchematicsControl.ApplyHighlightVisuals(false, 1.0);
                return;
            }

            this._blinkSelectedPhaseVisible = !this._blinkSelectedPhaseVisible;
            this.TabSchematicsControl.ApplyHighlightVisuals(true, this.GetCurrentBlinkFactor(true));
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
        // Closes single popup when pressing Escape while the main window is focused.
        // F11 opens the schematics fullscreen window.
        // ###########################################################################################
        private void OnMainKeyDownCloseSinglePopup(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.F11)
            {
                this.OpenSchematicsFullscreenWindow();
                e.Handled = true;
                return;
            }

            if (e.Key != Key.Escape)
                return;

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

            this.TabSchematicsControl.highlightRectsBySchematicAndLabel = await Task.Run(() =>
                TabSchematics.BuildHighlightRects(this._currentBoardData, this._localRegion));

            var previouslySelectedKeys = new HashSet<string>(
                this.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>()
                    .Select(i => i.SelectionKey) ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;
            var componentItems = BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
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
        }

        // ###########################################################################################
        // Refreshes the component filter list based on search text, and highlights the result.
        // ###########################################################################################
        public void OnComponentSearchTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            if (this._currentBoardData == null || this._suppressCategoryFilterSave)
                return;

            var activeCategories = new HashSet<string>(
                this.CategoryFilterListBox.SelectedItems?.Cast<string>() ?? Enumerable.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);

            var searchTerm = this.ComponentSearchTextBox?.Text ?? string.Empty;

            var componentItems = BuildComponentItems(this._currentBoardData, this._localRegion, activeCategories, searchTerm);

            this._suppressComponentHighlightUpdate = true;
            this.ComponentFilterListBox.ItemsSource = componentItems;

            var highlightLabels = new List<string>();

            if (!string.IsNullOrWhiteSpace(searchTerm))
            {
                try { this.ComponentFilterListBox.SelectAll(); } catch { }

                highlightLabels = componentItems
                    .Select(item => item.BoardLabel)
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToList();
            }

            this._suppressComponentHighlightUpdate = false;

            this.TabSchematicsControl.UpdateHighlightsForComponents(highlightLabels);

            // Forward the search term to filter the Overview tab's list
            this.TabOverview.ApplyFilter(searchTerm);
        }

        // ###########################################################################################
        // Returns true when the current board has at least one component explicitly tagged as PAL or NTSC.
        // ###########################################################################################
        internal bool CurrentBoardHasExplicitRegionComponents()
        {
            return HasExplicitRegionComponents(this._currentBoardData);
        }

        // ###########################################################################################
        // Returns true when the provided board has at least one component explicitly tagged as PAL or NTSC.
        // ###########################################################################################
        private static bool HasExplicitRegionComponents(BoardData? boardData)
        {
            if (boardData == null)
                return false;

            return boardData.Components.Any(component =>
                string.Equals(component.Region?.Trim(), "PAL", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(component.Region?.Trim(), "NTSC", StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Creates placeholder content for the Schematics tab while fullscreen mode is active.
        // Keeps the thumbnail list available so another schematic can be selected without closing
        // the fullscreen window first.
        // ###########################################################################################
        private Control CreateSchematicsFullscreenPlaceholder()
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
        // Opens the existing schematics viewer in a separate maximized window.
        // ###########################################################################################
        private void OpenSchematicsFullscreenWindow()
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

            this._schematicsFullscreenWindow.Closed += (_, _) =>
            {
                this._schematicsFullscreenWindow = null;
            };

            this.PositionFullscreenWindowOnSameScreen(this._schematicsFullscreenWindow);
            this._schematicsFullscreenWindow.WindowState = Avalonia.Controls.WindowState.Maximized;
            this._schematicsFullscreenWindow.Show();

            this.TabSchematicsControl.RefreshAfterHostChanged();
            this._schematicsFullscreenWindow.Focus();
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

            var hardwareName = this.HardwareComboBox.SelectedItem as string ?? string.Empty;
            var boardName = this.BoardComboBox.SelectedItem as string ?? string.Empty;

            var window = new ComponentContributionWindow();
            window.LoadComponent(
                this._currentBoardData,
                DataManager.DataRoot,
                hardwareName,
                boardName,
                this._localRegion,
                boardLabel);

            this.PositionFullscreenWindowOnSameScreen(window);
            window.WindowState = Avalonia.Controls.WindowState.Maximized;
            window.Show(this);
            window.Focus();
        }

        // ###########################################################################################
    }
}