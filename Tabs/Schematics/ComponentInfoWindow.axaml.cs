using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.IcTesting;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;


namespace CRT
{
    // ###########################################################################################
    // View model for a single component image entry shown in the thumbnail gallery.
    // ###########################################################################################
    public sealed class ComponentImageItem
    {
        public Bitmap? ImageSource { get; set; }
        public string Label { get; set; } = string.Empty;
        public string Pin { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ExpectedOscilloscopeReading { get; set; } = string.Empty;
        public string TimeDiv { get; set; } = string.Empty;
        public string VoltsDiv { get; set; } = string.Empty;
        public string TriggerLevelVolts { get; set; } = string.Empty;
        public string Note { get; set; } = string.Empty;
        public ComponentImageEntry? SourceEntry { get; set; }
        public bool LabelVisible => !string.IsNullOrEmpty(this.Label);
    }

    // ###########################################################################################
    // View model for a single local file entry shown in the local files list.
    // ###########################################################################################
    public sealed class ComponentLocalFileItem
    {
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // View model for a single link entry shown in the links list.
    // ###########################################################################################
    public sealed class ComponentLinkItem
    {
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    // ###########################################################################################
    // Popup window that displays detailed component information in a split-panel layout.
    // ###########################################################################################
    public partial class ComponentInfoWindow : Window
    {
        private readonly List<Bitmap> _loadedBitmaps = new List<Bitmap>();
        private CancellationTokenSource? _loadCts;
        private string _pinBuffer = string.Empty;
        private CancellationTokenSource? _pinBufferCts;
        private CancellationTokenSource? _pinFlashCts;
        private string _localRegion = "PAL";
        private string _displayTextFallback = string.Empty;
        private List<ComponentEntry> _allComponentEntries = new List<ComponentEntry>();
        private List<ComponentImageEntry> _allComponentImages = new List<ComponentImageEntry>();
        private string _boardLabel = string.Empty;
        private string _dataRoot = string.Empty;
        private bool _suppressThumbnailSelection = false;
        private bool _suppressRegionToggle = false;
        private double _normalWidth = 680.0;
        private double _normalHeight = 420.0;
        private int _normalX;
        private int _normalY;
        private bool _hasExplicitRegionComponents = false;

        // Image matrix for zoom and pan capabilities
        private Matrix _imageMatrix = Matrix.Identity;
        private bool _isPanningImage = false;
        private Point _panStartPoint;
        private Matrix _panStartMatrix;

        private bool _hasSeenOscilloscopeSessionTitleState;
        private bool _hasActiveOscilloscopeSessionTitleState;
        private DateTime _lastOscilloscopeKeyboardCommandUtc = DateTime.MinValue;
        private int _oscilloscopeKeyboardCommandInFlight;
        private Bitmap? thisTemporaryCapturedScopeBitmap;


        // ###########################################################################################
        // When true, the window closes itself whenever it loses focus to another window.
        // ###########################################################################################
        public bool CloseOnDeactivate { get; set; }

        public ComponentInfoWindow()
        {
            this.InitializeComponent();

            // Aggressively steal focus from the Main window's TextBox when interacting with this window
            this.PointerPressed += (_, _) => this.Focus();
            this.PointerEntered += (_, _) => this.Focus();

            this.Opened += (_, _) => this.RestoreKeyboardFocusToPopup();
            this.Activated += (_, _) => this.RestoreKeyboardFocusToPopup();

            // Seed the normal-size tracker and restore saved window size, splitters and state
            this._normalWidth = UserSettings.HasComponentInfoWindowLayout
                ? UserSettings.ComponentInfoWindowWidth
                : 680.0;

            this._normalHeight = UserSettings.HasComponentInfoWindowLayout
                ? UserSettings.ComponentInfoWindowHeight
                : 420.0;

            this._normalX = UserSettings.HasComponentInfoWindowLayout ? UserSettings.ComponentInfoWindowX : 0;
            this._normalY = UserSettings.HasComponentInfoWindowLayout ? UserSettings.ComponentInfoWindowY : 0;

            if (UserSettings.HasComponentInfoWindowLayout)
            {
                this.Width = UserSettings.ComponentInfoWindowWidth;
                this.Height = UserSettings.ComponentInfoWindowHeight;

                double ratio = Math.Clamp(UserSettings.ComponentInfoWindowLeftColumnRatio, 0.1, 0.9);
                this.RootGrid.ColumnDefinitions[0].Width = new GridLength(ratio, GridUnitType.Star);
                this.RootGrid.ColumnDefinitions[2].Width = new GridLength(1.0 - ratio, GridUnitType.Star);

                double thumbHeight = Math.Max(40.0, UserSettings.ComponentInfoWindowThumbnailRowHeight);
                this.LeftPanelGrid.RowDefinitions[2].Height = new GridLength(thumbHeight, GridUnitType.Pixel);

                if (string.Equals(UserSettings.ComponentInfoWindowState, "Maximized", StringComparison.OrdinalIgnoreCase))
                    this.WindowState = WindowState.Maximized;
            }

            // Restore switch states from persisted settings
            this.MousewheelZoomCheckBox.IsChecked =
                string.Equals(UserSettings.ComponentInfoScrollAction, "Image zoom", StringComparison.OrdinalIgnoreCase);

            this.NumpadOscilloscopeSwitch.IsChecked =
                string.Equals(UserSettings.ComponentInfoKeyboardHandling, "Control oscilloscope", StringComparison.OrdinalIgnoreCase);

            this.SyncOscilloscopeCheckBox.IsChecked = UserSettings.ComponentInfoOscilloscopeSyncEnabled;

            this.UpdateOscilloscopeControlsAvailability();

            // Replace Checked/Unchecked with IsCheckedChanged
            this.MousewheelZoomCheckBox.IsCheckedChanged += this.OnMousewheelZoomSwitchChanged;
            this.NumpadOscilloscopeSwitch.IsCheckedChanged += this.OnNumpadOscilloscopeSwitchChanged;
            this.SyncOscilloscopeCheckBox.IsCheckedChanged += this.OnSyncOscilloscopeCheckBoxChanged;

            this.ThumbnailList.SelectionChanged += this.OnThumbnailSelectionChanged;

            this.IcTestPanel.CloseRequested += this.OnIcTestPanelCloseRequested;

            // Map the interactions to the expanded top-panel boundaries area instead of the local box
            this.MainImageClickArea.PointerPressed += this.OnMainImageClickAreaPointerPressed;
            this.MainImageClickArea.PointerMoved += this.OnMainImageClickAreaPointerMoved;
            this.MainImageClickArea.PointerReleased += this.OnMainImageClickAreaPointerReleased;

            // Tunnel phase: intercepts key events before any child control (e.g. TextBox) sees them,
            // so arrow key navigation always works regardless of which control has focus.
            this.AddHandler(
                KeyDownEvent,
                this.OnWindowKeyDown,
                RoutingStrategies.Tunnel);

            // Tunnel phase: intercepts scroll wheel events on the left panel so scrolling
            // navigates thumbnails while allowing the right panel's ScrollViewer to work normally.
            this.LeftPanelGrid.AddHandler(
                PointerWheelChangedEvent,
                this.OnLeftPanelPointerWheelChanged,
                RoutingStrategies.Tunnel);

            // Keep _normalWidth/_normalHeight up to date so they always reflect the last
            // non-maximized dimensions regardless of how the window is closed.
            this.SizeChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this._normalWidth = this.Width;
                    this._normalHeight = this.Height;
                }
            };

            // Keep _normalX/_normalY up to date so a reopened single-instance popup can restore
            // the last on-screen position instead of re-cascading from the main window.
            this.PositionChanged += (_, _) =>
            {
                if (this.WindowState == WindowState.Normal)
                {
                    this._normalX = this.Position.X;
                    this._normalY = this.Position.Y;
                }
            };

            this.Deactivated += (_, _) =>
            {
                if (!this.CloseOnDeactivate)
                    return;

                // Immediately abort the close if we are hovering a component (re-use the window)
                if (this.Owner is Main mainOwner && mainOwner.isHoveringComponent)
                    return;

                this.Close();
            };

            this.Closing += (_, _) =>
            {
                string state = this.WindowState == WindowState.Maximized ? "Maximized" : "Normal";

                // Always use _normalWidth/_normalHeight so maximized dimensions never overwrite
                // the restored size that will be used when the window opens in Normal state.
                double leftWidth = this.LeftPanelGrid.Bounds.Width;
                double splitterThickness = 4.0;
                double rightWidth = this.RootGrid.Bounds.Width - leftWidth - splitterThickness;
                double leftRatio = (leftWidth + rightWidth) > 0.0
                    ? leftWidth / (leftWidth + rightWidth)
                    : 0.5;

                double thumbHeight = this.ThumbnailList.Bounds.Height;
                if (thumbHeight <= 0.0)
                    thumbHeight = UserSettings.ComponentInfoWindowThumbnailRowHeight;

                UserSettings.SaveComponentInfoWindowLayout(state, this._normalWidth, this._normalHeight, leftRatio, thumbHeight, this._normalX, this._normalY);
            };

            this.Closed += (_, _) =>
            {
                this.ClearTemporaryCapturedOscilloscopeImage();
                this._loadCts?.Cancel();
                this._pinBufferCts?.Cancel();
                this._pinFlashCts?.Cancel();
                foreach (var bmp in this._loadedBitmaps)
                    bmp.Dispose();
                this._loadedBitmaps.Clear();
            };
        }

        // ###########################################################################################
        // Restores a keyboard focus target inside the popup whenever it opens or becomes active.
        // This keeps Escape and other keyboard shortcuts working after clicking the native window frame.
        // ###########################################################################################
        private void RestoreKeyboardFocusToPopup()
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (!this.IsVisible)
                    return;

                this.Focus();
                this.RootGrid.Focus();
            }, DispatcherPriority.Input);
        }

        // ###########################################################################################
        // Intercepts key events at the tunnel phase so Escape, Left, Right and Enter always work
        // regardless of which child control currently has focus.
        // ###########################################################################################
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Left)
            {
                this.NavigateThumbnails(-1);
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Right)
            {
                this.NavigateThumbnails(1);
                e.Handled = true;
                return;
            }

            if (this.NumpadOscilloscopeSwitch.IsEnabled &&
                this.NumpadOscilloscopeSwitch.IsChecked == true &&
                this.TryHandleOscilloscopeKeyboardCommand(e.Key))
            {
                e.Handled = true;
                return;
            }

            if (e.Key == Key.Space || e.Key == Key.Enter)
            {
                this.ThumbnailList.SelectedIndex = 0;
                this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
                e.Handled = true;
                return;
            }

            // Digit keys: top-row digits always control pin selection.
            // Numpad digits do the same only when numpad-to-oscilloscope control is disabled.
            int digitValue = -1;
            if (e.Key >= Key.D0 && e.Key <= Key.D9)
                digitValue = (int)e.Key - (int)Key.D0;
            else if (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)
            {
                if (this.NumpadOscilloscopeSwitch.IsEnabled &&
                    this.NumpadOscilloscopeSwitch.IsChecked == true)
                    return;

                digitValue = (int)e.Key - (int)Key.NumPad0;
            }

            if (digitValue >= 0)
            {
                this.HandlePinDigit((char)('0' + digitValue));
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Routes popup keyboard shortcuts to the oscilloscope tab when oscilloscope keyboard control
        // is enabled for the component info window.
        // ###########################################################################################
        private bool TryHandleOscilloscopeKeyboardCommand(Key key)
        {
            switch (key)
            {
                case Key.Add:
                    return this.TryQueueOscilloscopeTimeDivStep(-1);

                case Key.Subtract:
                    return this.TryQueueOscilloscopeTimeDivStep(1);

                case Key.Up:
                    return this.TryQueueOscilloscopeTriggerLevelStep(1);

                case Key.Down:
                    return this.TryQueueOscilloscopeTriggerLevelStep(-1);

                case Key.NumPad1:
                    return this.TryQueueOscilloscopeVoltsDivSet(1.0);

                case Key.NumPad2:
                    return this.TryQueueOscilloscopeVoltsDivSet(2.0);

                case Key.Decimal:
                    return this.TryQueueOscilloscopeKeyboardCommand(
                        () => this.RunOscilloscopePaletteAsync(ScopeCommandPalette.ClearStatistics));

                case Key.Enter:
                    return this.TryQueueOscilloscopeKeyboardCommand(
                        this.CaptureAndDisplayOscilloscopeImageAsync);

                case Key.Multiply:
                    return this.TryQueueOscilloscopeKeyboardCommand(
                        () => this.RunOscilloscopePaletteAsync(ScopeCommandPalette.Single));

                case Key.Divide:
                    return this.TryQueueOscilloscopeKeyboardCommand(
                        () => this.RunOscilloscopePaletteAsync(ScopeCommandPalette.Run));

                default:
                    return false;
            }
        }

        // ###########################################################################################
        // Applies the oscilloscope Debounce-Time interval and a single-flight guard so held keys
        // do not flood the oscilloscope with repeated keyboard-driven commands.
        // ###########################################################################################
        private bool TryQueueOscilloscopeKeyboardCommand(Func<Task> commandAsync)
        {
            DateTime nowUtc = DateTime.UtcNow;
            TimeSpan minimumInterval = this.GetOscilloscopeKeyboardCommandMinimumInterval();

            if (nowUtc - this._lastOscilloscopeKeyboardCommandUtc < minimumInterval)
            {
                return true;
            }

            if (Interlocked.CompareExchange(ref this._oscilloscopeKeyboardCommandInFlight, 1, 0) != 0)
            {
                return true;
            }

            this._lastOscilloscopeKeyboardCommandUtc = nowUtc;
            _ = this.ExecuteOscilloscopeKeyboardCommandAsync(commandAsync);
            return true;
        }

        // ###########################################################################################
        // Executes one throttled oscilloscope keyboard command and releases the single-flight gate
        // afterward so the next accepted keypress can run.
        // ###########################################################################################
        private async Task ExecuteOscilloscopeKeyboardCommandAsync(Func<Task> commandAsync)
        {
            try
            {
                await commandAsync();
            }
            finally
            {
                this._lastOscilloscopeKeyboardCommandUtc = DateTime.UtcNow;
                Interlocked.Exchange(ref this._oscilloscopeKeyboardCommandInFlight, 0);
            }
        }

        // ###########################################################################################
        // Queues one TIME/DIV keyboard step on the oscilloscope tab so repeated Add/Subtract
        // keypresses are buffered there instead of being dropped by this popup window.
        // ###########################################################################################
        private bool TryQueueOscilloscopeTimeDivStep(int offset)
        {
            if (this.Owner is not Main mainOwner)
            {
                return true;
            }

            mainOwner.TabOscilloscopeControl.QueueTimeDivKeyboardStep(offset);
            return true;
        }

        // ###########################################################################################
        // Queues one fixed VOLTS/DIV keyboard selection on the oscilloscope tab so rapid numpad
        // requests use latest-wins behavior instead of the popup window's single-flight gate.
        // ###########################################################################################
        private bool TryQueueOscilloscopeVoltsDivSet(double voltsPerDiv)
        {
            if (this.Owner is not Main mainOwner)
            {
                return true;
            }

            mainOwner.TabOscilloscopeControl.QueueVoltsDivKeyboardSet(voltsPerDiv);
            return true;
        }

        // ###########################################################################################
        // Appends the typed digit to the pin buffer, immediately navigates to the first image whose
        // Pin matches the buffer, shows a flash overlay, and clears the buffer after a debounce pause.
        // ###########################################################################################
        private async void HandlePinDigit(char digit)
        {
            this._pinBuffer += digit;

            // Reset the debounce timer so the buffer is only cleared after typing pauses
            this._pinBufferCts?.Cancel();
            this._pinBufferCts = new CancellationTokenSource();
            var bufferCts = this._pinBufferCts;

            var items = this.ThumbnailList.ItemsSource as List<ComponentImageItem>;
            if (items != null && items.Count > 0)
            {
                int matchIndex = items.FindIndex(item =>
                    string.Equals(item.Pin, this._pinBuffer, StringComparison.OrdinalIgnoreCase));

                if (matchIndex >= 0)
                {
                    this.ThumbnailList.SelectedIndex = matchIndex;
                    this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
                    this.ShowPinFlashAsync($"{this._pinBuffer}");
                }
                else
                {
                    this.ShowPinFlashAsync("Not found");
                }
            }

            try
            {
                await Task.Delay(600, bufferCts.Token);
                this._pinBuffer = string.Empty;
            }
            catch (OperationCanceledException) { }
        }

        // ###########################################################################################
        // Displays a large centered "Pin X" label over the main image for 1.2 seconds.
        // Cancels and replaces any currently running flash.
        // ###########################################################################################
        private async void ShowPinFlashAsync(string text)
        {
            this._pinFlashCts?.Cancel();
            this._pinFlashCts = new CancellationTokenSource();
            var cts = this._pinFlashCts;

            this.PinFlashText.Text = text;
            this.PinFlashBorder.IsVisible = true;

            try
            {
                await Task.Delay(800, cts.Token);
                this.PinFlashBorder.IsVisible = false;
            }
            catch (OperationCanceledException) { }
        }

        // ###########################################################################################
        // Intercepts scroll wheel events at the tunnel phase and maps them to thumbnail navigation.
        // Scroll up → next (right), scroll down → previous (left).
        // ###########################################################################################
        private void OnWindowPointerWheelChanged(object? sender, PointerWheelEventArgs e)
        {
            if (e.Delta.Y > 0)
                this.NavigateThumbnails(1);
            else if (e.Delta.Y < 0)
                this.NavigateThumbnails(-1);

            e.Handled = true;
        }

        // ###########################################################################################
        // Intercepts scroll wheel events at the tunnel phase on the left panel and maps them to
        // thumbnail navigation. Scroll up → next (right), scroll down → previous (left).
        // ###########################################################################################
        private void OnLeftPanelPointerWheelChanged(object? sender, PointerWheelEventArgs e)
        {
            var posInContainer = e.GetPosition(this.MainImageClickArea);
            bool isPointerOverImage = posInContainer.X >= 0 && posInContainer.Y >= 0 &&
                                      posInContainer.X <= this.MainImageClickArea.Bounds.Width &&
                                      posInContainer.Y <= this.MainImageClickArea.Bounds.Height;

            if (string.Equals(UserSettings.ComponentInfoScrollAction, "Image zoom", StringComparison.OrdinalIgnoreCase) && isPointerOverImage)
            {
                // We base our scaling layout transforms natively on the inner image dimensions accurately.
                var pos = e.GetPosition(this.MainImageContainer);
                double delta = ViewportMath.ComputeWheelZoomFactor(e.Delta.Y, 1.2);

                double newScale = this._imageMatrix.M11 * delta;

                // Stop zooming out past the original 100% boundary limit. Snaps back precisely to exact initial layout matrix limits.
                if (newScale <= 1.0)
                {
                    this.ResetImageZoom();
                    e.Handled = true;
                    return;
                }

                if (newScale > 10.0)
                    return;

                var zoomMatrix = Matrix.CreateTranslation(-pos.X, -pos.Y)
                               * Matrix.CreateScale(delta, delta)
                               * Matrix.CreateTranslation(pos.X, pos.Y);

                this._imageMatrix = zoomMatrix * this._imageMatrix;

                if (this.MainImageContainer.RenderTransform is MatrixTransform mt)
                    mt.Matrix = this._imageMatrix;
                else
                    this.MainImageContainer.RenderTransform = new MatrixTransform(this._imageMatrix);

                e.Handled = true;
            }
            else
            {
                // Original navigation mode or pointer is located securely over the thumbnail panel
                if (e.Delta.Y > 0)
                    this.NavigateThumbnails(1);
                else if (e.Delta.Y < 0)
                    this.NavigateThumbnails(-1);

                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Handles panning setup with right clicks. Left clicks clear zoom and jump to first thumbnail.
        // ###########################################################################################
        private void OnMainImageClickAreaPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            var pointer = e.GetCurrentPoint(this.MainImageClickArea);

            if (pointer.Properties.IsRightButtonPressed)
            {
                this._isPanningImage = true;
                this._panStartPoint = e.GetPosition(this.MainImageClickArea);
                this._panStartMatrix = this._imageMatrix;
                this.MainImageClickArea.Cursor = new Cursor(StandardCursorType.SizeAll);
                e.Pointer.Capture(this.MainImageClickArea);
                e.Handled = true;
            }
            else if (pointer.Properties.IsLeftButtonPressed)
            {
                this.ThumbnailList.SelectedIndex = 0;
                this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Performs matrix transforms while capturing mouse to visually drag the zoomed image location.
        // ###########################################################################################
        private void OnMainImageClickAreaPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this._isPanningImage)
            {
                var point = e.GetPosition(this.MainImageClickArea);
                var delta = point - this._panStartPoint;
                this._imageMatrix = this._panStartMatrix * Matrix.CreateTranslation(delta.X, delta.Y);
                if (this.MainImageContainer.RenderTransform is MatrixTransform mt)
                    mt.Matrix = this._imageMatrix;
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Handles panning setup with right clicks. Left clicks clear zoom and jump to first thumbnail.
        // ###########################################################################################
        private void OnMainImageContainerPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            var pointer = e.GetCurrentPoint(this.MainImageContainer);

            if (pointer.Properties.IsRightButtonPressed)
            {
                this._isPanningImage = true;
                this._panStartPoint = e.GetPosition(this.MainImageContainer);
                this._panStartMatrix = this._imageMatrix;
                this.MainImageContainer.Cursor = new Cursor(StandardCursorType.SizeAll);
                e.Pointer.Capture(this.MainImageContainer);
                e.Handled = true;
            }
            else if (pointer.Properties.IsLeftButtonPressed)
            {
                this.ThumbnailList.SelectedIndex = 0;
                this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Performs matrix transforms while capturing mouse to visually drag the zoomed image location.
        // ###########################################################################################
        private void OnMainImageContainerPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this._isPanningImage)
            {
                var point = e.GetPosition(this.MainImageContainer);
                var delta = point - this._panStartPoint;
                this._imageMatrix = this._panStartMatrix * Matrix.CreateTranslation(delta.X, delta.Y);
                ((MatrixTransform)this.MainComponentImage.RenderTransform!).Matrix = this._imageMatrix;
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Securely resets zoom matrices entirely to fit exactly and perfectly bounds back inside margins.
        // ###########################################################################################
        private void ResetImageZoom()
        {
            this._imageMatrix = Matrix.Identity;
            if (this.MainImageContainer.RenderTransform is MatrixTransform mt)
            {
                mt.Matrix = this._imageMatrix;
            }
            else
            {
                this.MainImageContainer.RenderTransform = new MatrixTransform(this._imageMatrix);
            }
        }

        // ###########################################################################################
        // Finalizes drag status and releases cursor holds natively back to system expectations.
        // ###########################################################################################
        private void OnMainImageClickAreaPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this._isPanningImage)
            {
                this._isPanningImage = false;
                this.MainImageClickArea.Cursor = Cursor.Default;
                e.Pointer.Capture(null);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Finalizes drag status and releases cursor holds natively back to system expectations.
        // ###########################################################################################
        private void OnMainImageContainerPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this._isPanningImage)
            {
                this._isPanningImage = false;
                this.MainImageContainer.Cursor = Cursor.Default;
                e.Pointer.Capture(null);
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Saves the mousewheel zoom switch state to the persisted component info settings.
        // ###########################################################################################
        private void OnMousewheelZoomSwitchChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.ComponentInfoScrollAction =
                this.MousewheelZoomCheckBox.IsChecked == true
                    ? "Image zoom"
                    : "Image change";
        }

        // ###########################################################################################
        // Clicking the main image jumps back to the first thumbnail, identical to pressing Space.
        // ###########################################################################################
        private void OnMainImagePointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this.ThumbnailList.SelectedIndex = 0;
            this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
            e.Handled = true;
        }

        // ###########################################################################################
        // Updates popup content with the currently targeted component and loads matching images.
        // ###########################################################################################
        public void SetComponent(
            string boardLabel,
            string displayText,
            List<ComponentEntry> componentEntries,
            List<ComponentImageEntry> componentImages,
            List<ComponentLocalFileEntry> localFiles,
            List<ComponentLinkEntry> links,
            string region,
            string dataRoot,
            bool hasExplicitRegionComponents)
        {
            this.ClearTemporaryCapturedOscilloscopeImage();

            // Reset pin navigation state whenever a new component is loaded
            this._pinBufferCts?.Cancel();
            this._pinBuffer = string.Empty;
            this._pinFlashCts?.Cancel();
            this.PinFlashBorder.IsVisible = false;

            // Local files — filter by board label
            var matchingLocalFiles = localFiles
                .Where(f => string.Equals(f.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                .ToList();
            bool hasLocalFiles = matchingLocalFiles.Count > 0;
            this.LocalFilesSection.IsVisible = hasLocalFiles;
            this.LocalFilesItemsControl.ItemsSource = hasLocalFiles
                ? matchingLocalFiles
                    .Select(f => new ComponentLocalFileItem
                    {
                        Name = f.Name,
                        FullPath = Path.Combine(dataRoot, f.File.Replace('/', Path.DirectorySeparatorChar))
                    })
                    .ToList()
                : null;

            // Links — filter by board label
            var matchingLinks = links
                .Where(l => string.Equals(l.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                .ToList();
            bool hasLinks = matchingLinks.Count > 0;
            this.LinksSection.IsVisible = hasLinks;
            this.LinksItemsControl.ItemsSource = hasLinks
                ? matchingLinks
                    .Select(l => new ComponentLinkItem
                    {
                        Name = l.Name,
                        Url = l.Url
                    })
                    .ToList()
                : null;

            // Store state for region toggling; reset local region to the global value on each load
            this._boardLabel = boardLabel;
            this._displayTextFallback = displayText;
            this._allComponentEntries = componentEntries;
            this._allComponentImages = componentImages;
            this._dataRoot = dataRoot;
            this._localRegion = region;
            this._hasExplicitRegionComponents = hasExplicitRegionComponents;
            this.UpdateRegionButtonsState();

            // Reset selection on initial load so a lingering pin from a previous component
            // is never accidentally matched against this component's image list
            this.RefreshImages(resetSelection: true);
        }

        // ###########################################################################################
        // Updates the main image, NoImageText, info overlay, counter and note when selection changes.
        // Also schedules a debounced oscilloscope auto-sync when the selected image contains scope data.
        // ###########################################################################################
        private void OnThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressThumbnailSelection)
                return;

            this.ClearTemporaryCapturedOscilloscopeImage();
            this.ResetImageZoom();

            var selected = this.ThumbnailList.SelectedItem as ComponentImageItem;
            this.MainComponentImage.Source = selected?.ImageSource;
            this.NoImageText.IsVisible = selected?.ImageSource == null;
            this.UpdateInfoOverlay();
            this.UpdateImageCounter();
            this.UpdateImageNote(selected);
            this.ScheduleSelectedOscilloscopeImageSync(selected);
        }

        // ###########################################################################################
        // Opens the clicked component local file in the OS default application.
        // ###########################################################################################
        private void OnLocalFileButtonClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button { Tag: ComponentLocalFileItem item })
                return;

            this.OpenExternalTarget(item.FullPath);
        }

        // ###########################################################################################
        // Opens the clicked component link in the OS default browser.
        // ###########################################################################################
        private void OnLinkButtonClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button { Tag: ComponentLinkItem item })
                return;

            this.OpenExternalTarget(item.Url);
        }

        // ###########################################################################################
        // Opens a file path or URL using the operating system's default handler only after strict
        // validation. URLs are limited to HTTP/HTTPS and local files must remain inside data-root.
        // ###########################################################################################
        private void OpenExternalTarget(string target)
        {
            ExternalTargetLauncher.TryOpen(target, this._dataRoot);
        }

        // ###########################################################################################
        // Opens the help page describing numpad oscilloscope controls.
        // ###########################################################################################
        private void OnNumpadOscilloscopeHelpClick(object? sender, RoutedEventArgs e)
        {
            this.OpenExternalTarget("https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Controlling-oscilloscope-with-keyboard");
        }

        // ###########################################################################################
        // Opens the help page describing MiniPro programmer usage.
        // ###########################################################################################
        private void OnMiniProHelpClick(object? sender, RoutedEventArgs e)
        {
            this.OpenExternalTarget("https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/MiniPro-programmer");
        }

        // ###########################################################################################
        // Opens the help page describing oscilloscope synchronization.
        // ###########################################################################################
        private void OnSyncOscilloscopeHelpClick(object? sender, RoutedEventArgs e)
        {
            this.OpenExternalTarget("https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Synchronize-oscilloscope");
        }

        // ###########################################################################################
        // Refreshes the image info overlays from the currently selected thumbnail item.
        // ###########################################################################################
        private void UpdateInfoOverlay()
        {
            var selected = this.ThumbnailList.SelectedItem as ComponentImageItem;

            string? pin = selected?.Label;
            string? name = selected?.Name;

            // Suppress the pin/name label when it is identical to the name label
            bool pinSameAsName = !string.IsNullOrWhiteSpace(pin) &&
                                 string.Equals(pin, name, StringComparison.OrdinalIgnoreCase);

            SetInfoLabel(this.InfoPinBorder, this.InfoPinText, pinSameAsName ? null : pin);
            SetInfoLabel(this.InfoNameBorder, this.InfoNameText, name);
            SetInfoLabel(this.InfoOscBorder, this.InfoOscText, selected?.ExpectedOscilloscopeReading);

            SetInfoLabelPair(
                this.InfoScopeTimeDivBorder,
                this.InfoScopeTimeDivPrefixText,
                this.InfoScopeTimeDivValueText,
                "T/DIV:",
                selected?.TimeDiv);

            SetInfoLabelPair(
                this.InfoScopeVoltsDivBorder,
                this.InfoScopeVoltsDivPrefixText,
                this.InfoScopeVoltsDivValueText,
                "V/DIV:",
                selected?.VoltsDiv);

            SetInfoLabelPair(
                this.InfoScopeTriggerBorder,
                this.InfoScopeTriggerPrefixText,
                this.InfoScopeTriggerValueText,
                "T:",
                selected?.TriggerLevelVolts);
        }                

        // ###########################################################################################
        // Shows or hides the "Image note" section based on the selected thumbnail's note text.
        // ###########################################################################################
        private void UpdateImageNote(ComponentImageItem? item)
        {
            string? note = item?.Note;
            bool show = !string.IsNullOrWhiteSpace(note);
            this.ImageNoteSection.IsVisible = show;
            if (show)
                this.InfoNote.Text = note!.Trim();
        }

        // ###########################################################################################
        // Shows or hides a single info label border depending on whether value is non-empty.
        // ###########################################################################################
        private static void SetInfoLabel(Border border, TextBlock textBlock, string? value)
        {
            bool show = !string.IsNullOrWhiteSpace(value);
            border.IsVisible = show;
            if (show)
                textBlock.Text = value;
        }

        // ###########################################################################################
        // Shows or hides a compact two-part info label where the prefix stays normal and the value
        // is rendered bold. Hidden when the value is empty.
        // ###########################################################################################
        private static void SetInfoLabelPair(Border border, TextBlock prefixTextBlock, TextBlock valueTextBlock, string prefix, string? value)
        {
            string trimmed = ScopeFormatting.NormalizeScopeOverlayValue(value);
            bool show = !string.IsNullOrWhiteSpace(trimmed);

            border.IsVisible = show;
            if (!show)
            {
                prefixTextBlock.Text = string.Empty;
                valueTextBlock.Text = string.Empty;
                return;
            }

            prefixTextBlock.Text = prefix;
            valueTextBlock.Text = trimmed;
        }

        // ###########################################################################################
        // Refreshes the "Image X of Y" counter; hidden when fewer than 2 images are loaded.
        // ###########################################################################################
        private void UpdateImageCounter()
        {
            var items = this.ThumbnailList.ItemsSource as List<ComponentImageItem>;
            int total = items?.Count ?? 0;
            int index = this.ThumbnailList.SelectedIndex;

            bool show = total > 1 && index >= 0;
            this.ImageCounterBorder.IsVisible = show;

            if (show)
                this.ImageCounterText.Text = $"Image {index + 1} of {total}";
        }

        // ###########################################################################################
        // Moves the thumbnail selection left or right by the given delta and scrolls it into view.
        // Wraps around: going left from the first item lands on the last, and vice versa.
        // ###########################################################################################
        private void NavigateThumbnails(int delta)
        {
            var items = this.ThumbnailList.ItemsSource as List<ComponentImageItem>;
            if (items == null || items.Count == 0)
                return;

            int newIndex = (this.ThumbnailList.SelectedIndex + delta + items.Count) % items.Count;
            if (newIndex == this.ThumbnailList.SelectedIndex)
                return;

            this.ThumbnailList.SelectedIndex = newIndex;
            this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
        }

        // ###########################################################################################
        // Loads component images on a background thread, then populates the gallery and main image.
        // Entries without a File value are excluded so no empty thumbnail placeholders are shown.
        // ###########################################################################################
        private async void LoadImagesAsync(List<ComponentImageEntry> entries, string dataRoot, string? preservePin = null)
        {
            this._loadCts?.Cancel();
            this._loadCts = new CancellationTokenSource();
            var cts = this._loadCts;

            var displayableEntries = entries
                .Where(ComponentImageQueries.HasDisplayableImageFile)
                .ToList();

            if (displayableEntries.Count == 0)
            {
                this.DisposeLoadedBitmaps();
                this.ScheduleSelectedOscilloscopeImageSync(null);
                return;
            }

            var loaded = await Task.Run(() =>
            {
                var result = new List<(ComponentImageEntry Entry, Bitmap? Bitmap)>();

                foreach (var entry in displayableEntries)
                {
                    if (cts.Token.IsCancellationRequested)
                        break;

                    Bitmap? bitmap = null;

                    try
                    {
                        var fullPath = Path.Combine(dataRoot, entry.File.Replace('/', Path.DirectorySeparatorChar));
                        if (File.Exists(fullPath))
                        {
                            bitmap = new Bitmap(fullPath);
                        }
                    }
                    catch
                    {
                        // A malformed File value (e.g. from hand-edited or contributed board data)
                        // should skip this one image, not crash the whole load - this method has
                        // no caller-side exception handling since it is async void.
                    }

                    result.Add((entry, bitmap));
                }

                return result;
            });

            if (cts.Token.IsCancellationRequested)
            {
                foreach (var (_, bmp) in loaded)
                    bmp?.Dispose();
                return;
            }

            // Stage new bitmaps before touching the UI so old images stay visible until the swap
            var oldBitmaps = new List<Bitmap>(this._loadedBitmaps);
            this._loadedBitmaps.Clear();

            foreach (var (_, bmp) in loaded)
            {
                if (bmp != null)
                    this._loadedBitmaps.Add(bmp);
            }

            var items = loaded
                .Select(x => new ComponentImageItem
                {
                    ImageSource = x.Bitmap,
                    Label = ComponentImageQueries.BuildImageLabel(x.Entry),
                    Pin = x.Entry.Pin.Trim(),
                    Name = x.Entry.Name,
                    ExpectedOscilloscopeReading = x.Entry.ExpectedOscilloscopeReading,
                    TimeDiv = x.Entry.TimeDiv,
                    VoltsDiv = x.Entry.VoltsDiv,
                    TriggerLevelVolts = x.Entry.TriggerLevelVolts,
                    Note = x.Entry.Note,
                    SourceEntry = x.Entry
                })
                .ToList();

            // Resolve target index: restore same pin if found in the new set, otherwise first item
            int targetIndex = 0;
            if (!string.IsNullOrEmpty(preservePin))
            {
                int pinIndex = items.FindIndex(item =>
                    string.Equals(item.Pin, preservePin, StringComparison.OrdinalIgnoreCase));
                if (pinIndex >= 0)
                    targetIndex = pinIndex;
            }

            // Suppress selection events during the atomic ItemsSource + SelectedIndex swap.
            // Without this, the transient null-selection state would blank the main image (blink).
            this._suppressThumbnailSelection = true;
            this.ThumbnailList.ItemsSource = items;
            if (items.Count > 0)
            {
                this.ThumbnailList.SelectedIndex = targetIndex;
                this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
            }
            this._suppressThumbnailSelection = false;

            // Manually apply what OnThumbnailSelectionChanged would have done
            var selected = this.ThumbnailList.SelectedItem as ComponentImageItem;
            this.MainComponentImage.Source = selected?.ImageSource;
            this.NoImageText.IsVisible = selected?.ImageSource == null;
            this.UpdateInfoOverlay();
            this.UpdateImageCounter();
            this.UpdateImageNote(selected);
            this.ScheduleSelectedOscilloscopeImageSync(selected);

            // Dispose the previous bitmaps only after the UI has fully transitioned to the new set
            foreach (var bmp in oldBitmaps)
                bmp.Dispose();
        }

        // ###########################################################################################
        // Clears UI image references, resets the image note section, and disposes loaded bitmaps.
        // ###########################################################################################
        private void DisposeLoadedBitmaps()
        {
            this.ClearTemporaryCapturedOscilloscopeImage();
            this.MainComponentImage.Source = null;
            this.ThumbnailList.ItemsSource = null;
            this.ImageNoteSection.IsVisible = false;

            foreach (var bmp in this._loadedBitmaps)
                bmp.Dispose();

            this._loadedBitmaps.Clear();
        }

        // ###########################################################################################
        // Updates the PAL and NTSC button captions with per-region image counters.
        // Empty image regions are included in both counters.
        // ###########################################################################################
        private void UpdateRegionButtonCounters()
        {
            int palCount = ComponentImageQueries.CountImagesForRegion(this._allComponentImages, this._boardLabel, "PAL");
            int ntscCount = ComponentImageQueries.CountImagesForRegion(this._allComponentImages, this._boardLabel, "NTSC");

            this.PalRegionButton.Content = $"PAL ({palCount})";
            this.NtscRegionButton.Content = $"NTSC ({ntscCount})";
        }

        // ###########################################################################################
        // Re-filters the stored image list for the current local region and triggers an async reload.
        // Entries without a File value are excluded from the thumbnail gallery.
        // ###########################################################################################
        private void RefreshImages(bool resetSelection = false)
        {
            // Update text fields immediately so they reflect the new region before images finish loading
            this.RefreshComponentText();

            // Capture the current pin so the same gallery position can be restored after the reload,
            // unless this is a full component reset where selection should always start at index 0.
            var currentPin = resetSelection ? null : (this.ThumbnailList.SelectedItem as ComponentImageItem)?.Pin;

            var matchingEntries = this._allComponentImages
                .Where(img =>
                    string.Equals(img.BoardLabel, this._boardLabel, StringComparison.OrdinalIgnoreCase) &&
                    ComponentImageQueries.HasDisplayableImageFile(img) &&
                    ComponentImageQueries.IsImageVisibleInRegion(img, this._localRegion))
                .ToList();

            this.LoadImagesAsync(matchingEntries, this._dataRoot, currentPin);
        }

        // ###########################################################################################
        // Switches the local region to PAL and reloads images without touching the global setting.
        // ###########################################################################################
        private void OnPalRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;
            this._localRegion = "PAL";
            this.UpdateRegionButtonsState();
            this.RefreshImages();
        }

        // ###########################################################################################
        // Switches the local region to NTSC and reloads images without touching the global setting.
        // ###########################################################################################
        private void OnNtscRegionClick(object? sender, RoutedEventArgs e)
        {
            if (this._suppressRegionToggle)
                return;
            this._localRegion = "NTSC";
            this.UpdateRegionButtonsState();
            this.RefreshImages();
        }

        // ###########################################################################################
        // Updates the region toggle and button states to match the current local region.
        // Hides only the region buttons when the board has no explicit PAL/NTSC components,
        // while keeping the switches and Close button visible.
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;
            bool isNtsc = string.Equals(this._localRegion, "NTSC", StringComparison.OrdinalIgnoreCase);

            this.PalRegionButton.IsVisible = this._hasExplicitRegionComponents;
            this.NtscRegionButton.IsVisible = this._hasExplicitRegionComponents;
            this.UpdateRegionButtonCounters();

            if (this.PalRegionButton.Parent is Grid footerGrid && footerGrid.ColumnDefinitions.Count >= 5)
            {
                footerGrid.ColumnDefinitions[0].Width = this._hasExplicitRegionComponents
                    ? GridLength.Auto
                    : new GridLength(0, GridUnitType.Pixel);

                footerGrid.ColumnDefinitions[1].Width = this._hasExplicitRegionComponents
                    ? new GridLength(12, GridUnitType.Pixel)
                    : new GridLength(0, GridUnitType.Pixel);
            }

            this.NtscRegionButton.Classes.Set("active", isNtsc);
            this.PalRegionButton.Classes.Set("active", !isNtsc);

            this._suppressRegionToggle = false;
            this.UpdateRegionLabel();
        }

        // ###########################################################################################
        // Updates the region label overlay in the top-left corner using the same color schema
        // as the region label in the Schematics tab, bound dynamically to the current local region.
        // ###########################################################################################
        private void UpdateRegionLabel()
        {
            string colorPrefix = this._localRegion.ToUpperInvariant() switch
            {
                "PAL" => "Schematics_Region_PAL",
                "NTSC" => "Schematics_Region_NTSC",
                _ => "SchematicsRegion"
            };

            this.InfoRegionText.Text = this._localRegion;
            this.InfoRegionBorder.IsVisible = this._hasExplicitRegionComponents;

            this.InfoRegionBorder.Bind(
                Border.BackgroundProperty,
                this.GetResourceObservable($"{colorPrefix}_Bg"));

            this.InfoRegionBorder.Bind(
                Border.BorderBrushProperty,
                this.GetResourceObservable($"{colorPrefix}_Border"));

            this.InfoRegionText.Bind(
                TextBlock.ForegroundProperty,
                this.GetResourceObservable($"{colorPrefix}_Fg"));
        }

        // ###########################################################################################
        // Updates all region-sensitive text fields (title, category/part-number, description)
        // to reflect the current local region without affecting the global setting.
        // ###########################################################################################
        private void RefreshComponentText()
        {
            var entry = ComponentImageQueries.PickComponentEntry(this._allComponentEntries, this._localRegion);

            // Title: BoardLabel | FriendlyName | TechnicalNameOrValue (non-empty parts joined)
            var titleParts = new List<string>(3);
            if (!string.IsNullOrWhiteSpace(this._boardLabel))
                titleParts.Add(this._boardLabel.Trim());
            if (!string.IsNullOrWhiteSpace(entry?.FriendlyName))
                titleParts.Add(entry.FriendlyName.Trim());
            if (!string.IsNullOrWhiteSpace(entry?.TechnicalNameOrValue))
                titleParts.Add(entry.TechnicalNameOrValue.Trim());

            string titleText = titleParts.Count > 0
                ? string.Join(" | ", titleParts)
                : this._displayTextFallback;

            this.TitleText.Text = titleText;
            this.ApplyOscilloscopeSessionTitleState();

            // Category | Part-number
            string category = entry?.Category ?? string.Empty;
            string partNumber = entry?.PartNumber ?? string.Empty;
            var catPartParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(category))
                catPartParts.Add(category.Trim());
            if (!string.IsNullOrWhiteSpace(partNumber))
                catPartParts.Add(partNumber.Trim());
            bool hasCatPart = catPartParts.Count > 0;
            this.InfoCategoryPartNumber.IsVisible = hasCatPart;
            if (hasCatPart)
                this.InfoCategoryPartNumber.Text = string.Join(" | ", catPartParts);

            // One-liner description
            string description = entry?.Description ?? string.Empty;
            bool hasDescription = !string.IsNullOrWhiteSpace(description);
            this.OneLinerSection.IsVisible = hasDescription;
            this.InfoDescription.Text = hasDescription ? description.Trim() : string.Empty;

            this.UpdateTestAffordance(entry);
        }

        private IcTestEntry? _activeTestEntry;

// ###########################################################################################
// Shows the "Test this IC" affordance only for IC components that have a test-catalogue
// entry (matched on Technical name/value, not the unreliable Part-number). Functional-only
// parts show a disabled button with an honest label.
// ###########################################################################################
private void UpdateTestAffordance(ComponentEntry? entry)
{
    // The overlay shows a snapshot for whichever IC was active when opened — if the
    // displayed component changes underneath it (e.g. clicking another chip on the
    // schematic), it would otherwise keep showing stale test content.
    this.OnIcTestPanelCloseRequested();

    this._activeTestEntry = null;
    if (!UserSettings.EnableMiniproExperimentalMode)
    {
        this.TestSection.IsVisible = false;
        return;
    }
    bool isIc = string.Equals(entry?.Category?.Trim(), "IC", StringComparison.OrdinalIgnoreCase);
    var cat = isIc ? IcTestCatalogue.Lookup(entry!.TechnicalNameOrValue) : null;
    if (cat is null)
    {
        this.TestSection.IsVisible = false;
        return;
    }
    this._activeTestEntry = cat;
    Logger.Info($"IC test affordance shown for [{this._boardLabel}] [{entry!.TechnicalNameOrValue}] (kind={cat.Kind}, support={cat.Support})");
    this.TestSection.IsVisible = true;
    this.TestButton.IsEnabled = true;   // always openable — the panel disables Run itself for non-testable parts
    this.TestButton.Content = cat.IsTestable
        ? "Test IC with MiniPro programmer"
        : cat.IsFunctionalOnly
            ? "View test info (functional-only)"
            : "View test info (not supported)";
}

private void OnTestClick(object? sender, RoutedEventArgs e)
{
    if (this._activeTestEntry is null) return;
    Logger.Info($"Opening IC test panel for [{this._boardLabel}] [{this._activeTestEntry.Id}]");
    this.IcTestPanel.Load(this._activeTestEntry, this._boardLabel);
    this.IcTestPanel.IsVisible = true;
}

private void OnIcTestPanelCloseRequested() => this.IcTestPanel.IsVisible = false;

        // ###########################################################################################
        // Queues oscilloscope auto-sync for the currently selected image.
        // The oscilloscope tab handles debounce and latest-wins processing.
        // ###########################################################################################
        private void ScheduleSelectedOscilloscopeImageSync(ComponentImageItem? selectedItem)
        {
            if (this.Owner is not Main mainOwner)
            {
                return;
            }

            if (!this.CanSendOscilloscopeCommands() ||
                !ComponentImageQueries.IsOscilloscopeImage(selectedItem?.SourceEntry))
            {
                mainOwner.TabOscilloscopeControl.QueueComponentImageOscilloscopeSync(null);
                return;
            }

            mainOwner.TabOscilloscopeControl.QueueComponentImageOscilloscopeSync(selectedItem!.SourceEntry);
        }

        // ###########################################################################################
        // Updates the popup window title with the oscilloscope session suffix when the oscilloscope
        // has been connected at least once for the current app session.
        // ###########################################################################################
        public void UpdateOscilloscopeSessionTitleState(bool hasSeenOscilloscopeSession, bool hasActiveOscilloscopeSession)
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    () => this.UpdateOscilloscopeSessionTitleState(hasSeenOscilloscopeSession, hasActiveOscilloscopeSession),
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            this._hasSeenOscilloscopeSessionTitleState = hasSeenOscilloscopeSession;
            this._hasActiveOscilloscopeSessionTitleState = hasActiveOscilloscopeSession;

            this.UpdateOscilloscopeControlsAvailability();
            this.ApplyOscilloscopeSessionTitleState();
        }

        // ###########################################################################################
        // Applies the oscilloscope session suffix to the popup title while preserving the current
        // component-specific base title text.
        // ###########################################################################################
        private void ApplyOscilloscopeSessionTitleState()
        {
            string baseTitle = !string.IsNullOrWhiteSpace(this.TitleText.Text)
                ? this.TitleText.Text
                : ScopeFormatting.GetMainWindowTitleBase(this.Title ?? string.Empty);

            // Unlike the main window this popup does not report a pending auto-connect - it only has
            // something to say once a session has actually existed. A session can also still be live
            // from before the oscilloscope tab was switched off, and the builder drops the suffix for
            // that case: the popup must say nothing about an oscilloscope the user has turned off.
            this.Title = ScopeFormatting.BuildOscilloscopeWindowTitle(
                baseTitle,
                UserSettings.EnableNetworkConnectedOscilloscopeTab,
                this._hasSeenOscilloscopeSessionTitleState,
                this._hasActiveOscilloscopeSessionTitleState);
        }

        // ###########################################################################################
        // Saves the numpad oscilloscope control switch state to the persisted component info settings.
        // ###########################################################################################
        private void OnNumpadOscilloscopeSwitchChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.ComponentInfoKeyboardHandling =
                this.NumpadOscilloscopeSwitch.IsChecked == true
                    ? "Control oscilloscope"
                    : "Control image pin selection";
        }

        // ###########################################################################################
        // Returns the oscilloscope keyboard command throttle interval using the same Debounce-Time
        // value that the oscilloscope tab already reads from the main Excel data file.
        // ###########################################################################################
        private TimeSpan GetOscilloscopeKeyboardCommandMinimumInterval()
        {
            if (this.Owner is not Main mainOwner)
            {
                return TimeSpan.FromMilliseconds(250);
            }

            int debounceDelayMilliseconds = mainOwner.TabOscilloscopeControl.GetComponentImageSyncDebounceDelayMilliseconds();
            return TimeSpan.FromMilliseconds(Math.Max(0, debounceDelayMilliseconds));
        }

        // ###########################################################################################
        // Requests direct execution of a named oscilloscope command palette on the active session
        // using the oscilloscope tab's existing SCPI pipeline and command logging.
        // ###########################################################################################
        private async Task RunOscilloscopePaletteAsync(ScopeCommandPalette palette)
        {
            if (this.Owner is not Main mainOwner)
            {
                return;
            }

            await mainOwner.TabOscilloscopeControl.RunPaletteAsync(palette, CancellationToken.None);
        }

        // ###########################################################################################
        // Queues one trigger-level keyboard step on the oscilloscope tab so repeated Up/Down
        // keypresses are buffered there instead of being dropped by this popup window.
        // ###########################################################################################
        private bool TryQueueOscilloscopeTriggerLevelStep(int direction)
        {
            if (this.Owner is not Main mainOwner)
            {
                return true;
            }

            mainOwner.TabOscilloscopeControl.QueueTriggerLevelKeyboardStep(direction);
            return true;
        }

        // ###########################################################################################
        // Hides the popup's oscilloscope rows entirely when the user has turned off the network
        // connected oscilloscope tab, and otherwise enables them only while an active oscilloscope
        // session exists. The checked states are preserved so behavior resumes automatically after
        // reconnect - or after the oscilloscope tab is switched back on.
        //
        // Clearing IsEnabled is what actually disables the feature: CanSendOscilloscopeCommands and
        // the numpad key handler both gate on it, so nothing can reach the oscilloscope from a popup
        // whose rows are hidden.
        // ###########################################################################################
        private void UpdateOscilloscopeControlsAvailability()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.UpdateOscilloscopeControlsAvailability,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            bool isOscilloscopeTabEnabled = UserSettings.EnableNetworkConnectedOscilloscopeTab;
            bool isOscilloscopeAvailable = isOscilloscopeTabEnabled && this._hasActiveOscilloscopeSessionTitleState;

            this.SyncOscilloscopeRow.IsVisible = isOscilloscopeTabEnabled;
            this.NumpadOscilloscopeRow.IsVisible = isOscilloscopeTabEnabled;

            this.NumpadOscilloscopeSwitch.IsEnabled = isOscilloscopeAvailable;
            this.SyncOscilloscopeCheckBox.IsEnabled = isOscilloscopeAvailable;

            // Drop any sync this popup still has queued, the same way OnSyncOscilloscopeCheckBoxChanged
            // does when sync is switched off - otherwise a debounced request from just before the tab
            // was disabled could still reach the oscilloscope.
            if (!isOscilloscopeTabEnabled && this.Owner is Main mainOwner)
            {
                mainOwner.TabOscilloscopeControl.QueueComponentImageOscilloscopeSync(null);
            }
        }

        // ###########################################################################################
        // Captures a live oscilloscope image for the currently selected component image, saves it
        // through the oscilloscope tab, and temporarily shows it in the large preview area.
        // ###########################################################################################
        private async Task CaptureAndDisplayOscilloscopeImageAsync()
        {
            if (this.Owner is not Main mainOwner)
            {
                return;
            }

            ComponentImageItem? selectedItem = this.ThumbnailList.SelectedItem as ComponentImageItem;
            ComponentImageEntry? selectedEntry = selectedItem?.SourceEntry;

            if (selectedEntry == null)
            {
                return;
            }

            string displayedRegion = !string.IsNullOrWhiteSpace(selectedEntry.Region)
                ? selectedEntry.Region.Trim()
                : this._localRegion;

            await Dispatcher.UIThread.InvokeAsync(
                this.ShowFetchScopeImageOverlay,
                DispatcherPriority.Background);

            try
            {
                string? savedFilePath = await mainOwner.TabOscilloscopeControl.CaptureAndSaveOscilloscopeImageAsync(
                    selectedEntry,
                    displayedRegion,
                    CancellationToken.None);

                if (string.IsNullOrWhiteSpace(savedFilePath) || !File.Exists(savedFilePath))
                {
                    return;
                }

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.ShowTemporaryCapturedOscilloscopeImage(savedFilePath),
                    DispatcherPriority.Background);
            }
            finally
            {
                await Dispatcher.UIThread.InvokeAsync(
                    this.HideFetchScopeImageOverlay,
                    DispatcherPriority.Background);
            }
        }

        // ###########################################################################################
        // Loads a just-captured oscilloscope image from disk and temporarily replaces the large
        // image preview until the user navigates away from the current thumbnail selection.
        // ###########################################################################################
        private void ShowTemporaryCapturedOscilloscopeImage(string savedFilePath)
        {
            this.ClearTemporaryCapturedOscilloscopeImage();

            this.thisTemporaryCapturedScopeBitmap = new Bitmap(savedFilePath);
            this.MainComponentImage.Source = this.thisTemporaryCapturedScopeBitmap;
            this.NoImageText.IsVisible = false;
            this.CapturedScopeImageText.Text = $"Saved image as [{Path.GetFileName(savedFilePath)}]";
            this.CapturedScopeImageBorder.IsVisible = true;

            this.thisCapturedScopeImagePath = savedFilePath;
            this.UpdateAttachCapturedImageButton();
        }

        // ###########################################################################################
        // The capture the "Attach image to worklog" button acts on - the PNG most recently written
        // by CaptureAndDisplayOscilloscopeImageAsync. Cleared alongside the banner, so the button can
        // never act on a file belonging to an earlier capture.
        // ###########################################################################################
        private string? thisCapturedScopeImagePath;

        // ###########################################################################################
        // Shows the attach button only when there is both a capture to attach AND a workbook to
        // attach it to.
        //
        // Hidden outright when the worklog feature is switched off, matching how the Workbooks tab
        // and the worklog bar are gated (Main.ApplyWorklogBarVisibility) - a user who has turned the
        // feature off should see no trace of it here either. Also hidden when the board has no
        // workbook yet: the dialog's whole premise is filing into the active one, and there is
        // nothing useful to offer without it.
        // ###########################################################################################
        private void UpdateAttachCapturedImageButton()
        {
            bool hasCapture = !string.IsNullOrWhiteSpace(this.thisCapturedScopeImagePath);

            this.AttachCapturedImageButton.IsVisible =
                hasCapture &&
                UserSettings.EnableWorklog &&
                this.ResolveActiveWorkbookForCapture() != null;
        }

        // ###########################################################################################
        // The workbook a capture would be filed into - the board's ACTIVE workbook, resolved through
        // Main so this flow cannot disagree with the worklog bar about which one that is.
        // ###########################################################################################
        private WorkbookRecord? ResolveActiveWorkbookForCapture()
        {
            if (this.Owner is not Main mainOwner)
            {
                return null;
            }

            string boardKey = mainOwner.GetCurrentBoardKey();
            return string.IsNullOrWhiteSpace(boardKey)
                ? null
                : Main.ResolveActiveWorkbookForBoard(boardKey);
        }

        // ###########################################################################################
        // Files the just-captured oscilloscope image into a worklog entry.
        //
        // One modal does the whole job: it names the workbook, ranks the entries (component matches
        // first - see WorklogAttachTargets) and takes the comment. The capture is already safely
        // written to the oscilloscope image folder before any of this, so cancelling here, or a
        // failed attach, costs the user nothing but the filing.
        //
        // "Create new worklog" opens the full editor on a draft with the photo already attached,
        // rather than making the user save an entry and come back for it - probing before anything
        // has been written down is exactly how diagnosis starts. The draft carries no marked area
        // (ShowMarkedArea false, the supported "parked badge" state), because area-marking needs the
        // schematic view and the user is at the bench with a probe in hand.
        // ###########################################################################################
        private async void OnAttachCapturedImageClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AttachCapturedImageToWorklogAsync();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to attach captured oscilloscope image to a worklog: {ex.Message}");
            }
        }

        private async Task AttachCapturedImageToWorklogAsync()
        {
            string? capturedPath = this.thisCapturedScopeImagePath;
            if (string.IsNullOrWhiteSpace(capturedPath) || !File.Exists(capturedPath))
            {
                return;
            }

            var workbook = this.ResolveActiveWorkbookForCapture();
            if (workbook == null)
            {
                return;
            }

            var dialog = new WorklogAttachCaptureWindow();
            dialog.Initialize(
                capturedPath,
                workbook,
                WorklogManager.GetEntries(workbook.Id),
                this._boardLabel);

            var result = await dialog.ShowDialog<WorklogAttachCaptureWindow.AttachResult?>(this);
            if (result == null)
            {
                return;
            }

            if (result.EntryId == null)
            {
                await this.CreateWorklogEntryForCaptureAsync(workbook.Id, capturedPath, result.Comment);
                return;
            }

            // Nothing below this point can leave the entry invisible: it is filed into an existing
            // worklog, which already has whatever schematic and area the user gave it.

            var outcome = WorklogAttachmentWriter.AttachToEntry(
                workbook.Id,
                result.EntryId.Value,
                capturedPath,
                WorklogAttachmentStorage.PhotoFilePrefix,
                result.Comment);

            if (outcome != WorklogAttachmentWriter.AttachOutcome.Added)
            {
                Logger.Warning($"Could not attach captured image to worklog entry [#{result.EntryId.Value}]: {outcome}");
                return;
            }

            this.CapturedScopeImageText.Text =
                $"Attached to worklog #{result.EntryId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)}";

            this.RefreshWorklogSurfacesAfterAttach();
        }

        // ###########################################################################################
        // The "Create new worklog" branch: opens the full editor on a draft carrying the capture.
        //
        // The photo is attached AFTER InitializeForNewEntry, which is the order AttachCapturedPhoto
        // documents - the draft's attachment folder is named after the id that call reserves.
        //
        // The entry MUST be filed against a real schematic. Both surfaces that draw worklog entries
        // filter by SchematicName - the Schematics tab's overlay (RefreshWorklogEntriesList) and the
        // Workbooks board pane - so an entry saved with a blank one is invisible on both, with no
        // way to reach it from the board at all. It shipped that way once and was reported: the
        // worklog saved fine and then appeared nowhere. The schematic showing on the Schematics tab
        // is the right one to use, since that is the board view the user is working against.
        //
        // The AREA is deliberately left unset, with ShowMarkedArea off: this entry was born at the
        // oscilloscope with a probe in hand, not by dragging a rectangle, so there is no area to
        // record. That is the supported "parked pill" state - the entry shows as a "#N" pill in the
        // schematic panel's top-right corner rather than as a rectangle on the board. Ticking "Show
        // marked area" later gives it a real, draggable square (see WorklogDefaultAreaGeometry) -
        // which is why the schematic BITMAP is handed over rather than a null: that geometry needs
        // the board's pixel size, and without it the tick produced the zero-sized rect the geometry
        // exists to prevent. See ResolveSchematicBitmapForCapture.
        // ###########################################################################################
        private async Task CreateWorklogEntryForCaptureAsync(int workbookId, string capturedPath, string comment)
        {
            var editor = new WorklogEntryEditorWindow();

            editor.InitializeForNewEntry(
                workbookId,
                this.ResolveSchematicNameForCapture(),
                default,
                this.ResolveSchematicBitmapForCapture());

            editor.SetShowMarkedAreaForNewEntry(false);

            // The outcome is acted on, not merely logged. A failed copy leaves the Photos section
            // empty, and an editor that opens looking exactly like a successful one says the capture
            // was filed when it was not - along with the comment the user typed into the attach
            // dialog. The banner reports it instead, and the editor is not opened at all: there is
            // nothing to create a worklog around.
            if (!editor.AttachCapturedPhoto(capturedPath, comment))
            {
                Logger.Warning($"Could not attach captured image to a new worklog in workbook [#{workbookId.ToString(System.Globalization.CultureInfo.InvariantCulture)}] - the new worklog was not opened");
                this.CapturedScopeImageText.Text = "Could not attach the image to a worklog";
                return;
            }

            await editor.ShowDialog(this);

            this.RefreshWorklogSurfacesAfterAttach();
        }

        // ###########################################################################################
        // Which schematic a worklog created from a capture belongs to - the one currently showing on
        // the Schematics tab, which is the board view the user is working against.
        // ###########################################################################################
        private string ResolveSchematicNameForCapture()
        {
            if (this.Owner is not Main mainOwner)
            {
                return string.Empty;
            }

            return mainOwner.TabSchematicsControl?.GetCurrentSchematicName() ?? string.Empty;
        }

        // ###########################################################################################
        // The board image the editor draws the marked area against - the same full-resolution bitmap
        // the Schematics tab hands over on its own "Add worklog" path, taken from the same field.
        //
        // Passing null here is not harmless: the editor's EnsureMarkedAreaExistsWhenShown returns
        // early without one, so ticking "Show marked area" on a worklog created from a capture left
        // a zero-sized rectangle - which draws as nothing and can never be grabbed and dragged into
        // place. That is precisely the bug WorklogDefaultAreaGeometry exists to fix, and this entry
        // kind - born parked, with no area at all - is the one most likely to have the box ticked
        // later.
        //
        // The editor does not dispose it (see its own note that the bitmap belongs to the caller),
        // so handing over the tab's live field is safe; a null one simply means no board is loaded.
        // ###########################################################################################
        private Bitmap? ResolveSchematicBitmapForCapture()
        {
            if (this.Owner is not Main mainOwner)
            {
                return null;
            }

            return mainOwner.TabSchematicsControl?.currentFullResBitmap;
        }

        // ###########################################################################################
        // Pushes a just-attached photo out to every worklog surface, through Main.RefreshWorklogBar -
        // the one funnel every worklog change already passes through, so the Workbooks tab, the bar
        // and the Schematics tab's overlay cannot go stale in a case that funnel already handles.
        //
        // Deliberately not poking TabWorkbooks directly: that tab holds decoded schematic bitmaps
        // whose lifetime is tied to its attach/detach cycle, and a second refresh path into it is
        // exactly the re-entrancy its OnDetachedFromVisualTree comment warns about.
        // ###########################################################################################
        private void RefreshWorklogSurfacesAfterAttach()
        {
            if (this.Owner is Main mainOwner)
            {
                mainOwner.RefreshWorklogBar();
            }
        }

        // ###########################################################################################
        // Clears the temporary oscilloscope capture preview and restores the normal main-image
        // behavior driven by the selected thumbnail item.
        // ###########################################################################################
        private void ClearTemporaryCapturedOscilloscopeImage()
        {
            this.HideFetchScopeImageOverlay();
            this.CapturedScopeImageBorder.IsVisible = false;

            // Cleared with the banner, so the attach button can never act on a file belonging to an
            // earlier capture after the user has navigated to another thumbnail.
            this.thisCapturedScopeImagePath = null;
            this.AttachCapturedImageButton.IsVisible = false;

            if (this.thisTemporaryCapturedScopeBitmap != null)
            {
                this.thisTemporaryCapturedScopeBitmap.Dispose();
                this.thisTemporaryCapturedScopeBitmap = null;
            }
        }

        // ###########################################################################################
        // Shows a full-image overlay so the user can see that a new oscilloscope capture is in progress.
        // ###########################################################################################
        private void ShowFetchScopeImageOverlay()
        {
            this.FetchScopeImageOverlayBorder.IsVisible = true;
        }

        // ###########################################################################################
        // Hides the temporary full-image overlay once the oscilloscope capture has completed or failed.
        // ###########################################################################################
        private void HideFetchScopeImageOverlay()
        {
            this.FetchScopeImageOverlayBorder.IsVisible = false;
        }

        // ###########################################################################################
        // Returns true when popup-driven oscilloscope commands are allowed for the current session.
        // ###########################################################################################
        private bool CanSendOscilloscopeCommands()
        {
            return this.SyncOscilloscopeCheckBox.IsEnabled &&
                   this.SyncOscilloscopeCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Saves the oscilloscope sync switch state and refreshes the current popup-driven sync request.
        // ###########################################################################################
        private void OnSyncOscilloscopeCheckBoxChanged(object? sender, RoutedEventArgs e)
        {
            bool isEnabled = this.SyncOscilloscopeCheckBox.IsChecked == true;
            UserSettings.ComponentInfoOscilloscopeSyncEnabled = isEnabled;

            if (this.Owner is not Main mainOwner)
            {
                return;
            }

            if (!this.CanSendOscilloscopeCommands())
            {
                mainOwner.TabOscilloscopeControl.QueueComponentImageOscilloscopeSync(null);
                return;
            }

            this.ScheduleSelectedOscilloscopeImageSync(this.ThumbnailList.SelectedItem as ComponentImageItem);
        }

    }
}