using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        public string Note { get; set; } = string.Empty;
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
        private bool _hasExplicitRegionComponents = false;

        // Image matrix for zoom and pan capabilities
        private Matrix _imageMatrix = Matrix.Identity;
        private bool _isPanningImage = false;
        private Point _panStartPoint;
        private Matrix _panStartMatrix;


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

            // Seed the normal-size tracker and restore saved window size, splitters and state
            this._normalWidth = UserSettings.HasComponentInfoWindowLayout
                ? UserSettings.ComponentInfoWindowWidth
                : 680.0;
            this._normalHeight = UserSettings.HasComponentInfoWindowLayout
                ? UserSettings.ComponentInfoWindowHeight
                : 420.0;

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

            // Restore scroll action state to layout mapping
            if (string.Equals(UserSettings.ComponentInfoScrollAction, "Image zoom", StringComparison.OrdinalIgnoreCase))
                this.ScrollActionCombo.SelectedIndex = 1;
            else
                this.ScrollActionCombo.SelectedIndex = 0;

            this.ScrollActionCombo.SelectionChanged += this.OnScrollActionComboSelectionChanged;
            this.ThumbnailList.SelectionChanged += this.OnThumbnailSelectionChanged;

            // Map the interactions to the expanded top-panel boundaries area instead of the local box
            this.MainImageClickArea.PointerPressed += this.OnMainImageClickAreaPointerPressed;
            this.MainImageClickArea.PointerMoved += this.OnMainImageClickAreaPointerMoved;
            this.MainImageClickArea.PointerReleased += this.OnMainImageClickAreaPointerReleased;

            this.LocalFilesList.SelectionChanged += this.OnLocalFilesSelectionChanged;
            this.LinksList.SelectionChanged += this.OnLinksSelectionChanged;

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

                UserSettings.SaveComponentInfoWindowLayout(state, this._normalWidth, this._normalHeight, leftRatio, thumbHeight);
            };

            this.Closed += (_, _) =>
            {
                this._loadCts?.Cancel();
                this._pinBufferCts?.Cancel();
                this._pinFlashCts?.Cancel();
                foreach (var bmp in this._loadedBitmaps)
                    bmp.Dispose();
                this._loadedBitmaps.Clear();
            };
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

            if (e.Key == Key.Space || e.Key == Key.Enter)
            {
                this.ThumbnailList.SelectedIndex = 0;
                this.ThumbnailList.ScrollIntoView(this.ThumbnailList.SelectedItem!);
                e.Handled = true;
                return;
            }

            // Digit keys (top row and numpad): accumulate into a pin number and navigate to the matching image
            int digitValue = -1;
            if (e.Key >= Key.D0 && e.Key <= Key.D9)
                digitValue = (int)e.Key - (int)Key.D0;
            else if (e.Key >= Key.NumPad0 && e.Key <= Key.NumPad9)
                digitValue = (int)e.Key - (int)Key.NumPad0;

            if (digitValue >= 0)
            {
                this.HandlePinDigit((char)('0' + digitValue));
                e.Handled = true;
            }
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
                double delta = e.Delta.Y > 0 ? 1.2 : 0.8333333333333334;

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
        // Saves off the scroll action dropdown to global standard user settings configurations
        // ###########################################################################################
        private void OnScrollActionComboSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.ScrollActionCombo.SelectedItem is ComboBoxItem item && item.Content is string action)
            {
                UserSettings.ComponentInfoScrollAction = action;
            }
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
            this.LocalFilesList.ItemsSource = hasLocalFiles
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
            this.LinksList.ItemsSource = hasLinks
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
        // ###########################################################################################
        private void OnThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this._suppressThumbnailSelection)
                return;

            this.ResetImageZoom(); // Wipes previous zoom data clean every time a completely new index overrides selection
            var selected = this.ThumbnailList.SelectedItem as ComponentImageItem;
            this.MainComponentImage.Source = selected?.ImageSource;
            this.NoImageText.IsVisible = selected?.ImageSource == null;
            this.UpdateInfoOverlay();
            this.UpdateImageCounter();
            this.UpdateImageNote(selected);
        }

        // ###########################################################################################
        // Opens the selected local file in the OS default application, then clears the selection.
        // ###########################################################################################
        private void OnLocalFilesSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.LocalFilesList.SelectedItem is not ComponentLocalFileItem item)
                return;

            this.LocalFilesList.SelectedIndex = -1;

            if (string.IsNullOrWhiteSpace(item.FullPath))
                return;

            try
            {
                Process.Start(new ProcessStartInfo(item.FullPath) { UseShellExecute = true });
            }
            catch { }
        }

        // ###########################################################################################
        // Opens the selected link URL in the OS default browser, then clears the selection.
        // ###########################################################################################
        private void OnLinksSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.LinksList.SelectedItem is not ComponentLinkItem item)
                return;

            this.LinksList.SelectedIndex = -1;

            if (string.IsNullOrWhiteSpace(item.Url))
                return;

            try
            {
                Process.Start(new ProcessStartInfo(item.Url) { UseShellExecute = true });
            }
            catch { }
        }

        // ###########################################################################################
        // Refreshes the three top-left info labels from the currently selected thumbnail item.
        // ###########################################################################################
        private void UpdateInfoOverlay()
        {
            var selected = this.ThumbnailList.SelectedItem as ComponentImageItem;

            string? pin = selected?.Label;
            string? name = selected?.Name;

            // Suppress the pin/name label when it is identical to the name label
            bool pinSameAsName = !string.IsNullOrWhiteSpace(pin) &&
                                 string.Equals(pin, name, StringComparison.OrdinalIgnoreCase);

            this.SetInfoLabel(this.InfoPinBorder, this.InfoPinText, pinSameAsName ? null : pin);
            this.SetInfoLabel(this.InfoNameBorder, this.InfoNameText, name);
            this.SetInfoLabel(this.InfoOscBorder, this.InfoOscText, selected?.ExpectedOscilloscopeReading);
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
        private void SetInfoLabel(Border border, TextBlock textBlock, string? value)
        {
            bool show = !string.IsNullOrWhiteSpace(value);
            border.IsVisible = show;
            if (show)
                textBlock.Text = value;
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
        // ###########################################################################################
        private async void LoadImagesAsync(List<ComponentImageEntry> entries, string dataRoot, string? preservePin = null)
        {
            this._loadCts?.Cancel();
            this._loadCts = new CancellationTokenSource();
            var cts = this._loadCts;

            if (entries.Count == 0)
            {
                this.DisposeLoadedBitmaps();
                return;
            }

            var loaded = await Task.Run(() =>
            {
                var result = new List<(ComponentImageEntry Entry, Bitmap? Bitmap)>();

                foreach (var entry in entries)
                {
                    if (cts.Token.IsCancellationRequested)
                        break;

                    if (string.IsNullOrWhiteSpace(entry.File))
                    {
                        result.Add((entry, null));
                        continue;
                    }

                    var fullPath = Path.Combine(dataRoot, entry.File.Replace('/', Path.DirectorySeparatorChar));
                    Bitmap? bitmap = null;

                    if (File.Exists(fullPath))
                    {
                        try { bitmap = new Bitmap(fullPath); }
                        catch { }
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
                    Label = BuildImageLabel(x.Entry),
                    Pin = x.Entry.Pin.Trim(),
                    Name = x.Entry.Name,
                    ExpectedOscilloscopeReading = x.Entry.ExpectedOscilloscopeReading,
                    Note = x.Entry.Note
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

            // Dispose the previous bitmaps only after the UI has fully transitioned to the new set
            foreach (var bmp in oldBitmaps)
                bmp.Dispose();
        }

        // ###########################################################################################
        // Builds the thumbnail overlay label: "Pin X" when a pin number exists, otherwise the name.
        // ###########################################################################################
        private static string BuildImageLabel(ComponentImageEntry entry)
        {
            if (!string.IsNullOrWhiteSpace(entry.Pin))
                return $"Pin {entry.Pin.Trim()}";

            if (!string.IsNullOrWhiteSpace(entry.Name))
                return entry.Name.Trim();

            return string.Empty;
        }

        // ###########################################################################################
        // Clears UI image references, resets the image note section, and disposes loaded bitmaps.
        // ###########################################################################################
        private void DisposeLoadedBitmaps()
        {
            this.MainComponentImage.Source = null;
            this.ThumbnailList.ItemsSource = null;
            this.ImageNoteSection.IsVisible = false;
            foreach (var bmp in this._loadedBitmaps)
                bmp.Dispose();
            this._loadedBitmaps.Clear();
        }

        // ###########################################################################################
        // Returns true when an image is visible for the requested region.
        // Empty image regions are treated as shared and count for both PAL and NTSC.
        // ###########################################################################################
        private static bool IsImageVisibleInRegion(ComponentImageEntry image, string region)
        {
            return string.IsNullOrWhiteSpace(image.Region) ||
                   string.Equals(image.Region.Trim(), region, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Counts how many images belong to the current component for the requested region.
        // Empty image regions are included in both counters.
        // ###########################################################################################
        private int CountImagesForRegion(string region)
        {
            return this._allComponentImages.Count(img =>
                string.Equals(img.BoardLabel, this._boardLabel, StringComparison.OrdinalIgnoreCase) &&
                IsImageVisibleInRegion(img, region));
        }

        // ###########################################################################################
        // Updates the PAL and NTSC button captions with per-region image counters.
        // Empty image regions are included in both counters.
        // ###########################################################################################
        private void UpdateRegionButtonCounters()
        {
            int palCount = this.CountImagesForRegion("PAL");
            int ntscCount = this.CountImagesForRegion("NTSC");

            this.PalRegionButton.Content = $"PAL ({palCount})";
            this.NtscRegionButton.Content = $"NTSC ({ntscCount})";
        }

        // ###########################################################################################
        // Re-filters the stored image list for the current local region and triggers an async reload.
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
                    IsImageVisibleInRegion(img, this._localRegion))
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
        // while keeping the scroll-action selector and Close button left-aligned and visible.
        // ###########################################################################################
        private void UpdateRegionButtonsState()
        {
            this._suppressRegionToggle = true;
            bool isNtsc = string.Equals(this._localRegion, "NTSC", StringComparison.OrdinalIgnoreCase);

            this.PalRegionButton.IsVisible = this._hasExplicitRegionComponents;
            this.NtscRegionButton.IsVisible = this._hasExplicitRegionComponents;
            this.UpdateRegionButtonCounters();

            if (this.PalRegionButton.Parent is Grid footerGrid && footerGrid.ColumnDefinitions.Count >= 7)
            {
                footerGrid.ColumnDefinitions[0].Width = this._hasExplicitRegionComponents
                    ? GridLength.Auto
                    : new GridLength(0, GridUnitType.Pixel);

                footerGrid.ColumnDefinitions[1].Width = this._hasExplicitRegionComponents
                    ? new GridLength(4, GridUnitType.Pixel)
                    : new GridLength(0, GridUnitType.Pixel);

                footerGrid.ColumnDefinitions[2].Width = this._hasExplicitRegionComponents
                    ? GridLength.Auto
                    : new GridLength(0, GridUnitType.Pixel);

                footerGrid.ColumnDefinitions[3].Width = this._hasExplicitRegionComponents
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

            this.InfoRegionText.Bind(
                TextBlock.ForegroundProperty,
                this.GetResourceObservable($"{colorPrefix}_Fg"));
        }

        // ###########################################################################################
        // Selects the best-fit ComponentEntry for the given region:
        // exact region match → generic (empty region) → first available → null.
        // ###########################################################################################
        private ComponentEntry? PickComponentEntry(string region)
        {
            if (this._allComponentEntries.Count == 0)
                return null;

            var regionMatch = this._allComponentEntries.FirstOrDefault(e =>
                string.Equals(e.Region?.Trim(), region, StringComparison.OrdinalIgnoreCase));
            if (regionMatch != null)
                return regionMatch;

            var generic = this._allComponentEntries.FirstOrDefault(e =>
                string.IsNullOrWhiteSpace(e.Region));
            if (generic != null)
                return generic;

            return this._allComponentEntries[0];
        }

        // ###########################################################################################
        // Updates all region-sensitive text fields (title, category/part-number, description)
        // to reflect the current local region without affecting the global setting.
        // ###########################################################################################
        private void RefreshComponentText()
        {
            var entry = this.PickComponentEntry(this._localRegion);

            // Title: BoardLabel | FriendlyName | TechnicalNameOrValue (non-empty parts joined)
            var titleParts = new List<string>(3);
            if (!string.IsNullOrWhiteSpace(this._boardLabel))
                titleParts.Add(this._boardLabel.Trim());
            if (!string.IsNullOrWhiteSpace(entry?.FriendlyName))
                titleParts.Add(entry!.FriendlyName.Trim());
            if (!string.IsNullOrWhiteSpace(entry?.TechnicalNameOrValue))
                titleParts.Add(entry!.TechnicalNameOrValue.Trim());

            string titleText = titleParts.Count > 0
                ? string.Join(" | ", titleParts)
                : this._displayTextFallback;

            this.TitleText.Text = titleText;
            this.Title = titleText;

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
        }

        // ###########################################################################################
        // Closes the window when the Close button is clicked.
        // ###########################################################################################
        private void OnCloseButtonClick(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }

    }
}