using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Classic_Repair_Toolbox.TabSchematics;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using TabSchematics;

namespace CRT;

public partial class TabSchematics : UserControl
{
    public Main? MainWindow { get; set; }

    // Zoom
    internal Matrix schematicsMatrix = Matrix.Identity;

    // Thumbnails
    internal List<SchematicThumbnail> currentThumbnails = new(); // compliant with .NET6

    // Full-res viewer
    internal Bitmap? currentFullResBitmap;
    internal CancellationTokenSource? fullResLoadCts;

    // Panning
    private bool isPanning;
    private Point panStartPoint;
    private Matrix panStartMatrix;

    // Highlights
    internal Dictionary<string, HighlightSpatialIndex> highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);
    internal Dictionary<string, BoardSchematicEntry> schematicByName = new(StringComparer.OrdinalIgnoreCase);

    // Highlight rects per schematic per board label — built at board load, used for on-demand highlighting
    internal Dictionary<string, Dictionary<string, List<Rect>>> highlightRectsBySchematicAndLabel = new(StringComparer.OrdinalIgnoreCase);

    // Insert logic class declaration
    internal PolylineManagement? polylineManager;

    public TabSchematics()
    {
        InitializeComponent();
    }

    // ###########################################################################################
    // Handle manual row clicks for scaled label visibilities.
    // ###########################################################################################
    private void OnLabelBoardRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelBoard.IsChecked = !this.CheckLabelBoard.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelTechnicalRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelTechnical.IsChecked = !this.CheckLabelTechnical.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelFriendlyRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelFriendly.IsChecked = !this.CheckLabelFriendly.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelSelectedOnlyRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelSelectedOnly.IsChecked = !this.CheckLabelSelectedOnly.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Initializes the control by injecting the parent main window instance and wiring events.
    // ###########################################################################################
    public void Initialize(Main mainWindow)
    {
        this.MainWindow = mainWindow;

        // Restore initial states from User Settings
        this.CheckLabelBoard.IsChecked = UserSettings.SchematicsLabelBoard;
        this.CheckLabelTechnical.IsChecked = UserSettings.SchematicsLabelTechnical;
        this.CheckLabelFriendly.IsChecked = UserSettings.SchematicsLabelFriendly;
        this.CheckLabelSelectedOnly.IsChecked = UserSettings.SchematicsLabelSelectedOnly;

        bool isLabelsExpanded = UserSettings.SchematicsLabelsPanelExpanded;
        this.LabelsListPanel.IsVisible = isLabelsExpanded;
        this.ToggleLabelsPanelButton.Content = isLabelsExpanded ? "Collapse" : "Expand";

        this.CheckLabelBoard.IsCheckedChanged += (s, e) =>
        {
            UserSettings.SchematicsLabelBoard = this.CheckLabelBoard.IsChecked == true;
            this.UpdateComponentLabels();
        };
        this.CheckLabelTechnical.IsCheckedChanged += (s, e) =>
        {
            UserSettings.SchematicsLabelTechnical = this.CheckLabelTechnical.IsChecked == true;
            this.UpdateComponentLabels();
        };
        this.CheckLabelFriendly.IsCheckedChanged += (s, e) =>
        {
            UserSettings.SchematicsLabelFriendly = this.CheckLabelFriendly.IsChecked == true;
            this.UpdateComponentLabels();
        };
        this.CheckLabelSelectedOnly.IsCheckedChanged += (s, e) =>
        {
            UserSettings.SchematicsLabelSelectedOnly = this.CheckLabelSelectedOnly.IsChecked == true;
            this.UpdateComponentLabels();
        };

        this.ToggleLabelsPanelButton.Click += (s, e) =>
        {
            bool willBeExpanded = !this.LabelsListPanel.IsVisible;
            this.LabelsListPanel.IsVisible = willBeExpanded;
            this.ToggleLabelsPanelButton.Content = willBeExpanded ? "Collapse" : "Expand";
            UserSettings.SchematicsLabelsPanelExpanded = willBeExpanded;
        };

        this.polylineManager = new PolylineManagement(this, this.SchematicsPolylineCanvas);

        this.polylineManager.TraceStatsChanged += stats =>
        {
            Dispatcher.UIThread.Post(() => this.BuildTracesListPanel(stats));
        };

        // Saves active lines down dynamically over to disk
        this.polylineManager.TracesModified += () =>
        {
            var boardKey = this.MainWindow?.GetCurrentBoardKey();
            var schematicName = (this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail)?.Name;

            if (!string.IsNullOrEmpty(boardKey) && !string.IsNullOrEmpty(schematicName))
            {
                var export = this.polylineManager.ExportTraces();
                TraceStorage.SaveTraces(boardKey, schematicName, export);
            }
        };

        this.polylineManager.PaletteColorsChanged += colors =>
        {
            Dispatcher.UIThread.Post(() => this.RebuildDynamicPalette(colors));
        };
        this.RebuildDynamicPalette(this.polylineManager.PaletteColors);

        this.polylineManager.PaletteStateChanged += (visible, point) =>
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (visible)
                {
                    this.TraceFloatingPalette.Margin = new Avalonia.Thickness(point.X + 15, point.Y + 15, 0, 0);
                    this.TraceFloatingPalette.IsVisible = true;
                }
                else
                {
                    this.TraceFloatingPalette.IsVisible = false;
                }
            });
        };

        this.ClearTracesButton.Click += (s, e) =>
        {
            // Execute absolute clearing that triggers our new JSON saver event 
            this.polylineManager?.ClearAllTracesAndSave();
        };

        this.polylineManager.UndoStateChanged += hasUndo =>
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.UndoDeletedTraceButton.IsVisible = hasUndo;

                // Keep panel populated even when 0 lines remain if undo limit evaluates to valid stack lengths accurately 
                this.TracesPanel.IsVisible = (this.TracesListPanel.Children.Count > 0) || hasUndo;
            });
        };

        this.UndoDeletedTraceButton.Click += (s, e) =>
        {
            this.polylineManager?.UndoLastDeletion();
        };

        this.SchematicsSplitter.AddHandler(
            InputElement.PointerReleasedEvent,
            this.OnSchematicsSplitterPointerReleased,
            RoutingStrategies.Bubble,
            handledEventsToo: true);

        this.SchematicsImage.RenderTransformOrigin = RelativePoint.TopLeft;
        this.SchematicsHighlightsOverlay.RenderTransformOrigin = RelativePoint.TopLeft;

        this.SchematicsContainer.PropertyChanged += (s, e) =>
        {
            if (e.Property == Visual.BoundsProperty) this.ClampSchematicsMatrix();
        };

        this.SchematicsImage.PropertyChanged += (s, e) =>
        {
            if (e.Property == Visual.BoundsProperty) this.ClampSchematicsMatrix();
        };

        this.SchematicsThumbnailList.SelectionChanged += this.OnSchematicsThumbnailSelectionChanged;
        this.SchematicsContainer.PointerExited += this.OnSchematicsPointerExited;
    }

    private void OnTraceColorPickerPropertyChanged(object? sender, Avalonia.AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property.Name == "Color" && e.NewValue is Color c)
        {
            this.polylineManager?.ChangeActiveColor(c);
            this.polylineManager?.AddOrReplacePaletteColor(c);
            if (this.CustomColorButton != null)
            {
                this.CustomColorButton.Background = new SolidColorBrush(c);
            }
        }
    }

    // ###########################################################################################
    // Dynamically rebuilds Standard Ellipses mapped to the floating Palette context window.
    // ###########################################################################################
    private void RebuildDynamicPalette(List<Color> colors)
    {
        this.DynamicPaletteColorsPanel.Children.Clear();

        foreach (var c in colors)
        {
            var ellipse = new Avalonia.Controls.Shapes.Ellipse
            {
                Fill = new SolidColorBrush(c),
                Width = 18,
                Height = 18,
                Stroke = Brushes.White,
                StrokeThickness = 1,
                Cursor = new Cursor(StandardCursorType.Hand)
            };

            ellipse.PointerPressed += this.OnPaletteColorClicked;
            this.DynamicPaletteColorsPanel.Children.Add(ellipse);
        }
    }

    // ###########################################################################################
    // Generates the code-behind view layout tracking the amounts and visibility configs per line color.
    // ###########################################################################################
    private void BuildTracesListPanel(Dictionary<Color, int> stats)
    {
        this.TracesListPanel.Children.Clear();
        int totalCounts = 0;

        foreach (var kvp in stats)
        {
            Color colorItem = kvp.Key;
            int count = kvp.Value;
            totalCounts += count;

            var row = new StackPanel
            {
                Orientation = Avalonia.Layout.Orientation.Horizontal,
                Spacing = 6,
                Margin = new Avalonia.Thickness(0, 1),
                Background = Brushes.Transparent, // Ensures the empty space between items catches mouse clicks
                Cursor = new Cursor(StandardCursorType.Hand)
            };

            var cb = new CheckBox
            {
                MinHeight = 0,
                Margin = new Avalonia.Thickness(0),
                Padding = new Avalonia.Thickness(0),
                CornerRadius = new Avalonia.CornerRadius(2),
                IsChecked = this.polylineManager?.GetColorVisibility(colorItem) ?? true,
                IsHitTestVisible = false // Disable direct native hits so the row handles everything universally
            };

            cb.IsCheckedChanged += (s, e) =>
            {
                this.polylineManager?.SetVisibilityByColor(colorItem, cb.IsChecked == true);
            };

            // Wrapped CheckBox slightly larger, as the template natively has invisible touch padding 
            // that shrinks its visual square smaller than its layout bounds.
            var cbContainer = new Viewbox
            {
                Width = 20,
                Height = 20,
                Child = cb,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                IsHitTestVisible = false
            };

            var borderBrushCandidate = this.TryFindResource("AppThemeBorderBrush", out var res) ? res as IBrush : Brushes.Gray;

            var border = new Border
            {
                Width = 48,
                Height = 14,
                Background = new SolidColorBrush(colorItem),
                BorderBrush = borderBrushCandidate,
                BorderThickness = new Avalonia.Thickness(1),
                CornerRadius = new Avalonia.CornerRadius(2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };

            var textForegroundCandidate = this.TryFindResource("SchematicsComponent_Fg", out var textRes) ? textRes as IBrush : Brushes.Black;

            var txt = new TextBlock
            {
                Text = $"({count})",
                FontSize = 11,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center,
                Foreground = textForegroundCandidate
            };

            // Clicking anywhere on the row flips the active status of the checkbox
            row.PointerPressed += (s, e) =>
            {
                if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
                {
                    cb.IsChecked = !cb.IsChecked;
                    e.Handled = true;
                }
            };

            row.Children.Add(cbContainer);
            row.Children.Add(border);
            row.Children.Add(txt);

            this.TracesListPanel.Children.Add(row);
        }

        this.TracesPanel.IsVisible = totalCounts > 0;
    }

    // ###########################################################################################
    // Applies the locally clicked palette color onto the currently active trace line.
    // ###########################################################################################
    private void OnPaletteColorClicked(object? sender, PointerPressedEventArgs e)
    {
        if (sender is Avalonia.Controls.Shapes.Ellipse ellipse && ellipse.Fill is ISolidColorBrush brush)
        {
            this.polylineManager?.ChangeActiveColor(brush.Color);

            // Sync standard ColorPicker thumb AND button surface internally reflecting context accurately
            this.SetTraceColorPickerColor(brush.Color);
            this.CustomColorButton.Background = brush;
        }
        e.Handled = true;
    }

    // ###########################################################################################
    // Updates UI visibility for all contextual label borders at once.
    // ###########################################################################################
    public void HideLabels()
    {
        this.SchematicsNameBorder.IsVisible = false;
        this.SchematicsRegionBorder.IsVisible = false;
        this.SchematicsHoverLabelBorder.IsVisible = false;
    }

    // ###########################################################################################
    // Safely retrieves and sets the ColorView inside the unmapped flyout surface.
    // ###########################################################################################
    private void SetTraceColorPickerColor(Color color)
    {
        if (this.CustomColorButton.Flyout is Avalonia.Controls.Flyout flyout && flyout.Content != null)
        {
            var propInfo = flyout.Content.GetType().GetProperty("Color");
            propInfo?.SetValue(flyout.Content, color);
        }
    }

    // ###########################################################################################
    // Removes the currently targeted polyline trace securely via UI.
    // ###########################################################################################
    private void OnPaletteDeleteClicked(object? sender, PointerPressedEventArgs e)
    {
        this.polylineManager?.DeleteActivePolyline();
        e.Handled = true;
    }

    // ###########################################################################################
    // Handles mouse wheel zoom on the Schematics image, centered on the cursor position.
    // ###########################################################################################
    private void OnSchematicsZoom(object? sender, PointerWheelEventArgs e)
    {
        var pos = e.GetPosition(this.SchematicsImage);
        double delta = e.Delta.Y > 0 ? AppConfig.SchematicsZoomFactor : 1.0 / AppConfig.SchematicsZoomFactor;

        double newScale = this.schematicsMatrix.M11 * delta;

        if (newScale > AppConfig.SchematicsMaxZoom)
            return;

        if (newScale < AppConfig.SchematicsMinZoom)
        {
            this.schematicsMatrix = Matrix.Identity;
            ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
            ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;

            this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
            this.SchematicsHighlightsOverlay.InvalidateVisual();

            e.Handled = true;
            return;
        }

        // Build a zoom matrix centered at the cursor position in image-local space
        var zoomMatrix = Matrix.CreateTranslation(-pos.X, -pos.Y)
                       * Matrix.CreateScale(delta, delta)
                       * Matrix.CreateTranslation(pos.X, pos.Y);

        this.schematicsMatrix = zoomMatrix * this.schematicsMatrix;
        this.ClampSchematicsMatrix();

        e.Handled = true;
    }

    // ###########################################################################################
    // Handles right-click for panning on the schematic view and selection toggling on release.
    // Left-click selects hovered component, and single-click opens the component info popup.
    // Also routes pointer presses to the polyline manager if appropriate.
    // ###########################################################################################
    private void OnSchematicsPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);
        var pointer = e.GetCurrentPoint(this.SchematicsContainer);

        bool hoveringComponent = this.TryGetHoveredBoardLabel(point, out var boardLabel, out var displayText);

        if (TryInvert(this.schematicsMatrix, out var inv))
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerPressed(point, localPoint, pointer, hoveringComponent))
            {
                e.Handled = true;
                return;
            }
        }

        if (pointer.Properties.IsRightButtonPressed)
        {
            this.isPanning = true;
            this.panStartPoint = point;
            this.panStartMatrix = this.schematicsMatrix;

            // Call hide UI first to avoid overwriting the panning cursor 
            this.HideSchematicsHoverUi();
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);

            e.Pointer.Capture(this.SchematicsContainer);
            e.Handled = true;
            return;
        }

        if (pointer.Properties.IsLeftButtonPressed && hoveringComponent)
        {
            this.SelectComponentByBoardLabel(boardLabel);

            if (e.ClickCount == 1 && this.MainWindow != null)
                this.MainWindow.OpenComponentInfoPopup(boardLabel, displayText);

            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Translates the schematics image while the right mouse button is held down.
    // routes movement and shift key state to Polyline Manager.
    // ###########################################################################################
    private void OnSchematicsPointerMoved(object? sender, PointerEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);
        bool isShiftDown = e.KeyModifiers.HasFlag(KeyModifiers.Shift);

        if (TryInvert(this.schematicsMatrix, out var inv))
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerMoved(localPoint, isShiftDown))
            {
                e.Handled = true;
                return;
            }
        }

        if (this.isPanning)
        {
            var delta = point - this.panStartPoint;
            this.schematicsMatrix = this.panStartMatrix * Matrix.CreateTranslation(delta.X, delta.Y);
            this.ClampSchematicsMatrix();
            e.Handled = true;
            return;
        }

        this.UpdateSchematicsHoverUi(point);
    }

    // ###########################################################################################
    // Exits pan mode when the right mouse button is released, or finalized polyline logic.
    // Also evaluates if the release qualifies as a stationary right-click to toggle selection.
    // ###########################################################################################
    private void OnSchematicsPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        var point = e.GetPosition(this.SchematicsContainer);

        if (TryInvert(this.schematicsMatrix, out var inv))
        {
            var localPoint = new Point(
                (point.X * inv.M11) + (point.Y * inv.M21) + inv.M31,
                (point.X * inv.M12) + (point.Y * inv.M22) + inv.M32);

            if (this.polylineManager != null && this.polylineManager.OnPointerReleased(point, localPoint))
            {
                e.Handled = true;
                return;
            }
        }

        if (!this.isPanning)
            return;

        this.isPanning = false;
        e.Pointer.Capture(null);

        // Determine if movement was small enough to be interpreted as a right-click rather than a pan
        var delta = point - this.panStartPoint;
        if (Math.Abs(delta.X) < 4 && Math.Abs(delta.Y) < 4)
        {
            if (this.TryGetHoveredBoardLabel(point, out var boardLabel, out var displayText))
            {
                this.ToggleComponentSelectionByBoardLabel(boardLabel);
            }
        }

        this.UpdateSchematicsHoverUi(e.GetPosition(this.SchematicsContainer));
        e.Handled = true;
    }

    // ###########################################################################################
    // Updates the displayed region and schematic name overlays.
    // ###########################################################################################
    public void UpdateOverlayLabels()
    {
        var selected = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;

        string schematicName = selected?.Name ?? string.Empty;
        this.SchematicsNameLabel.Text = schematicName;
        this.SchematicsNameBorder.IsVisible = !string.IsNullOrWhiteSpace(schematicName);

        string rawRegion = UserSettings.Region?.Trim() ?? string.Empty;
        this.SchematicsRegionLabel.Text = string.IsNullOrWhiteSpace(rawRegion) ? "All Regions" : rawRegion;
        this.SchematicsRegionBorder.IsVisible = this.SchematicsNameBorder.IsVisible;

        string regionKey = rawRegion.ToUpperInvariant();

        string colorPrefix = regionKey switch
        {
            "PAL" => "Region_PAL",
            "NTSC" => "Region_NTSC",
            _ => "SchematicsRegion"
        };

        this.SchematicsRegionBorder.Bind(
            Border.BackgroundProperty,
            this.GetResourceObservable($"{colorPrefix}_Bg"));

        this.SchematicsRegionLabel.Bind(
            TextBlock.ForegroundProperty,
            this.GetResourceObservable($"{colorPrefix}_Fg"));
    }

    // ###########################################################################################
    // Loads the full-resolution image for the selected thumbnail and sets up the highlight overlay.
    // ###########################################################################################
    private async void OnSchematicsThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        this.UpdateOverlayLabels();

        this.fullResLoadCts?.Cancel();
        this.fullResLoadCts = new CancellationTokenSource();
        var cts = this.fullResLoadCts;

        var selected = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;

        this.SchematicsImage.Source = null;
        this.SchematicsMissingImageText.IsVisible = false; // Hide while loading
        this.schematicsMatrix = Matrix.Identity;
        ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;

        this.SchematicsHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        if (selected == null || string.IsNullOrEmpty(selected.ImageFilePath))
            return;

        // Save the newly selected schematic for this board
        var boardKey = this.MainWindow?.GetCurrentBoardKey();
        if (!string.IsNullOrEmpty(boardKey))
        {
            UserSettings.SetLastSchematicForBoard(boardKey, selected.Name);
        }

        var bitmap = await Task.Run(() =>
        {
            if (cts.Token.IsCancellationRequested) return null;

            try { return new Bitmap(selected.ImageFilePath); }
            catch (Exception ex)
            {
                Logger.Warning($"Cannot load image file [{selected.ImageFilePath}] - [{ex.Message}]");
                return null;
            }
        }, cts.Token);

        if (cts.Token.IsCancellationRequested)
        {
            bitmap?.Dispose();
            return;
        }

        this.currentFullResBitmap?.Dispose();
        this.currentFullResBitmap = bitmap;
        this.SchematicsImage.Source = bitmap;

        if (bitmap != null)
        {
            this.SchematicsMissingImageText.IsVisible = false;

            // Always set BitmapPixelSize so the overlay can render as soon as a component is selected,
            // even if no highlight index exists yet at the time this schematic loads.
            this.SchematicsHighlightsOverlay.BitmapPixelSize = bitmap.PixelSize;

            if (this.highlightIndexBySchematic.TryGetValue(selected.Name, out var index) &&
                this.schematicByName.TryGetValue(selected.Name, out var schematic))
            {
                this.SchematicsHighlightsOverlay.HighlightIndex = index;
                this.SchematicsHighlightsOverlay.HighlightColor = ParseColorOrDefault(schematic.MainImageHighlightColor, Colors.IndianRed);
                this.SchematicsHighlightsOverlay.HighlightOpacity = ParseOpacityOrDefault(schematic.MainHighlightOpacity, 0.20);
            }
        }
        else
        {
            this.SchematicsMissingImageText.IsVisible = true;
        }

        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHighlightsOverlay.InvalidateVisual();

        // Populate trace database for this exact board/schematic setup immediately before render logic finishes 
        if (this.polylineManager != null && !string.IsNullOrEmpty(boardKey))
        {
            var loaded = TraceStorage.GetTraces(boardKey, selected.Name);
            this.polylineManager.ImportTraces(loaded);
        }

        // Defer a clamp call so the engine can measure and center the new image layout 
        // immediately instead of waiting for a window resize or banner collapse.
        Dispatcher.UIThread.Post(() =>
        {
            this.ClampSchematicsMatrix();
            this.UpdateComponentLabels();
        });

        // Defer a clamp call so the engine can measure and center the new image layout 
        // immediately instead of waiting for a window resize or banner collapse.
        Dispatcher.UIThread.Post(() => this.ClampSchematicsMatrix());
    }

    // ###########################################################################################
    // Saves the schematics/thumbnail split ratio for the current board after the drag ends.
    // Deferred via Post to ensure Bounds reflects the completed layout pass.
    // ###########################################################################################
    private void OnSchematicsSplitterPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        var boardKey = this.MainWindow?.GetCurrentBoardKey();
        if (string.IsNullOrEmpty(boardKey))
        {
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            var leftWidth = this.SchematicsContainer.Bounds.Width;
            var rightWidth = this.SchematicsThumbnailList.Bounds.Width;
            var total = leftWidth + rightWidth;
            if (total <= 0)
            {
                return;
            }
            UserSettings.SetSchematicsSplitterRatio(boardKey, leftWidth / total);
        });
    }

    // ###########################################################################################
    // Clears the main schematics image and resets the zoom and highlight overlay state.
    // ###########################################################################################
    public void ResetSchematicsViewer()
    {
        this.polylineManager?.Reset(); // Enforce state wipe on board reload
        this.SchematicsLabelsCanvas.Children.Clear();

        // Reset color indicator fallback
        this.SetTraceColorPickerColor(Colors.White);
        this.CustomColorButton.Background = Brushes.White;

        this.fullResLoadCts?.Cancel();
        this.fullResLoadCts = null;

        this.currentFullResBitmap?.Dispose();
        this.currentFullResBitmap = null;

        this.SchematicsNameBorder.IsVisible = false;
        this.SchematicsRegionBorder.IsVisible = false;

        this.SchematicsImage.Source = null;
        this.SchematicsMissingImageText.IsVisible = false;

        this.schematicsMatrix = Matrix.Identity;
        ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;

        this.SchematicsHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        this.isPanning = false;
        this.HideSchematicsHoverUi();
    }

    // ###########################################################################################
    // Returns the rectangle (in the image control's local coordinate space) that the actual
    // bitmap content occupies, accounting for Stretch="Uniform" letterboxing on either axis.
    // ###########################################################################################
    internal Rect GetImageContentRect()
    {
        var imageSize = this.SchematicsImage.Bounds.Size;
        var bitmap = this.currentFullResBitmap;

        if (bitmap == null || imageSize.Width <= 0 || imageSize.Height <= 0)
            return new Rect(imageSize);

        double containerAspect = imageSize.Width / imageSize.Height;
        double bitmapAspect = bitmap.Size.Width / bitmap.Size.Height;

        double contentX, contentY, contentWidth, contentHeight;

        if (bitmapAspect > containerAspect)
        {
            contentWidth = imageSize.Width;
            contentHeight = imageSize.Width / bitmapAspect;
            contentX = 0;
            contentY = (imageSize.Height - contentHeight) / 2.0;
        }
        else
        {
            contentHeight = imageSize.Height;
            contentWidth = imageSize.Height * bitmapAspect;
            contentX = (imageSize.Width - contentWidth) / 2.0;
            contentY = 0;
        }

        return new Rect(contentX, contentY, contentWidth, contentHeight);
    }

    // ###########################################################################################
    // Clamps the current schematics matrix so no empty space is visible inside the container.
    // ###########################################################################################
    private void ClampSchematicsMatrix()
    {
        var containerSize = this.SchematicsContainer.Bounds.Size;
        if (containerSize.Width <= 0 || containerSize.Height <= 0)
            return;

        var contentRect = this.GetImageContentRect();
        double scale = this.schematicsMatrix.M11;
        double tx = this.schematicsMatrix.M31;
        double ty = this.schematicsMatrix.M32;

        var transformedRect = contentRect.TransformToAABB(this.schematicsMatrix);

        double scaledWidth = transformedRect.Width;
        double scaledHeight = transformedRect.Height;
        double scaledLeft = transformedRect.Left;
        double scaledTop = transformedRect.Top;
        double scaledRight = transformedRect.Right;
        double scaledBottom = transformedRect.Bottom;

        if (scaledWidth >= containerSize.Width)
        {
            if (scaledLeft > 0) tx -= scaledLeft;
            else if (scaledRight < containerSize.Width) tx += containerSize.Width - scaledRight;
        }
        else
        {
            tx = (containerSize.Width - scaledWidth) / 2.0 - scale * contentRect.Left;
        }

        if (scaledHeight >= containerSize.Height)
        {
            if (scaledTop > 0) ty -= scaledTop;
            else if (scaledBottom < containerSize.Height) ty += containerSize.Height - scaledBottom;
        }
        else
        {
            ty = -(scale * contentRect.Top);
        }

        this.schematicsMatrix = new Matrix(scale, 0, 0, scale, tx, ty);
        ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsPolylineCanvas.RenderTransform!).Matrix = this.schematicsMatrix; // Apply to Polyline layer
        ((MatrixTransform)this.SchematicsLabelsCanvas.RenderTransform!).Matrix = this.schematicsMatrix; // Apply to Labels layer

        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHighlightsOverlay.InvalidateVisual();

        this.polylineManager?.UpdateScaleFactor(scale); // Ensure line visuals behave consistently
        this.UpdateComponentLabelsScale(scale);
    }

    // ###########################################################################################
    // Applies inverse scale to mapped labels so they remain standard text size regardless of zoom.
    // ###########################################################################################
    private void UpdateComponentLabelsScale(double scale)
    {
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;
        foreach (var child in this.SchematicsLabelsCanvas.Children)
        {
            if (child is Border container && container.RenderTransform is TransformGroup group)
            {
                var st = group.Children.OfType<ScaleTransform>().FirstOrDefault();
                if (st != null)
                {
                    st.ScaleX = inverseScale;
                    st.ScaleY = inverseScale;
                }
            }
        }
    }

    // ###########################################################################################
    // Generates floating labels exactly above relevant components based on checkbox matrix.
    // ###########################################################################################
    public void UpdateComponentLabels()
    {
        this.SchematicsLabelsCanvas.Children.Clear();

        if (this.CheckLabelBoard.IsChecked != true &&
            this.CheckLabelTechnical.IsChecked != true &&
            this.CheckLabelFriendly.IsChecked != true)
        {
            return;
        }

        if (this.MainWindow == null || this.currentFullResBitmap == null)
            return;

        double imgWidth = this.currentFullResBitmap.PixelSize.Width;
        double imgHeight = this.currentFullResBitmap.PixelSize.Height;
        if (imgWidth <= 0 || imgHeight <= 0)
            return;

        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb == null)
            return;

        if (!this.highlightRectsBySchematicAndLabel.TryGetValue(selectedThumb.Name, out var byLabel))
            return;

        var visibleItems = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<Main.ComponentListItem>().ToList() ?? new List<Main.ComponentListItem>();
        var selectedItems = this.MainWindow.ComponentFilterListBox.SelectedItems?.Cast<Main.ComponentListItem>().ToList() ?? new List<Main.ComponentListItem>();

        bool selectedOnly = this.CheckLabelSelectedOnly.IsChecked == true;
        var itemsToLoop = selectedOnly ? selectedItems : visibleItems;
        var seenLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var contentRect = this.GetImageContentRect();

        var borderBrushCandidate = this.TryFindResource("AppThemeBorderBrush", out var res) ? res as IBrush : Brushes.Gray;
        var textForegroundCandidate = this.TryFindResource("SchematicsComponent_Fg", out var textRes) ? textRes as IBrush : Brushes.White;
        var textBackgroundCandidate = this.TryFindResource("SchematicsComponent_Bg", out var bgRes) ? bgRes as IBrush : new SolidColorBrush(Color.Parse("#CC000000"));

        double currentScale = this.schematicsMatrix.M11;
        double inverseScale = currentScale > 0 ? 1.0 / currentScale : 1.0;

        foreach (var item in itemsToLoop)
        {
            if (string.IsNullOrWhiteSpace(item.BoardLabel)) continue;
            if (!seenLabels.Add(item.BoardLabel)) continue;

            if (!byLabel.TryGetValue(item.BoardLabel, out var rects) || rects.Count == 0) continue;

            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;
            foreach (var r in rects)
            {
                if (r.X < minX) minX = r.X;
                if (r.Y < minY) minY = r.Y;
                if (r.Right > maxX) maxX = r.Right;
                if (r.Bottom > maxY) maxY = r.Bottom;
            }

            double centerX = minX + (maxX - minX) / 2.0;
            double centerY = minY + (maxY - minY) / 2.0;

            double localX = contentRect.X + (centerX / imgWidth) * contentRect.Width;
            double localY = contentRect.Y + (centerY / imgHeight) * contentRect.Height;

            var parts = item.SelectionKey?.Split('\u001F') ?? Array.Empty<string>();
            string friendlyName = parts.Length > 1 ? parts[1] : string.Empty;
            string technicalName = parts.Length > 2 ? parts[2] : string.Empty;

            var lines = new List<string>();
            if (this.CheckLabelBoard.IsChecked == true && !string.IsNullOrWhiteSpace(item.BoardLabel)) lines.Add(item.BoardLabel);
            if (this.CheckLabelTechnical.IsChecked == true && !string.IsNullOrWhiteSpace(technicalName)) lines.Add(technicalName);
            if (this.CheckLabelFriendly.IsChecked == true && !string.IsNullOrWhiteSpace(friendlyName)) lines.Add(friendlyName);

            if (lines.Count == 0) continue;

            var tb = new TextBlock
            {
                Text = string.Join("\n", lines),
                FontSize = 11,
                Foreground = textForegroundCandidate,
                FontWeight = FontWeight.Bold,
                TextAlignment = TextAlignment.Center
            };

            var innerBorder = new Border
            {
                Background = textBackgroundCandidate,
                BorderBrush = borderBrushCandidate,
                BorderThickness = new Avalonia.Thickness(1),
                CornerRadius = new Avalonia.CornerRadius(4),
                Padding = new Avalonia.Thickness(6, 4),
                Child = tb
            };

            var transformGroup = new TransformGroup();
            transformGroup.Children.Add(new ScaleTransform(inverseScale, inverseScale));

            var container = new Border
            {
                IsHitTestVisible = false,
                Child = innerBorder,
                RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative),
                RenderTransform = transformGroup
            };

            // Resolves exact dynamic layout dimensions to physically center everything exactly in the component
            container.SizeChanged += (s, ev) =>
            {
                Canvas.SetLeft(container, localX - (ev.NewSize.Width / 2.0));
                Canvas.SetTop(container, localY - (ev.NewSize.Height / 2.0));
            };

            Canvas.SetLeft(container, localX);
            Canvas.SetTop(container, localY);

            this.SchematicsLabelsCanvas.Children.Add(container);
        }
    }

    // ###########################################################################################
    // Creates a pre-scaled bitmap from a full-resolution source image.
    // ###########################################################################################
    public static RenderTargetBitmap CreateScaledThumbnail(Bitmap source, int maxWidth)
    {
        double scale = Math.Min(1.0, (double)maxWidth / source.PixelSize.Width);
        int tw = Math.Max(1, (int)(source.PixelSize.Width * scale));
        int th = Math.Max(1, (int)(source.PixelSize.Height * scale));

        var imageControl = new Image { Source = source, Stretch = Stretch.Uniform };
        imageControl.Measure(new Size(tw, th));
        imageControl.Arrange(new Rect(0, 0, tw, th));

        var rtb = new RenderTargetBitmap(new PixelSize(tw, th), new Vector(96, 96));
        rtb.Render(imageControl);
        return rtb;
    }

    // ###########################################################################################
    // Composites highlight rectangles onto a base thumbnail and returns the new rendered bitmap.
    // ###########################################################################################
    public static RenderTargetBitmap CreateHighlightedThumbnail(
        IImage baseThumbnail, PixelSize originalPixelSize,
        HighlightSpatialIndex index, BoardSchematicEntry schematic, double opacityMultiplier = 1.0)
    {
        int tw = 1, th = 1;
        if (baseThumbnail is RenderTargetBitmap rtb)
        {
            tw = rtb.PixelSize.Width;
            th = rtb.PixelSize.Height;
        }
        else if (baseThumbnail is Bitmap bmp)
        {
            tw = bmp.PixelSize.Width;
            th = bmp.PixelSize.Height;
        }

        var root = new Grid();
        var image = new Image
        {
            Source = baseThumbnail,
            Stretch = Stretch.Uniform,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
            VerticalAlignment = Avalonia.Layout.VerticalAlignment.Top
        };

        var overlay = new SchematicHighlightsOverlay
        {
            HighlightIndex = index,
            BitmapPixelSize = originalPixelSize,
            ViewMatrix = Matrix.Identity,
            HighlightColor = ParseColorOrDefault(schematic.ThumbnailImageHighlightColor, Colors.IndianRed),
            HighlightOpacity = ParseOpacityOrDefault(schematic.ThumbnailHighlightOpacity, 0.20) * Math.Clamp(opacityMultiplier, 0.0, 1.0),
            IsHitTestVisible = false
        };

        root.Children.Add(image);
        root.Children.Add(overlay);

        root.Measure(new Size(tw, th));
        root.Arrange(new Rect(0, 0, tw, th));

        var result = new RenderTargetBitmap(new PixelSize(tw, th), new Vector(96, 96));
        result.Render(root);
        return result;
    }

    // ###########################################################################################
    // Rebuilds highlight indices from the selected board labels, then applies highlight visuals
    // to the main schematic and all thumbnails.
    // ###########################################################################################
    public void UpdateHighlightsForComponents(List<string> boardLabels)
    {
        this.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

        if (boardLabels.Count > 0)
        {
            foreach (var (schematicName, byLabel) in this.highlightRectsBySchematicAndLabel)
            {
                var rects = new List<Rect>();
                foreach (var label in boardLabels)
                {
                    if (byLabel.TryGetValue(label, out var labelRects))
                        rects.AddRange(labelRects);
                }

                if (rects.Count > 0)
                    this.highlightIndexBySchematic[schematicName] = new HighlightSpatialIndex(rects);
            }
        }

        bool hasSelection = boardLabels.Count > 0;
        if (this.MainWindow != null)
        {
            this.ApplyHighlightVisuals(hasSelection, this.MainWindow.GetCurrentBlinkFactor(hasSelection));
            this.MainWindow.UpdateBlinkTimer(hasSelection);
        }

        this.UpdateComponentLabels();
    }

    // ###########################################################################################
    // Applies current highlight visuals (including blink phase) to main schematic and thumbnails.
    // ###########################################################################################
    public void ApplyHighlightVisuals(bool hasSelection, double blinkFactor)
    {
        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb != null &&
            this.highlightIndexBySchematic.TryGetValue(selectedThumb.Name, out var mainIndex) &&
            this.schematicByName.TryGetValue(selectedThumb.Name, out var mainSchematic))
        {
            this.SchematicsHighlightsOverlay.HighlightIndex = mainIndex;
            this.SchematicsHighlightsOverlay.BitmapPixelSize = this.currentFullResBitmap?.PixelSize ?? new PixelSize(0, 0);
            this.SchematicsHighlightsOverlay.HighlightColor = ParseColorOrDefault(mainSchematic.MainImageHighlightColor, Colors.IndianRed);
            this.SchematicsHighlightsOverlay.HighlightOpacity = ParseOpacityOrDefault(mainSchematic.MainHighlightOpacity, 0.20) * blinkFactor;
        }
        else
        {
            this.SchematicsHighlightsOverlay.HighlightIndex = null;
        }

        this.SchematicsHighlightsOverlay.InvalidateVisual();

        foreach (var thumb in this.currentThumbnails)
        {
            if (thumb.BaseThumbnail == null)
                continue;

            bool hasMatch = false;

            if (this.highlightIndexBySchematic.TryGetValue(thumb.Name, out var thumbIndex) &&
                this.schematicByName.TryGetValue(thumb.Name, out var thumbSchematic))
            {
                hasMatch = true;
                var highlighted = CreateHighlightedThumbnail(thumb.BaseThumbnail, thumb.OriginalPixelSize, thumbIndex, thumbSchematic, blinkFactor);
                var old = thumb.ImageSource;
                thumb.ImageSource = highlighted;
                if (!ReferenceEquals(old, thumb.BaseThumbnail))
                    (old as IDisposable)?.Dispose();
            }
            else
            {
                if (!ReferenceEquals(thumb.ImageSource, thumb.BaseThumbnail))
                {
                    var old = thumb.ImageSource;
                    thumb.ImageSource = thumb.BaseThumbnail;
                    (old as IDisposable)?.Dispose();
                }
            }

            bool isRelevantForDimming = !hasSelection || hasMatch;
            thumb.VisualOpacity = isRelevantForDimming ? 1.0 : 0.35;
            thumb.IsMatchForSelection = hasSelection && hasMatch;
        }
    }

    // ###########################################################################################
    // Clears hover label and resets schematic cursor.
    // ###########################################################################################
    private void HideSchematicsHoverUi()
    {
        this.SchematicsHoverLabelBorder.IsVisible = false;
        this.SchematicsHoverLabelText.Text = string.Empty;
        this.SchematicsContainer.Cursor = Cursor.Default;

        if (this.MainWindow != null)
            this.MainWindow.isHoveringComponent = false;
    }

    // ###########################################################################################
    // Clears hover UI when pointer exits schematic area.
    // ###########################################################################################
    private void OnSchematicsPointerExited(object? sender, PointerEventArgs e)
    {
        if (this.isPanning)
            return;

        this.HideSchematicsHoverUi();
    }

    // ###########################################################################################
    // Updates hover label/cursor from current pointer position.
    // ###########################################################################################
    private void UpdateSchematicsHoverUi(Point pointerInContainer)
    {
        if (this.TryGetHoveredBoardLabel(pointerInContainer, out _, out var displayText))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Hand);
            this.SchematicsHoverLabelText.Text = displayText;
            this.SchematicsHoverLabelBorder.IsVisible = true;
            if (this.MainWindow != null) this.MainWindow.isHoveringComponent = true;
            return;
        }

        if (this.MainWindow != null) this.MainWindow.isHoveringComponent = false;
        this.HideSchematicsHoverUi();
    }

    // ###########################################################################################
    // Resolves hovered board label and the exact text shown in component selector.
    // Includes components that are visible in the selector even when not selected/highlighted.
    // ###########################################################################################
    private bool TryGetHoveredBoardLabel(Point pointerInContainer, out string boardLabel, out string displayText)
    {
        boardLabel = string.Empty;
        displayText = string.Empty;

        if (this.currentFullResBitmap == null || this.MainWindow == null)
            return false;

        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb == null)
            return false;

        if (!this.highlightRectsBySchematicAndLabel.TryGetValue(selectedThumb.Name, out var byLabel))
            return false;

        if (!TryInvert(this.schematicsMatrix, out var inv))
            return false;

        var localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        var contentRect = this.GetImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0 || !contentRect.Contains(localPoint))
            return false;

        double px = ((localPoint.X - contentRect.X) / contentRect.Width) * this.currentFullResBitmap.PixelSize.Width;
        double py = ((localPoint.Y - contentRect.Y) / contentRect.Height) * this.currentFullResBitmap.PixelSize.Height;
        var pixelPoint = new Point(px, py);

        var visibleItems = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<Main.ComponentListItem>().ToList() ?? new List<Main.ComponentListItem>();
        var seenLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in visibleItems)
        {
            if (string.IsNullOrWhiteSpace(item.BoardLabel))
                continue;

            if (!seenLabels.Add(item.BoardLabel))
                continue;

            if (!byLabel.TryGetValue(item.BoardLabel, out var rects))
                continue;

            if (!rects.Any(r => r.Contains(pixelPoint)))
                continue;

            boardLabel = item.BoardLabel;
            displayText = item.DisplayText;
            return true;
        }

        return false;
    }

    // ###########################################################################################
    // Selects first component row matching board label and scrolls it into view.
    // ###########################################################################################
    private void SelectComponentByBoardLabel(string boardLabel)
    {
        if (this.MainWindow == null) return;
        var items = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<Main.ComponentListItem>().ToList() ?? new List<Main.ComponentListItem>();
        int index = items.FindIndex(i => string.Equals(i.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase));
        if (index < 0)
            return;

        this.MainWindow.ComponentFilterListBox.Selection.Select(index);
        this.MainWindow.ComponentFilterListBox.ScrollIntoView(items[index]);
    }

    // ###########################################################################################
    // Deselects all component rows that match the given board label.
    // ###########################################################################################
    private void DeselectComponentByBoardLabel(string boardLabel)
    {
        if (this.MainWindow == null) return;
        var items = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<Main.ComponentListItem>().ToList() ?? new List<Main.ComponentListItem>();
        if (items.Count == 0)
            return;

        for (int i = 0; i < items.Count; i++)
        {
            if (string.Equals(items[i].BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
                this.MainWindow.ComponentFilterListBox.Selection.Deselect(i);
        }
    }

    // ###########################################################################################
    // Toggles selection for a component board label.
    // ###########################################################################################
    private void ToggleComponentSelectionByBoardLabel(string boardLabel)
    {
        if (this.MainWindow == null) return;
        bool hasSelectedMatch = this.MainWindow.ComponentFilterListBox.SelectedItems?
            .Cast<Main.ComponentListItem>()
            .Any(i => string.Equals(i.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)) ?? false;

        if (hasSelectedMatch)
            this.DeselectComponentByBoardLabel(boardLabel);
        else
            this.SelectComponentByBoardLabel(boardLabel);
    }

    // ###########################################################################################
    // Tries to invert a 2D affine matrix.
    // ###########################################################################################
    private static bool TryInvert(Matrix m, out Matrix inv)
    {
        double a = m.M11, b = m.M12, c = m.M21, d = m.M22, e = m.M31, f = m.M32;
        double det = (a * d) - (b * c);

        if (Math.Abs(det) < 1e-12)
        {
            inv = Matrix.Identity;
            return false;
        }

        double idet = 1.0 / det;
        double na = d * idet, nb = -b * idet, nc = -c * idet, nd = a * idet;
        double ne = -((e * na) + (f * nc)), nf = -((e * nb) + (f * nd));

        inv = new Matrix(na, nb, nc, nd, ne, nf);
        return true;
    }

    // ###########################################################################################
    // Builds per-schematic highlight rect lookups.
    // ###########################################################################################
    public static Dictionary<string, Dictionary<string, List<Rect>>> BuildHighlightRects(BoardData boardData, string region)
    {
        var componentRegionsByLabel = boardData.Components
            .Where(c => !string.IsNullOrWhiteSpace(c.BoardLabel))
            .GroupBy(c => c.BoardLabel, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                g => g.Key,
                g => g.Select(c => c.Region?.Trim() ?? string.Empty)
                      .Where(r => !string.IsNullOrWhiteSpace(r))
                      .Distinct(StringComparer.OrdinalIgnoreCase)
                      .ToList(),
                StringComparer.OrdinalIgnoreCase);

        bool IsVisibleByRegion(string boardLabel)
        {
            if (!componentRegionsByLabel.TryGetValue(boardLabel, out var regionsForLabel)) return true;
            if (regionsForLabel.Count == 0) return true;
            return regionsForLabel.Any(r => string.Equals(r, region, StringComparison.OrdinalIgnoreCase));
        }

        var result = new Dictionary<string, Dictionary<string, List<Rect>>>(StringComparer.OrdinalIgnoreCase);

        foreach (var h in boardData.ComponentHighlights)
        {
            if (string.IsNullOrWhiteSpace(h.SchematicName) || string.IsNullOrWhiteSpace(h.BoardLabel)) continue;
            if (!IsVisibleByRegion(h.BoardLabel)) continue;

            if (!TryParseDouble(h.X, out var x) || !TryParseDouble(h.Y, out var y) ||
                !TryParseDouble(h.Width, out var w) || !TryParseDouble(h.Height, out var hh))
                continue;

            if (w <= 0 || hh <= 0) continue;

            if (!result.TryGetValue(h.SchematicName, out var byLabel))
            {
                byLabel = new Dictionary<string, List<Rect>>(StringComparer.OrdinalIgnoreCase);
                result[h.SchematicName] = byLabel;
            }

            if (!byLabel.TryGetValue(h.BoardLabel, out var rects))
            {
                rects = new List<Rect>();
                byLabel[h.BoardLabel] = rects;
            }

            rects.Add(new Rect(x, y, w, hh));
        }

        return result;
    }

    public static bool TryParseDouble(string text, out double value)
        => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);

    public static Color ParseColorOrDefault(string text, Color fallback)
    {
        if (string.IsNullOrWhiteSpace(text)) return fallback;
        try { return Color.Parse(text.Trim()); }
        catch { return fallback; }
    }

    public static double ParseOpacityOrDefault(string text, double fallback)
    {
        if (!TryParseDouble(text, out var v)) return fallback;
        if (v > 1.0) v /= 100.0;
        return Math.Clamp(v, 0.0, 1.0);
    }
}