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
using System.Collections.ObjectModel;
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
    //    internal List<SchematicThumbnail> currentThumbnails = new(); // compliant with .NET6
    internal ObservableCollection<SchematicThumbnail> currentThumbnails = new();

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

    // Thumbnail drag/drop reordering
    private Point thisThumbnailDragStartPoint;
    private bool thisIsDraggingThumbnail;
    private SchematicThumbnail? thisDraggedThumbnail;
    private int thisDraggedThumbnailOriginalIndex = -1;
    private bool thisDraggedThumbnailWasSelected;
    private double thisDraggedThumbnailHeight = 120.0;
    private SchematicThumbnail? thisThumbnailDropPlaceholder;
    private double thisDraggedThumbnailWidth = 160.0;
    private Point thisThumbnailDragPointerOffsetInItem;
    private int thisThumbnailCurrentInsertIndex = -1;
    private Point thisThumbnailDragStartPointInList;
    private double thisThumbnailLastPointerYInList = double.NaN;
    private double thisThumbnailDragGhostFixedX;
    private bool thisSuppressThumbnailSelectionChanged;

    // Fullscreen
    private bool thisIsFullscreenMode;
    private GridLength thisRestoreLeftColumnWidth = new(1, GridUnitType.Star);
    private GridLength thisRestoreSplitterColumnWidth = new(4, GridUnitType.Pixel);
    private GridLength thisRestoreRightColumnWidth = new(1, GridUnitType.Star);
    private double thisRestoreRightColumnMinWidth = 100.0;

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

        DragDrop.SetAllowDrop(this.SchematicsThumbnailList, true);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DragOverEvent, this.OnThumbnailDragOver);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DropEvent, this.OnThumbnailDrop);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DragLeaveEvent, this.OnThumbnailDragLeave);

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

            var border = new Border
            {
                Width = 48,
                Height = 14,
                Background = new SolidColorBrush(colorItem),
                BorderThickness = new Avalonia.Thickness(1),
                CornerRadius = new Avalonia.CornerRadius(2),
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };

            // Set DynamicResource for BorderBrush
            border.Bind(Border.BorderBrushProperty, this.GetResourceObservable("AppThemeBorderBrush"));

            var txt = new TextBlock
            {
                Text = $"({count})",
                FontSize = 11,
                VerticalAlignment = Avalonia.Layout.VerticalAlignment.Center
            };

            // Set DynamicResource for Foreground
            txt.Bind(TextBlock.ForegroundProperty, this.GetResourceObservable("Schematics_Panels_Fg"));

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
    // The image control already fits the bitmap to the available area, so matrix scale 1.0 is
    // the true minimum zoom and must not be reduced further.
    // ###########################################################################################
    private void OnSchematicsZoom(object? sender, PointerWheelEventArgs e)
    {
        var pos = e.GetPosition(this.SchematicsImage);
        double delta = e.Delta.Y > 0 ? AppConfig.SchematicsZoomFactor : 1.0 / AppConfig.SchematicsZoomFactor;

        double currentScale = this.schematicsMatrix.M11;
        double newScale = currentScale * delta;

        if (newScale > AppConfig.SchematicsMaxZoom)
            return;

        // The image is already fully fitted by Stretch="Uniform", so do not allow zooming out
        // below the baseline matrix scale of 1.0.
        if (delta < 1.0 && currentScale <= 1.0)
        {
            e.Handled = true;
            return;
        }

        // Snap cleanly back to the fitted baseline when zooming out crosses below 1.0.
        if (newScale < 1.0)
        {
            this.schematicsMatrix = Matrix.Identity;
            this.ClampSchematicsMatrix();
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

        // Fetch the active local region from the Main window, falling back to UserSettings if not attached
        string rawRegion = this.MainWindow?.LocalRegion?.Trim() ?? UserSettings.Region?.Trim() ?? string.Empty;
        this.SchematicsRegionLabel.Text = string.IsNullOrWhiteSpace(rawRegion) ? "All Regions" : rawRegion;

        bool hasExplicitRegions = this.MainWindow?.CurrentBoardHasExplicitRegionComponents() ?? true;
        this.SchematicsRegionBorder.IsVisible = this.SchematicsNameBorder.IsVisible && hasExplicitRegions;

        string regionKey = rawRegion.ToUpperInvariant();

        string colorPrefix = regionKey switch
        {
            "PAL" => "Schematics_Region_PAL",
            "NTSC" => "Schematics_Region_NTSC",
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
        if (this.thisSuppressThumbnailSelectionChanged)
            return;

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

        double currentScale = this.schematicsMatrix.M11;
        double inverseScale = currentScale > 0 ? 1.0 / currentScale : 1.0;

        foreach (var item in itemsToLoop)
        {
            if (string.IsNullOrWhiteSpace(item.BoardLabel)) continue;
            if (!seenLabels.Add(item.BoardLabel)) continue;

            if (!byLabel.TryGetValue(item.BoardLabel, out var rects) || rects.Count == 0) continue;

            var parts = item.SelectionKey?.Split('\u001F') ?? Array.Empty<string>();
            string friendlyName = parts.Length > 1 ? parts[1] : string.Empty;
            string technicalName = parts.Length > 2 ? parts[2] : string.Empty;

            var lines = new List<string>();
            if (this.CheckLabelBoard.IsChecked == true && !string.IsNullOrWhiteSpace(item.BoardLabel)) lines.Add(item.BoardLabel);
            if (this.CheckLabelTechnical.IsChecked == true && !string.IsNullOrWhiteSpace(technicalName)) lines.Add(technicalName);
            if (this.CheckLabelFriendly.IsChecked == true && !string.IsNullOrWhiteSpace(friendlyName)) lines.Add(friendlyName);

            if (lines.Count == 0) continue;

            string labelText = string.Join("\n", lines);

            foreach (var r in rects)
            {
                double centerX = r.X + (r.Width / 2.0);
                double centerY = r.Y + (r.Height / 2.0);

                double localX = contentRect.X + (centerX / imgWidth) * contentRect.Width;
                double localY = contentRect.Y + (centerY / imgHeight) * contentRect.Height;

                var tb = new TextBlock
                {
                    Text = labelText,
                    FontSize = 11,
                    FontWeight = FontWeight.Bold,
                    TextAlignment = TextAlignment.Center
                };
                tb.Bind(TextBlock.ForegroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Fg"));

                var innerBorder = new Border
                {
                    BorderThickness = new Avalonia.Thickness(1),
                    CornerRadius = new Avalonia.CornerRadius(4),
                    Padding = new Avalonia.Thickness(6, 4),
                    Child = tb
                };
                innerBorder.Bind(Border.BackgroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Bg"));
                innerBorder.Bind(Border.BorderBrushProperty, this.GetResourceObservable("AppThemeBorderBrush"));

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

    // ###########################################################################################
    // Safely resolves visual brushes from global Theme dictionaries, regardless of UI attach state.
    // ###########################################################################################
    private IBrush ResolveThemeBrush(string key, IBrush fallback)
    {
        if (this.TryFindResource(key, out var localRes) && localRes is IBrush localBrush)
            return localBrush;

        if (Application.Current != null)
        {
            var theme = Application.Current.ActualThemeVariant;
            if (Application.Current.TryGetResource(key, theme, out var appRes) && appRes is IBrush appBrush)
                return appBrush;
        }

        return fallback;
    }

    // ###########################################################################################
    // Loads thumbnails, applies saved user order, and removes stale saved entries automatically.
    // ###########################################################################################
    public void LoadSortedThumbnails(string boardKey, List<SchematicThumbnail> rawList)
    {
        var savedOrder = UserSettings.GetSchematicsOrder(boardKey);
        List<SchematicThumbnail> orderedList;

        if (savedOrder != null && savedOrder.Count > 0)
        {
            var orderLookup = savedOrder
                .Select((name, index) => new { name, index })
                .ToDictionary(x => x.name, x => x.index, StringComparer.OrdinalIgnoreCase);

            orderedList = rawList
                .OrderBy(x => orderLookup.TryGetValue(x.Name, out int orderIndex) ? orderIndex : int.MaxValue)
                .ToList();

            var currentNames = orderedList.Select(x => x.Name).ToList();
            if (!currentNames.SequenceEqual(savedOrder, StringComparer.OrdinalIgnoreCase))
            {
                UserSettings.SetSchematicsOrder(boardKey, currentNames);
            }
        }
        else
        {
            orderedList = rawList;
        }

        this.currentThumbnails.Clear();
        foreach (var thumbnail in orderedList)
        {
            this.currentThumbnails.Add(thumbnail);
        }

        this.SchematicsThumbnailList.ItemsSource = this.currentThumbnails;
    }

    // ###########################################################################################
    // Starts tracking a thumbnail for possible drag reorder.
    // Suppresses immediate ListBox selection so dragging does not replace the large schematic.
    // ###########################################################################################
    private void OnThumbnailPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;

        if (sender is not Control control || control.DataContext is not SchematicThumbnail thumbnail || thumbnail.IsDropPlaceholder)
            return;

        this.thisThumbnailDragStartPoint = e.GetPosition(this);
        this.thisThumbnailDragStartPointInList = e.GetPosition(this.SchematicsThumbnailList);
        this.thisThumbnailDragPointerOffsetInItem = e.GetPosition(control);
        this.thisIsDraggingThumbnail = true;
        this.thisDraggedThumbnail = thumbnail;
        this.thisDraggedThumbnailWasSelected = ReferenceEquals(this.SchematicsThumbnailList.SelectedItem, thumbnail);
        this.thisDraggedThumbnailHeight = Math.Max(control.Bounds.Height, 80.0);
        this.thisDraggedThumbnailWidth = Math.Max(control.Bounds.Width, 120.0);
        this.thisSuppressThumbnailSelectionChanged = true;

        e.Handled = true;
    }

    // ###########################################################################################
    // Begins drag/drop reordering once the pointer has moved far enough.
    // ###########################################################################################
    private async void OnThumbnailPointerMoved(object? sender, PointerEventArgs e)
    {
        if (!this.thisIsDraggingThumbnail || this.thisDraggedThumbnail == null)
            return;

        var point = e.GetPosition(this);
        var diff = this.thisThumbnailDragStartPoint - point;

        // Slightly higher threshold to avoid accidental detaching on tiny pointer jitter.
        if (Math.Abs(diff.X) <= 6 && Math.Abs(diff.Y) <= 6)
            return;

        if (sender is not Control control)
            return;

        this.thisIsDraggingThumbnail = false;

        this.thisDraggedThumbnailOriginalIndex = this.currentThumbnails.IndexOf(this.thisDraggedThumbnail);
        if (this.thisDraggedThumbnailOriginalIndex < 0)
            return;

        this.thisDraggedThumbnailHeight = Math.Max(control.Bounds.Height, 80.0);
        this.thisDraggedThumbnailWidth = Math.Max(control.Bounds.Width, 120.0);

        var pointerInList = e.GetPosition(this.SchematicsThumbnailList);
        this.thisThumbnailLastPointerYInList = pointerInList.Y;

        var transformToList = control.TransformToVisual(this.SchematicsThumbnailList);
        if (transformToList.HasValue)
        {
            var boundsInList = new Rect(control.Bounds.Size).TransformToAABB(transformToList.Value);
            this.thisThumbnailDragGhostFixedX = Math.Max(0, boundsInList.X);
        }
        else
        {
            this.thisThumbnailDragGhostFixedX = Math.Max(0, pointerInList.X - this.thisThumbnailDragPointerOffsetInItem.X);
        }

        this.currentThumbnails.RemoveAt(this.thisDraggedThumbnailOriginalIndex);
        this.thisThumbnailCurrentInsertIndex = this.thisDraggedThumbnailOriginalIndex;
        this.ShowThumbnailDropPlaceholder(this.thisDraggedThumbnailOriginalIndex);
        this.ShowThumbnailDragGhost(this.thisDraggedThumbnail, pointerInList);

        var dragData = new DataObject();
        dragData.Set("SchematicThumbnail", this.thisDraggedThumbnail);

        var effect = await DragDrop.DoDragDrop(e, dragData, DragDropEffects.Move);

        if (effect != DragDropEffects.Move && this.thisDraggedThumbnail != null)
        {
            this.RestoreDraggedThumbnail();
        }

        this.HideThumbnailDragGhost();
        this.ClearThumbnailDropPlaceholder();
        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
        this.thisDraggedThumbnailWasSelected = false;
        this.thisThumbnailCurrentInsertIndex = -1;
        this.thisThumbnailLastPointerYInList = double.NaN;

        e.Handled = true;
    }

    // ###########################################################################################
    // Updates the floating drag ghost to follow the mouse while preserving the original grab point.
    // Horizontal movement is locked so the detached thumbnail only moves vertically.
    // ###########################################################################################
    private void UpdateThumbnailDragGhostPosition(Point pointerInList)
    {
        if (!this.ThumbnailDragGhost.IsVisible || this.ThumbnailDragGhost.RenderTransform is not TranslateTransform transform)
            return;

        transform.X = this.thisThumbnailDragGhostFixedX;
        transform.Y = Math.Max(0, pointerInList.Y - this.thisThumbnailDragPointerOffsetInItem.Y);
    }

    // ###########################################################################################
    // Shows a detached visual thumbnail that follows the mouse during drag.
    // ###########################################################################################
    private void ShowThumbnailDragGhost(SchematicThumbnail thumbnail, Point pointerInList)
    {
        this.ThumbnailDragGhost.Width = this.thisDraggedThumbnailWidth;
        this.ThumbnailDragGhostName.Text = thumbnail.Name;
        this.ThumbnailDragGhostImage.Source = thumbnail.ImageSource ?? thumbnail.BaseThumbnail;
        this.ThumbnailDragGhost.IsVisible = true;
        this.UpdateThumbnailDragGhostPosition(pointerInList);
    }

    // ###########################################################################################
    // Hides the floating drag ghost after drop or cancel.
    // ###########################################################################################
    private void HideThumbnailDragGhost()
    {
        this.ThumbnailDragGhost.IsVisible = false;
        this.ThumbnailDragGhostName.Text = string.Empty;
        this.ThumbnailDragGhostImage.Source = null;
        this.thisThumbnailDragGhostFixedX = 0;

        if (this.ThumbnailDragGhost.RenderTransform is TranslateTransform transform)
        {
            transform.X = 0;
            transform.Y = 0;
        }
    }

    // ###########################################################################################
    // Stops local drag tracking when the pointer is released.
    // If no drag started, treat the interaction as a normal selection click.
    // ###########################################################################################
    private void OnThumbnailPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        if (this.thisIsDraggingThumbnail &&
            sender is Control control &&
            control.DataContext is SchematicThumbnail thumbnail &&
            !thumbnail.IsDropPlaceholder)
        {
            this.thisSuppressThumbnailSelectionChanged = false;
            this.SchematicsThumbnailList.SelectedItem = thumbnail;
            e.Handled = true;
        }

        this.thisIsDraggingThumbnail = false;
    }

    // ###########################################################################################
    // Resolves the thumbnail container border from the current visual source.
    // ###########################################################################################
    private Control? GetThumbnailContainer(Control? start)
    {
        Control? current = start;

        while (current != null)
        {
            if (current is Border border && border.Classes.Contains("ThumbnailBorder"))
                return border;

            current = current.Parent as Control;
        }

        return null;
    }

    // ###########################################################################################
    // Creates or moves the temporary placeholder item to the requested insert index.
    // The provided insert index is already in current collection coordinates, so no extra
    // reindex adjustment is applied after removing the existing placeholder.
    // ###########################################################################################
    private void ShowThumbnailDropPlaceholder(int insertIndex)
    {
        if (this.thisThumbnailDropPlaceholder == null)
        {
            this.thisThumbnailDropPlaceholder = new SchematicThumbnail
            {
                IsDropPlaceholder = true
            };
        }

        this.thisThumbnailDropPlaceholder.PlaceholderHeight = this.thisDraggedThumbnailHeight;
        this.thisThumbnailDropPlaceholder.PlaceholderWidth = this.thisDraggedThumbnailWidth;

        int existingIndex = this.currentThumbnails.IndexOf(this.thisThumbnailDropPlaceholder);

        if (existingIndex == insertIndex)
        {
            this.thisThumbnailCurrentInsertIndex = insertIndex;
            return;
        }

        if (existingIndex >= 0)
        {
            this.currentThumbnails.RemoveAt(existingIndex);
        }

        insertIndex = Math.Clamp(insertIndex, 0, this.currentThumbnails.Count);
        this.currentThumbnails.Insert(insertIndex, this.thisThumbnailDropPlaceholder);
        this.thisThumbnailCurrentInsertIndex = insertIndex;
    }

    // ###########################################################################################
    // Removes the temporary placeholder item from the thumbnails list.
    // ###########################################################################################
    private void ClearThumbnailDropPlaceholder()
    {
        if (this.thisThumbnailDropPlaceholder == null)
            return;

        int index = this.currentThumbnails.IndexOf(this.thisThumbnailDropPlaceholder);
        if (index >= 0)
        {
            this.currentThumbnails.RemoveAt(index);
        }
    }

    // ###########################################################################################
    // Restores the dragged thumbnail to its original position when no drop occurs.
    // ###########################################################################################
    private void RestoreDraggedThumbnail()
    {
        if (this.thisDraggedThumbnail == null)
            return;

        this.ClearThumbnailDropPlaceholder();

        int restoreIndex = this.thisDraggedThumbnailOriginalIndex;
        if (restoreIndex < 0 || restoreIndex > this.currentThumbnails.Count)
        {
            restoreIndex = this.currentThumbnails.Count;
        }

        this.currentThumbnails.Insert(restoreIndex, this.thisDraggedThumbnail);
        this.thisThumbnailCurrentInsertIndex = restoreIndex;
        this.thisThumbnailLastPointerYInList = double.NaN;

        if (this.thisDraggedThumbnailWasSelected)
        {
            this.SchematicsThumbnailList.SelectedItem = this.thisDraggedThumbnail;
        }

        this.thisSuppressThumbnailSelectionChanged = false;
    }

    // ###########################################################################################
    // Saves the current thumbnail order for the active board, excluding the placeholder item.
    // ###########################################################################################
    private void SaveCurrentThumbnailOrder()
    {
        var boardKey = this.MainWindow?.GetCurrentBoardKey();
        if (string.IsNullOrWhiteSpace(boardKey))
            return;

        var orderedNames = this.currentThumbnails
            .Where(x => !x.IsDropPlaceholder && !string.IsNullOrWhiteSpace(x.Name))
            .Select(x => x.Name)
            .ToList();

        UserSettings.SetSchematicsOrder(boardKey, orderedNames);
    }

    // ###########################################################################################
    // Updates the placeholder position while dragging over the thumbnail list.
    // Uses actual ListBoxItem row bounds and only moves in the current drag direction.
    // ###########################################################################################
    private void OnThumbnailDragOver(object? sender, DragEventArgs e)
    {
        if (!e.Data.Contains("SchematicThumbnail") || this.thisDraggedThumbnail == null)
        {
            e.DragEffects = DragDropEffects.None;
            e.Handled = true;
            return;
        }

        e.DragEffects = DragDropEffects.Move;

        var pointerInList = e.GetPosition(this.SchematicsThumbnailList);
        this.UpdateThumbnailDragGhostPosition(pointerInList);

        int placeholderIndex = this.thisThumbnailDropPlaceholder != null
            ? this.currentThumbnails.IndexOf(this.thisThumbnailDropPlaceholder)
            : -1;

        if (placeholderIndex < 0)
        {
            e.Handled = true;
            return;
        }

        if (double.IsNaN(this.thisThumbnailLastPointerYInList))
        {
            this.thisThumbnailLastPointerYInList = pointerInList.Y;
            e.Handled = true;
            return;
        }

        double deltaY = pointerInList.Y - this.thisThumbnailLastPointerYInList;
        if (Math.Abs(deltaY) < 0.1)
        {
            e.Handled = true;
            return;
        }

        bool isMovingUp = deltaY < 0;
        bool isMovingDown = deltaY > 0;

        double ghostTopY = pointerInList.Y - this.thisThumbnailDragPointerOffsetInItem.Y;
        double ghostBottomY = ghostTopY + this.thisDraggedThumbnailHeight;

        if (isMovingUp && placeholderIndex > 0)
        {
            var itemAbove = this.SchematicsThumbnailList.ContainerFromIndex(placeholderIndex - 1) as ListBoxItem;
            if (itemAbove != null)
            {
                var transform = itemAbove.TransformToVisual(this.SchematicsThumbnailList);
                if (transform.HasValue)
                {
                    var boundsAbove = new Rect(itemAbove.Bounds.Size).TransformToAABB(transform.Value);

                    if (ghostTopY <= boundsAbove.Bottom)
                    {
                        this.ShowThumbnailDropPlaceholder(placeholderIndex - 1);
                        this.thisThumbnailLastPointerYInList = pointerInList.Y;
                        e.Handled = true;
                        return;
                    }
                }
            }
        }

        if (isMovingDown && placeholderIndex < this.currentThumbnails.Count - 1)
        {
            var itemBelow = this.SchematicsThumbnailList.ContainerFromIndex(placeholderIndex + 1) as ListBoxItem;
            if (itemBelow != null)
            {
                var transform = itemBelow.TransformToVisual(this.SchematicsThumbnailList);
                if (transform.HasValue)
                {
                    var boundsBelow = new Rect(itemBelow.Bounds.Size).TransformToAABB(transform.Value);

                    if (ghostBottomY >= boundsBelow.Top)
                    {
                        this.ShowThumbnailDropPlaceholder(placeholderIndex + 1);
                        this.thisThumbnailLastPointerYInList = pointerInList.Y;
                        e.Handled = true;
                        return;
                    }
                }
            }
        }

        this.thisThumbnailLastPointerYInList = pointerInList.Y;
        e.Handled = true;
    }

    // ###########################################################################################
    // Keeps the placeholder visible even if the drag briefly leaves the list bounds.
    // ###########################################################################################
    private void OnThumbnailDragLeave(object? sender, RoutedEventArgs e)
    {
        e.Handled = true;
    }

    // ###########################################################################################
    // Finalizes thumbnail reordering by replacing the placeholder with the dragged item.
    // ###########################################################################################
    private void OnThumbnailDrop(object? sender, DragEventArgs e)
    {
        if (!e.Data.Contains("SchematicThumbnail") || this.thisDraggedThumbnail == null)
        {
            e.Handled = true;
            return;
        }

        int insertIndex = this.thisThumbnailDropPlaceholder != null
            ? this.currentThumbnails.IndexOf(this.thisThumbnailDropPlaceholder)
            : this.currentThumbnails.Count;

        this.ClearThumbnailDropPlaceholder();

        if (insertIndex < 0 || insertIndex > this.currentThumbnails.Count)
        {
            insertIndex = this.currentThumbnails.Count;
        }

        this.currentThumbnails.Insert(insertIndex, this.thisDraggedThumbnail);
        this.thisThumbnailCurrentInsertIndex = insertIndex;

        if (this.thisDraggedThumbnailWasSelected)
        {
            this.SchematicsThumbnailList.SelectedItem = this.thisDraggedThumbnail;
        }

        this.HideThumbnailDragGhost();
        this.SaveCurrentThumbnailOrder();

        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
        this.thisDraggedThumbnailWasSelected = false;
        this.thisThumbnailCurrentInsertIndex = -1;
        this.thisThumbnailLastPointerYInList = double.NaN;
        this.thisSuppressThumbnailSelectionChanged = false;

        e.Handled = true;
    }

    // ###########################################################################################
    // Expands the control into image-only mode before it is rehosted in the fullscreen window.
    // ###########################################################################################
    public void EnterFullscreenMode()
    {
        if (this.thisIsFullscreenMode)
            return;

        this.thisIsFullscreenMode = true;

        this.thisRestoreLeftColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[0].Width;
        this.thisRestoreSplitterColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[1].Width;
        this.thisRestoreRightColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[2].Width;
        this.thisRestoreRightColumnMinWidth = this.SchematicsInnerGrid.ColumnDefinitions[2].MinWidth;

        this.thisIsDraggingThumbnail = false;
        this.ClearThumbnailDropPlaceholder();
        this.HideThumbnailDragGhost();

        this.SchematicsInnerGrid.ColumnDefinitions[0].Width = new GridLength(1, GridUnitType.Star);
        this.SchematicsInnerGrid.ColumnDefinitions[1].Width = new GridLength(0, GridUnitType.Pixel);
        this.SchematicsInnerGrid.ColumnDefinitions[2].Width = new GridLength(0, GridUnitType.Pixel);
        this.SchematicsInnerGrid.ColumnDefinitions[2].MinWidth = 0;

        this.SchematicsSplitter.IsVisible = false;
        this.SchematicsThumbnailList.IsVisible = false;

        this.RefreshAfterHostChanged();
    }

    // ###########################################################################################
    // Restores the normal schematics tab layout after leaving fullscreen mode.
    // ###########################################################################################
    public void ExitFullscreenMode()
    {
        if (!this.thisIsFullscreenMode)
            return;

        this.thisIsFullscreenMode = false;

        this.SchematicsInnerGrid.ColumnDefinitions[0].Width = this.thisRestoreLeftColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[1].Width = this.thisRestoreSplitterColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[2].Width = this.thisRestoreRightColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[2].MinWidth = this.thisRestoreRightColumnMinWidth;

        this.SchematicsSplitter.IsVisible = true;
        this.SchematicsThumbnailList.IsVisible = true;

        this.RefreshAfterHostChanged();
    }

    // ###########################################################################################
    // Re-clamps and redraws the viewer after the control is moved between windows.
    // ###########################################################################################
    public void RefreshAfterHostChanged()
    {
        Dispatcher.UIThread.Post(() =>
        {
            this.ClampSchematicsMatrix();
            this.UpdateOverlayLabels();
            this.UpdateComponentLabels();
        }, DispatcherPriority.Background);

        Dispatcher.UIThread.Post(() =>
        {
            this.ClampSchematicsMatrix();
        }, DispatcherPriority.Background);
    }

}