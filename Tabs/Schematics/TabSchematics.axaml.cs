// ###########################################################################################
// TabSchematics - the Schematics tab.
//
// Renders board schematics and PCB images with overlays on top: component highlights, an
// interactive KiCad trace/copper overlay, and user-drawn polyline traces. It also hosts
// the component label editor and the MiniPro IC test panel.
//
// The class is big, so it is split across several files. Each part owns one area and
// documents its own scope in a header comment:
//
//   TabSchematics.Types.cs                    Private data types shared by the parts below
//   TabSchematics.Viewport.cs                 Zoom, pan and the schematic transform matrix
//   TabSchematics.Input.cs                    Pointer, wheel, gesture and keyboard event handlers
//   TabSchematics.Thumbnails.cs               The schematic thumbnail list and drag-to-reorder
//   TabSchematics.Highlights.cs               Component highlight overlays and on-schematic labels
//   TabSchematics.LabelEditor.cs              Label editor mode: lifecycle, save/validate, undo
//   TabSchematics.LabelEditor.TestSeams.cs    ...ForTests seams letting headless tests drive the
//                                             editor - see that file's own header
//   TabSchematics.LabelEditor.Interaction.cs  Label editor pointer/keyboard interaction
//   TabSchematics.LabelEditor.Snap.cs         Builds the snap context from tab state; the maths
//                                             itself is Handlers/Geometry/LabelEditorSnapGeometry
//   TabSchematics.KiCad.cs                    KiCad project state, selection and cache scopes
//   TabSchematics.KiCad.Panels.cs             The Important signals and Net connections panels
//   TabSchematics.KiCad.Render.cs             Draws the KiCad overlay onto the schematic
//   TabSchematics.KiCad.RenderCache.cs        Builds and caches the per-net PCB render nodes
//   TabSchematics.KiCad.Geometry.cs           KiCad world/screen mapping and zone geometry
//   TabSchematics.KiCad.HitTest.cs            Hover hit-testing over the KiCad overlay
//   TabSchematics.KiCad.Calibration.cs        Interactive KiCad trace calibration mode (maths in
//                                             Handlers/Geometry/KiCadCalibrationGeometry)
//   TabSchematics.Worklog.cs                  Worklog "Add worklog" area-drawing mode (which opens
//                                             the full editor), the saved-entry overlay and its
//                                             anchored and parked pills
//   TabSchematics.Worklog.TestSeams.cs        ...ForTests seams for the area-drawing flow - see
//                                             that file's own header
//   TabSchematics.Settings.cs                 Board-level and global setting rows
//
// They are one partial class, so state is shared; each field is declared in the part
// that owns it.
// ###########################################################################################

using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;
using Handlers.Theming;

namespace CRT;

// ###########################################################################################
// Shell for the Schematics tab: construction, one-time wiring in Initialize, the
// fullscreen/splitter layout, the user-drawn trace colour palette, and small shared
// parsing/theme helpers used by the other TabSchematics parts.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics : UserControl
{
    public Main? MainWindow { get; set; }

    public bool IsLabelEditorActive => this.thisIsLabelEditorMode || this.thisIsKiCadTraceCalibrationMode;

    // Full-res viewer
    internal Bitmap? currentFullResBitmap;

    internal CancellationTokenSource? fullResLoadCts;

    internal Dictionary<string, BoardSchematicEntry> schematicByName = new(StringComparer.OrdinalIgnoreCase);

    // Insert logic class declaration
    internal PolylineManagement? polylineManager;

    // Fullscreen
    private bool thisIsFullscreenMode;

    private GridLength thisRestoreLeftColumnWidth = new(1, GridUnitType.Star);

    private GridLength thisRestoreSplitterColumnWidth = new(4, GridUnitType.Pixel);

    private GridLength thisRestoreRightColumnWidth = new(1, GridUnitType.Star);

    private double thisRestoreRightColumnMinWidth = 100.0;

    // ###########################################################################################
    // Initializes the schematics tab control and wires the shared UI actions used by the viewer,
    // label editor, and KiCad calibration workflows.
    // ###########################################################################################
    public TabSchematics()
    {
        InitializeComponent();

        this.ClearImportantSignalsButton.Click += (_, _) => this.ClearImportantSignalsSelection();

        this.EnableLabelEditorButton.Click += (_, _) => this.BeginLabelEditorMode();
        this.CancelLabelEditorChangesButton.Click += (_, _) => this.CancelLabelEditorChanges();
        this.ApplyLabelEditorChangesButton.Click += (_, _) => this.ApplyLabelEditorChanges();

        this.ConfirmNewLabelEditorBoardLabelButton.Click += (_, _) => this.ConfirmNewLabelEditorPrompt();
        this.CancelNewLabelEditorBoardLabelButton.Click += (_, _) => this.CancelNewLabelEditorPrompt();

        this.NewLabelEditorBoardLabelTextBox.KeyDown += this.OnNewLabelEditorPromptKeyDown;
        this.NewLabelEditorCategoryComboBox.KeyDown += this.OnNewLabelEditorPromptKeyDown;

        this.ClearKiCadTraceSelectionButton.Click += (_, _) => this.ClearAllKiCadTraceSelections();

        this.BeginKiCadTraceCalibrationButton.Click += (_, _) => this.BeginKiCadTraceCalibrationMode();
        this.ApplyKiCadTraceCalibrationButton.Click += (_, _) => this.ApplyKiCadTraceCalibration();
        this.CancelKiCadTraceCalibrationButton.Click += (_, _) => this.CancelKiCadTraceCalibrationMode();

        this.CheckGlobalShowCalibrationTracesAndPads.IsChecked = true;
        this.CheckGlobalShowCalibrationTracesAndPads.IsCheckedChanged += (_, _) =>
        {
            if (this.thisIsKiCadTraceCalibrationMode)
            {
                this.RefreshKiCadOverlay(forceImmediate: true);
            }
        };
    }

    public void Initialize(Main mainWindow)
    {
        this.MainWindow = mainWindow;

        // Only this overlay instance draws worklog entry areas, so its dashed-border mode is set
        // once here rather than threaded through every ApplyState call.
        this.SchematicsWorklogEntryOverlay.UseDashedBorder = true;

        var thisPinchGestureRecognizer = new PinchGestureRecognizer();
        this.SchematicsContainer.GestureRecognizers.Add(thisPinchGestureRecognizer);

        var thisScrollGestureRecognizer = new ScrollGestureRecognizer
        {
            CanHorizontallyScroll = true,
            CanVerticallyScroll = true
        };
        this.SchematicsContainer.GestureRecognizers.Add(thisScrollGestureRecognizer);

        this.SchematicsContainer.Pinch += this.OnSchematicsPinch;
        this.SchematicsContainer.ScrollGesture += this.OnSchematicsScrollGesture;

        // Parked worklog pills sit in the schematic panel's top-right corner and step aside for
        // the "Netlist names" panel. Subscribed here rather than calling the layout from each of
        // the six places that set IsVisible - a seventh would be added one day and quietly not
        // move the pills. Bounds covers the panel resizing to fit a longer net name, and the
        // container's own Bounds covers the window or splitter being dragged.
        this.KiCadNetConnectionsPanel.PropertyChanged += this.OnWorklogParkedBadgeLayoutTriggerChanged;
        this.SchematicsContainer.PropertyChanged += this.OnWorklogParkedBadgeLayoutTriggerChanged;

        // Restore initial states from User Settings
        this.CheckLabelBoard.IsChecked = UserSettings.SchematicsLabelBoard;
        this.CheckLabelTechnical.IsChecked = UserSettings.SchematicsLabelTechnical;
        this.CheckLabelFriendly.IsChecked = UserSettings.SchematicsLabelFriendly;
        this.CheckLabelSelectedOnly.IsChecked = UserSettings.SchematicsLabelSelectedOnly;

        this.thisSuppressBoardSettingsChanged = true;
        this.thisSuppressGlobalSettingsChanged = true;
        this.CheckBoardMarkPin1OnSelectedComponent.IsChecked = true;
        this.CheckBoardShowTracesOnSelectedComponent.IsChecked = UserSettings.SchematicsShowTracesOnSelectedComponent;
        this.CheckGlobalShowTracesOnComponentSelect.IsChecked = UserSettings.SchematicsShowTracesOnComponentSelect;
        this.CheckGlobalShowOppositeSideTraces.IsChecked = UserSettings.SchematicsShowOppositeSideTraces;
        this.CheckGlobalShowZones.IsChecked = UserSettings.SchematicsShowZones;
        this.CheckGlobalHoverHighlightsTraces.IsChecked = true;
        this.CheckBoardContributorMode.IsChecked = UserSettings.ContributorMode;
        this.thisSuppressGlobalSettingsChanged = false;
        this.thisSuppressBoardSettingsChanged = false;

        this.UpdateGlobalSettingsControls();

        DragDrop.SetAllowDrop(this.SchematicsThumbnailList, true);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DragOverEvent, this.OnThumbnailDragOver);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DropEvent, this.OnThumbnailDrop);
        this.SchematicsThumbnailList.AddHandler(DragDrop.DragLeaveEvent, this.OnThumbnailDragLeave);

        this.AddHandler(
            InputElement.KeyDownEvent,
            this.OnSchematicsKeyDown,
            RoutingStrategies.Tunnel,
            handledEventsToo: true);

        this.AddHandler(
            InputElement.KeyUpEvent,
            this.OnSchematicsKeyUp,
            RoutingStrategies.Tunnel,
            handledEventsToo: true);

        UserSettings.InteractiveCadTraceHoverModeChanged += this.OnInteractiveCadTraceHoverModeChanged;

        bool isLabelsExpanded = UserSettings.SchematicsLabelsPanelExpanded;
        this.LabelsListPanel.IsVisible = isLabelsExpanded;
        this.ToggleLabelsPanelButton.Content = isLabelsExpanded ? "Collapse" : "Expand";

        bool isGlobalSettingsExpanded = UserSettings.SchematicsGlobalSettingsPanelExpanded;
        this.GlobalSettingsListPanel.IsVisible = isGlobalSettingsExpanded;
        this.ToggleGlobalSettingsPanelButton.Content = isGlobalSettingsExpanded ? "Collapse" : "Expand";

        bool isImportantSignalsExpanded = UserSettings.SchematicsImportantSignalsPanelExpanded;
        this.UpdateImportantSignalsPanelExpandedState(isImportantSignalsExpanded);
        this.UpdateImportantSignalsClearButtonState(false);

        bool isKiCadNetConnectionsExpanded = UserSettings.SchematicsNetConnectionsPanelExpanded;
        this.UpdateKiCadNetConnectionsPanelExpandedState(isKiCadNetConnectionsExpanded);

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

        this.CheckBoardMarkPin1OnSelectedComponent.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressBoardSettingsChanged)
            {
                return;
            }

            var boardKey = this.MainWindow?.GetCurrentBoardKey();
            if (string.IsNullOrWhiteSpace(boardKey))
            {
                return;
            }

            UserSettings.SetSchematicsMarkPin1OnSelectedComponentForBoard(
                boardKey,
                this.CheckBoardMarkPin1OnSelectedComponent.IsChecked == true);

            this.RefreshKiCadOverlay(forceImmediate: true);
        };

        this.CheckBoardShowTracesOnSelectedComponent.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressBoardSettingsChanged)
            {
                return;
            }

            UserSettings.SchematicsShowTracesOnSelectedComponent =
                this.CheckBoardShowTracesOnSelectedComponent.IsChecked == true;

            this.RefreshKiCadOverlay(forceImmediate: true);
        };

        this.CheckGlobalShowOppositeSideTraces.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressGlobalSettingsChanged)
            {
                return;
            }

            UserSettings.SchematicsShowOppositeSideTraces =
                this.CheckGlobalShowOppositeSideTraces.IsChecked == true;

            this.RefreshKiCadOverlay(forceImmediate: true);
        };

        this.CheckGlobalShowZones.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressGlobalSettingsChanged)
            {
                return;
            }

            UserSettings.SchematicsShowZones =
                this.CheckGlobalShowZones.IsChecked != false;

            this.RefreshKiCadOverlay(forceImmediate: true);
        };

        this.CheckGlobalShowTracesOnComponentSelect.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressGlobalSettingsChanged)
            {
                return;
            }

            UserSettings.SchematicsShowTracesOnComponentSelect =
                this.CheckGlobalShowTracesOnComponentSelect.IsChecked == true;

            this.RefreshKiCadOverlay(forceImmediate: true);
        };

        this.CheckGlobalHoverHighlightsTraces.IsCheckedChanged += (s, e) =>
        {
            this.ApplyInteractiveCadTraceHoverModeFromGlobalSettings();
            this.UpdateInteractiveCadTraceHoverModeUi();
            this.RefreshKiCadHoverPadUi();
            this.RefreshKiCadOverlay();
        };

        this.CheckBoardContributorMode.IsCheckedChanged += (s, e) =>
        {
            if (this.thisSuppressBoardSettingsChanged)
            {
                return;
            }

            bool isEnabled = this.CheckBoardContributorMode.IsChecked == true;
            UserSettings.ContributorMode = isEnabled;
        };

        this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsCheckedChanged +=
            this.OnSchematicsInteractiveCadTraceHoverModeChanged;

        this.ToggleLabelsPanelButton.Click += (s, e) =>
        {
            bool willBeExpanded = !this.LabelsListPanel.IsVisible;
            this.LabelsListPanel.IsVisible = willBeExpanded;
            this.ToggleLabelsPanelButton.Content = willBeExpanded ? "Collapse" : "Expand";
            UserSettings.SchematicsLabelsPanelExpanded = willBeExpanded;
        };

        this.ToggleGlobalSettingsPanelButton.Click += (s, e) =>
        {
            bool willBeExpanded = !this.GlobalSettingsListPanel.IsVisible;
            this.GlobalSettingsListPanel.IsVisible = willBeExpanded;
            this.ToggleGlobalSettingsPanelButton.Content = willBeExpanded ? "Collapse" : "Expand";
            UserSettings.SchematicsGlobalSettingsPanelExpanded = willBeExpanded;
        };

        this.ToggleImportantSignalsPanelButton.Click += (s, e) =>
        {
            bool willBeExpanded = !this.ImportantSignalsListPanel.IsVisible;
            this.UpdateImportantSignalsPanelExpandedState(willBeExpanded);
            this.UpdateImportantSignalsClearButtonState(this.ImportantSignalsPanel.IsVisible);
            UserSettings.SchematicsImportantSignalsPanelExpanded = willBeExpanded;
        };

        this.ImportantSignalsListBox.SelectionChanged += (s, e) =>
        {
            this.SyncSelectedImportantSignalsFromList();
            this.RefreshKiCadOverlay(forceImmediate: true);
            this.RefreshBlinkStateFromCurrentSelection();
        };

        this.ToggleKiCadNetConnectionsPanelButton.Click += (s, e) =>
        {
            bool willBeExpanded = !this.KiCadNetConnectionsContentPanel.IsVisible;
            this.UpdateKiCadNetConnectionsPanelExpandedState(willBeExpanded);
            UserSettings.SchematicsNetConnectionsPanelExpanded = willBeExpanded;
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
        this.SchematicsHoverHighlightsOverlay.RenderTransformOrigin = RelativePoint.TopLeft;

        this.SchematicsContainer.PropertyChanged += (s, e) =>
        {
            if (e.Property == Visual.BoundsProperty)
            {
                this.ClampSchematicsMatrix();
                this.UpdateComponentLabels();
                this.RefreshKiCadOverlay();
            }
        };

        this.SchematicsImage.PropertyChanged += (s, e) =>
        {
            if (e.Property == Visual.BoundsProperty)
            {
                this.ClampSchematicsMatrix();
                this.UpdateComponentLabels();
                this.RefreshKiCadOverlay();
            }
        };

        this.SchematicsThumbnailList.SelectionChanged += this.OnSchematicsThumbnailSelectionChanged;
        this.SchematicsContainer.PointerExited += this.OnSchematicsPointerExited;

        this.UpdateInteractiveCadTraceHoverModeUi();

        this.UpdateLabelEditorLockState();
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
            border.Bind(Border.BorderBrushProperty, this.GetResourceObservable("Schematics_Panel_TracesVisible_Trace_Border"));

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
    // Returns true when the pointer is currently inside any interactive overlay panel that should
    // consume wheel input instead of letting the schematic viewer zoom underneath it.
    // ###########################################################################################
    private bool IsPointerInsideInteractiveOverlayPanel(Point containerPoint)
    {
        bool IsInsideVisibleOverlay(Control? overlay)
        {
            if (overlay == null ||
                !overlay.IsVisible ||
                overlay.Bounds.Width <= 0 ||
                overlay.Bounds.Height <= 0)
            {
                return false;
            }

            Point? translatedTopLeft = overlay.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
            if (!translatedTopLeft.HasValue)
            {
                return false;
            }

            var overlayRect = new Rect(translatedTopLeft.Value, overlay.Bounds.Size);
            return overlayRect.Contains(containerPoint);
        }

        return IsInsideVisibleOverlay(this.GlobalSettingsPanel) ||
               IsInsideVisibleOverlay(this.LabelsPanel) ||
               IsInsideVisibleOverlay(this.ImportantSignalsPanel) ||
               IsInsideVisibleOverlay(this.TracesPanel) ||
               IsInsideVisibleOverlay(this.KiCadNetConnectionsPanel) ||
               IsInsideVisibleOverlay(this.TraceFloatingPalette) ||
               IsInsideVisibleOverlay(this.SchematicsLabelEditorMenuBorder) ||
               IsInsideVisibleOverlay(this.SchematicsNewLabelPromptBorder);
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
    // Resets only the active UI/session state and leaves persistent per-board KiCad runtime
    // caches intact so returning to a previously visited board can reuse them.
    // ###########################################################################################
    public void ResetSchematicsViewer()
    {
        this.thisIsKiCadOverlayRefreshQueued = false;
        this.thisKiCadOverlayRefreshRequestVersion = 0;
        this.thisKiCadOverlayLastRenderedVersion = 0;
        this.thisKiCadSchematicHoverHitTestCacheByKey.Clear();
        this.thisKiCadProject = null;
        this.thisCurrentKiCadRuntimeCacheScopeKey = string.Empty;

        lock (this.thisKiCadPcbNetRenderCacheSync)
        {
            this.thisKiCadPcbNetRenderBuildTaskByKey.Clear();
        }

        lock (this.thisKiCadPcbHoverHitTestCacheSync)
        {
            this.thisKiCadPcbHoverHitTestBuildTaskByKey.Clear();
        }

        this.SchematicsKiCadOverlayCanvas.ClearGeometry();
        ((MatrixTransform)this.SchematicsKiCadOverlayCanvas.RenderTransform!).Matrix = this.schematicsMatrix;

        this.KiCadNetConnectionsPanel.IsVisible = false;
        this.KiCadNetConnectionsHeaderTextBlock.Text = "Netlist name";
        this.KiCadNetConnectionsList.ItemsSource = null;
        this.UpdateKiCadNetConnectionsClearButtonState(false);

        this.UpdateImportantSignalsHeaderText();
        this.ImportantSignalsListBox.ItemsSource = null;
        this.ImportantSignalsListBox.SelectedItems?.Clear();
        this.ImportantSignalsPanel.IsVisible = false;
        this.UpdateImportantSignalsClearButtonState(false);

        this.thisSelectedKiCadReferences.Clear();
        this.thisSelectedKiCadNormalizedNetNames.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();

        this.thisHoveredKiCadNetName = null;
        this.thisHoveredKiCadPadNumber = null;
        this.thisHoveredComponentBoardLabel = null;
        this.thisSchematicsOnlySelectedBoardLabels.Clear();
        this.thisLockedKiCadNetNames.Clear();

        this.thisIsKiCadTraceCalibrationMode = false;
        this.thisKiCadTraceCalibrationDragMode = LabelEditorDragMode.None;
        this.thisKiCadCalibrationImageLeft = 0.0;
        this.thisKiCadCalibrationImageTop = 0.0;
        this.thisKiCadCalibrationImageRight = 0.0;
        this.thisKiCadCalibrationImageBottom = 0.0;
        this.thisKiCadCalibrationStartImageLeft = 0.0;
        this.thisKiCadCalibrationStartImageTop = 0.0;
        this.thisKiCadCalibrationStartImageRight = 0.0;
        this.thisKiCadCalibrationStartImageBottom = 0.0;

        this.thisIsDraggingThumbnail = false;
        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
        this.thisDraggedThumbnailWasSelected = false;
        this.thisThumbnailLastPointerYInList = double.NaN;
        this.thisThumbnailDragStartEventArgs = null;
        this.ClearThumbnailDropPlaceholder();
        this.HideThumbnailDragGhost();

        this.polylineManager?.Reset();
        this.SchematicsLabelsCanvas.Children.Clear();
        this.ResetComponentLabelVisualCaches();
        this.ResetWorklogOverlays();

        this.SetTraceColorPickerColor(Colors.White);
        this.CustomColorButton.Background = Brushes.White;

        this.ResetKiCadHoverHitTestThrottle();
        this.thisLastKiCadNetConnectionsSignature = string.Empty;
        this.thisLastThumbnailHighlightSignature = string.Empty;

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
        ((MatrixTransform)this.SchematicsHoverHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsLabelEditorOverlay.RenderTransform!).Matrix = this.schematicsMatrix;

        this.SchematicsHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        this.SchematicsHoverHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHoverHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHoverHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        this.thisIsLabelEditorMode = false;
        this.thisLabelEditorSchematicName = string.Empty;
        this.thisLastCreatedLabelEditorCategory = string.Empty;
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;
        this.thisPendingNewLabelEditorHighlight = null;
        this.thisIsDrawingLabelEditorRectangle = false;
        this.thisLabelEditorDraftRectangle = null;
        this.thisLabelEditorDragMode = LabelEditorDragMode.None;
        this.thisLabelEditorOriginalSelectionBounds = default;
        this.thisLabelEditorOriginalDragRectangles.Clear();
        this.thisLabelEditorWorkingHighlights.Clear();
        this.UpdateLabelEditorLockState();
        this.HideLabelEditorMenu();
        this.HideNewLabelEditorPrompt();

        this.SchematicsLabelEditorOverlay.ApplyState(
            rectangles: Array.Empty<Rect>(),
            selectedIndex: -1,
            selectedIndices: Array.Empty<int>(),
            selectionBounds: null,
            hoveredIndex: -1,
            draftRectangle: null,
            snapGuides: Array.Empty<(Point Start, Point End)>(),
            bitmapPixelSize: new PixelSize(0, 0),
            viewMatrix: this.schematicsMatrix,
            highlightColor: Colors.IndianRed,
            highlightOpacity: 0.20,
            isVisible: false);

        this.RestoreBoardSettings(string.Empty);

        this.isPanning = false;
        this.HideSchematicsHoverUi();
    }

    // ###########################################################################################
    // Safely resolves visual brushes from global Theme dictionaries, regardless of UI attach state.
    // ###########################################################################################
    private IBrush ResolveThemeBrush(string key, IBrush fallback) =>
        ThemeResources.ResolveForControl(this, key, fallback);

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
        this.thisThumbnailDragStartEventArgs = null;
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
    // Uses one full label refresh and then a follow-up clamp-only pass for final layout settling.
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

    // ###########################################################################################
    // Returns true when the schematics actions menu is allowed to be shown.
    // Contributor mode enables menu entry from empty-space right click, and active editor or KiCad
    // calibration workflows keep the same shared floating menu available.
    //
    // Never during worklog entry mode: on a contributor-mode board the first clause is true, so the
    // menu could still open while an area was being marked out - and it offers "Enable label editor"
    // and the calibration mode, which is exactly how two modes ended up active at once.
    // ###########################################################################################
    private bool CanShowSchematicsActionsMenu()
    {
        if (this.thisIsWorklogEntryMode)
        {
            return false;
        }

        return IsBoardContributorModeEnabled() ||
               this.thisIsLabelEditorMode ||
               this.thisIsKiCadTraceCalibrationMode;
    }

    // ###########################################################################################
    // Returns the currently selected schematic name, or an empty string if none is selected.
    //
    // internal rather than private so the component popup's "attach capture to worklog" flow can
    // name the schematic a new entry belongs to. An entry with a blank SchematicName is invisible
    // on BOTH surfaces that draw entries - this tab filters by name (RefreshWorklogEntriesList) and
    // so does the Workbooks board pane - so a worklog created from the oscilloscope has to be filed
    // against a real schematic or it can never be seen again.
    // ###########################################################################################
    internal string GetCurrentSchematicName()
    {
        return (this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail)?.Name?.Trim() ?? string.Empty;
    }

    // ###########################################################################################
    // Applies the persisted schematics/thumbnail split ratio for the supplied board key.
    // Called early so the splitter does not flash at the default centered position.
    // ###########################################################################################
    public void ApplySchematicsSplitterRatio(string boardKey)
    {
        double ratio = Math.Clamp(
            string.IsNullOrWhiteSpace(boardKey)
                ? 0.5
                : UserSettings.GetSchematicsSplitterRatio(boardKey),
            0.05,
            0.95);

        this.SchematicsInnerGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
        this.SchematicsInnerGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);
    }
}