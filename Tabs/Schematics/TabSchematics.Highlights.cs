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

namespace CRT;

// ###########################################################################################
// Component highlight overlays and the labels drawn on top of a schematic: selection by
// board label, blink visuals, hover UI, and the pooled standard component label visuals.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Highlights
    internal Dictionary<string, HighlightSpatialIndex> highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

    // Highlight rects per schematic per board label — built at board load, used for on-demand highlighting
    internal Dictionary<string, Dictionary<string, List<Rect>>> highlightRectsBySchematicAndLabel = new(StringComparer.OrdinalIgnoreCase);

    private double thisCurrentHighlightBlinkFactor = 1.0;

    private string? thisHoveredComponentBoardLabel;

    private readonly HashSet<string> thisSchematicsOnlySelectedBoardLabels = new(StringComparer.OrdinalIgnoreCase);

    private readonly List<Border> thisStandardLabelContainers = new();

    private readonly List<TextBlock> thisStandardLabelTextBlocks = new();

    private readonly List<ScaleTransform> thisStandardLabelScaleTransforms = new();

    private string thisLastStandardLabelVisualSignature = string.Empty;

    private string thisHighlightedBoardLabelsSignature = string.Empty;

    // ###########################################################################################
    // Which worklog entry the hover label is currently painted for, or 0 when it is in its default
    // component-hover appearance. Both UpdateSchematicsHoverUi's worklog branch and
    // ResetSchematicsHoverLabelToDefaultAppearance run on every PointerMoved frame, and each one
    // otherwise reassigns brushes on every one of those frames - the worklog side allocating two
    // SolidColorBrushes, the default side resolving three theme brushes. Repainting only when the
    // hovered entry actually changes keeps a hot path free of that.
    // ###########################################################################################
    private int thisHoverLabelPaintedWorklogEntryId;

    // ###########################################################################################
    // Cursor instances are IDisposable and were previously allocated per frame by the hover paths.
    // Two are enough for the whole tab, so they are created once and reused.
    // ###########################################################################################
    private readonly Cursor thisHandCursor = new(StandardCursorType.Hand);

    // ###########################################################################################
    // Updates UI visibility for all contextual label borders at once.
    // ###########################################################################################
    public void HideLabels()
    {
        this.SchematicsNameBorder.IsVisible = false;
        this.SchematicsRegionBorder.IsVisible = false;
        this.SchematicsHoverLabelBorder.IsVisible = false;
        this.SchematicsHoverPadBorder.IsVisible = false;
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

        this.SchematicsRegionBorder.Bind(
            Border.BorderBrushProperty,
            this.GetResourceObservable($"{colorPrefix}_Border"));

        this.SchematicsRegionLabel.Bind(
            TextBlock.ForegroundProperty,
            this.GetResourceObservable($"{colorPrefix}_Fg"));
    }

    // ###########################################################################################
    // Applies inverse scale to mapped labels so they remain standard text size regardless of zoom.
    // Uses cached editor or standard scale transforms directly when available to avoid repeated
    // child-tree scanning and LINQ allocations during pan and zoom.
    // ###########################################################################################
    private void UpdateComponentLabelsScale(double scale)
    {
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;

        if (this.thisIsLabelEditorMode && this.thisEditorLabelScaleTransforms.Count > 0)
        {
            for (int i = 0; i < this.thisEditorLabelScaleTransforms.Count; i++)
            {
                this.thisEditorLabelScaleTransforms[i].ScaleX = inverseScale;
                this.thisEditorLabelScaleTransforms[i].ScaleY = inverseScale;
            }

            return;
        }

        if (this.thisStandardLabelScaleTransforms.Count > 0)
        {
            for (int i = 0; i < this.thisStandardLabelScaleTransforms.Count; i++)
            {
                this.thisStandardLabelScaleTransforms[i].ScaleX = inverseScale;
                this.thisStandardLabelScaleTransforms[i].ScaleY = inverseScale;
            }

            return;
        }

        foreach (var child in this.SchematicsLabelsCanvas.Children)
        {
            if (child is Border container && container.RenderTransform is TransformGroup group)
            {
                var scaleTransform = group.Children.OfType<ScaleTransform>().FirstOrDefault();
                if (scaleTransform != null)
                {
                    scaleTransform.ScaleX = inverseScale;
                    scaleTransform.ScaleY = inverseScale;
                }
            }
        }
    }

    // ###########################################################################################
    // Generates floating labels exactly above relevant components based on checkbox matrix.
    // While label-editor mode is active, board labels from the working copy are shown immediately
    // so newly confirmed labels appear without waiting for Apply.
    // ###########################################################################################
    public void UpdateComponentLabels()
    {
        if (this.currentFullResBitmap == null)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
            this.ResetComponentLabelVisualCaches();
            return;
        }

        double imgWidth = this.currentFullResBitmap.PixelSize.Width;
        double imgHeight = this.currentFullResBitmap.PixelSize.Height;
        if (imgWidth <= 0 || imgHeight <= 0)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
            this.ResetComponentLabelVisualCaches();
            return;
        }

        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb == null)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
            this.ResetComponentLabelVisualCaches();
            return;
        }

        var contentRect = this.GetImageContentRect();

        double currentScale = this.schematicsMatrix.M11;
        double inverseScale = currentScale > 0 ? 1.0 / currentScale : 1.0;

        if (this.thisIsLabelEditorMode)
        {
            if (this.thisStandardLabelContainers.Count > 0)
            {
                this.SchematicsLabelsCanvas.Children.Clear();
                this.ResetStandardComponentLabelVisualCache();
            }

            var editorRows = this.thisLabelEditorWorkingHighlights
                .Where(row =>
                    string.Equals(row.SchematicName, selectedThumb.Name, StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(row.BoardLabel))
                .ToList();

            this.UpdateEditorComponentLabels(editorRows, contentRect, imgWidth, imgHeight, inverseScale);
            return;
        }

        if (this.thisEditorLabelContainers.Count > 0)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
            this.ResetEditorComponentLabelVisualCache();
        }

        if (this.CheckLabelBoard.IsChecked != true &&
            this.CheckLabelTechnical.IsChecked != true &&
            this.CheckLabelFriendly.IsChecked != true)
        {
            for (int i = 0; i < this.thisStandardLabelContainers.Count; i++)
            {
                this.thisStandardLabelContainers[i].IsVisible = false;
            }

            return;
        }

        if (this.MainWindow == null)
        {
            for (int i = 0; i < this.thisStandardLabelContainers.Count; i++)
            {
                this.thisStandardLabelContainers[i].IsVisible = false;
            }

            return;
        }

        if (!this.highlightRectsBySchematicAndLabel.TryGetValue(selectedThumb.Name, out var byLabel))
        {
            for (int i = 0; i < this.thisStandardLabelContainers.Count; i++)
            {
                this.thisStandardLabelContainers[i].IsVisible = false;
            }

            return;
        }

        var visibleItems = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<ComponentListItem>().ToList() ?? new List<ComponentListItem>();
        var selectedItems = this.MainWindow.ComponentFilterListBox.SelectedItems?.Cast<ComponentListItem>().ToList() ?? new List<ComponentListItem>();

        bool selectedOnly = this.CheckLabelSelectedOnly.IsChecked == true;
        var itemsToLoop = selectedOnly ? selectedItems : visibleItems;
        var seenLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var standardLabels = new List<(string Text, double LocalX, double LocalY)>();

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

                standardLabels.Add((labelText, localX, localY));
            }
        }

        this.UpdateStandardComponentLabels(standardLabels, inverseScale);
    }

    // ###########################################################################################
    // Rebuilds highlight indices from the selected board labels, then applies highlight visuals
    // to the main schematic and all thumbnails.
    // Also preserves schematic-only selections for components hidden by category/search filters.
    // ###########################################################################################
    public void UpdateHighlightsForComponents(List<string> boardLabels)
    {
        this.ClearVisibleBoardLabelsFromSchematicsOnlySelection();

        var effectiveBoardLabels = new HashSet<string>(
            boardLabels.Where(label => !string.IsNullOrWhiteSpace(label)),
            StringComparer.OrdinalIgnoreCase);

        foreach (string boardLabel in this.thisSchematicsOnlySelectedBoardLabels)
        {
            effectiveBoardLabels.Add(boardLabel);
        }

        this.thisHighlightedBoardLabelsSignature = string.Join(
            "\u001E",
            effectiveBoardLabels.OrderBy(label => label, StringComparer.OrdinalIgnoreCase));

        this.highlightIndexBySchematic = new(StringComparer.OrdinalIgnoreCase);

        if (effectiveBoardLabels.Count > 0)
        {
            foreach (var (schematicName, byLabel) in this.highlightRectsBySchematicAndLabel)
            {
                var rects = new List<Rect>();

                foreach (string boardLabel in effectiveBoardLabels)
                {
                    if (byLabel.TryGetValue(boardLabel, out var labelRects))
                    {
                        rects.AddRange(labelRects);
                    }
                }

                if (rects.Count > 0)
                {
                    this.highlightIndexBySchematic[schematicName] = new HighlightSpatialIndex(rects);
                }
            }
        }

        this.UpdateKiCadSelectionFromBoardLabels(effectiveBoardLabels);
        this.RefreshBlinkStateFromCurrentSelection();
        this.UpdateComponentLabels();
        this.RefreshHoveredComponentHighlightOverlay();
    }

    // ###########################################################################################
    // Returns true when any schematic selection is active that should participate in blink visuals.
    // ###########################################################################################
    public bool HasBlinkEligibleSelection()
    {
        if (this.highlightIndexBySchematic.Count > 0)
        {
            return true;
        }

        if (this.thisLockedKiCadNetNames.Count > 0)
        {
            return true;
        }

        if (this.thisSelectedImportantSignalDisplayNames.Count > 0)
        {
            return true;
        }

        string? hoveredNetName = this.GetActiveHoveredKiCadNetName();
        if (!string.IsNullOrWhiteSpace(hoveredNetName))
        {
            return true;
        }

        return false;
    }

    // ###########################################################################################
    // Recomputes blink timer state and reapplies visuals after KiCad net selection changes.
    // This keeps blinking responsive even when no component selection exists.
    // ###########################################################################################
    private void RefreshBlinkStateFromCurrentSelection()
    {
        bool hasBlinkEligibleSelection = this.HasBlinkEligibleSelection();
        bool hasComponentSelection = this.highlightIndexBySchematic.Count > 0;

        if (this.MainWindow != null)
        {
            this.MainWindow.UpdateBlinkTimer(hasBlinkEligibleSelection);
            this.ApplyHighlightVisuals(
                hasComponentSelection,
                this.MainWindow.GetCurrentBlinkFactor(hasBlinkEligibleSelection));
            return;
        }

        this.ApplyHighlightVisuals(hasComponentSelection, 1.0);
    }

    // ###########################################################################################
    // Applies current highlight visuals (including blink phase) to main schematic and thumbnails.
    // Thumbnail highlights must be regenerated when the component-selection blink phase changes
    // because their highlight overlay is baked into the rendered thumbnail bitmap.
    // ###########################################################################################
    public void ApplyHighlightVisuals(bool hasSelection, double blinkFactor)
    {
        this.thisCurrentHighlightBlinkFactor = Math.Clamp(blinkFactor, 0.0, 1.0);

        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb != null &&
            this.highlightIndexBySchematic.TryGetValue(selectedThumb.Name, out var mainIndex) &&
            this.schematicByName.TryGetValue(selectedThumb.Name, out var mainSchematic))
        {
            this.SchematicsHighlightsOverlay.HighlightIndex = mainIndex;
            this.SchematicsHighlightsOverlay.BitmapPixelSize = this.currentFullResBitmap?.PixelSize ?? new PixelSize(0, 0);
            this.SchematicsHighlightsOverlay.HighlightColor = RectGeometry.ParseColorOrDefault(mainSchematic.SchematicHighlightColor, Colors.IndianRed);
            this.SchematicsHighlightsOverlay.HighlightOpacity =
                RectGeometry.ParseOpacityOrDefault(mainSchematic.SchematicHighlightOpacity, 0.20) * this.thisCurrentHighlightBlinkFactor;
        }
        else
        {
            this.SchematicsHighlightsOverlay.HighlightIndex = null;
        }

        this.SchematicsHighlightsOverlay.InvalidateVisual();

        bool isLabelEditorModeActive = this.thisIsLabelEditorMode;
        string activeEditorSchematicName = isLabelEditorModeActive
            ? this.GetCurrentSchematicName()
            : string.Empty;

        bool hasKiCadSelection = this.HasKiCadSelectionForThumbnailDimming();
        bool hasAnyThumbnailSelection = hasSelection || hasKiCadSelection || isLabelEditorModeActive;

        string thumbnailSignature = this.BuildThumbnailHighlightSignature(hasSelection, hasKiCadSelection);

        if (!string.Equals(this.thisLastThumbnailHighlightSignature, thumbnailSignature, StringComparison.Ordinal))
        {
            this.thisLastThumbnailHighlightSignature = thumbnailSignature;

            foreach (var thumb in this.currentThumbnails)
            {
                if (thumb.BaseThumbnail == null)
                {
                    continue;
                }

                if (isLabelEditorModeActive)
                {
                    if (!ReferenceEquals(thumb.ImageSource, thumb.BaseThumbnail))
                    {
                        var old = thumb.ImageSource;
                        thumb.ImageSource = thumb.BaseThumbnail;
                        (old as IDisposable)?.Dispose();
                    }

                    bool isRelevantForEditor = string.Equals(
                        thumb.Name,
                        activeEditorSchematicName,
                        StringComparison.OrdinalIgnoreCase);

                    thumb.VisualOpacity = isRelevantForEditor ? 1.0 : 0.35;
                    thumb.IsMatchForSelection = false;
                    continue;
                }

                bool hasComponentMatch = false;
                bool hasKiCadMatch = false;

                if (this.highlightIndexBySchematic.TryGetValue(thumb.Name, out var thumbIndex) &&
                    this.schematicByName.TryGetValue(thumb.Name, out var thumbSchematic))
                {
                    hasComponentMatch = true;

                    var highlighted = CreateHighlightedThumbnail(
                        thumb.BaseThumbnail,
                        thumb.OriginalPixelSize,
                        thumbIndex,
                        thumbSchematic,
                        opacityMultiplier: this.thisCurrentHighlightBlinkFactor);

                    var old = thumb.ImageSource;
                    thumb.ImageSource = highlighted;

                    if (!ReferenceEquals(old, thumb.BaseThumbnail))
                    {
                        (old as IDisposable)?.Dispose();
                    }
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

                if (hasKiCadSelection)
                {
                    hasKiCadMatch = this.DoesSchematicContainSelectedKiCadContent(thumb.Name);
                }

                bool hasMatch = hasComponentMatch || hasKiCadMatch;
                bool isRelevantForDimming = !hasAnyThumbnailSelection || hasMatch;

                thumb.VisualOpacity = isRelevantForDimming ? 1.0 : 0.35;
                thumb.IsMatchForSelection = hasAnyThumbnailSelection && hasMatch;
            }
        }

        this.RefreshKiCadOverlay(forceImmediate: true);
    }

    // ###########################################################################################
    // Returns the hover label to its default component-hover appearance: theme colors, and no
    // worklog "#N"/state-icon parts.
    //
    // Every path that clears or rewrites the hover label must call this, because the worklog
    // branch of UpdateSchematicsHoverUi repaints the shared SchematicsHoverLabelBorder in the
    // hovered entry's category color with white text. Merely hiding the border is not enough -
    // the colors survive into whatever shows it next. That is what went wrong when only some of
    // the paths reset them: hovering a pill, moving into the KiCad net connections panel and
    // back out onto a plain component drew that component's name white-on-white in the previous
    // entry's color, with a stale "#5" and icon still beside it.
    //
    // Restores the DynamicResource bindings rather than assigning resolved brushes, the same way
    // UpdateOverlayLabels rebinds the region border. Assigning a brush would replace the binding
    // the AXAML set up with a permanent local value, so the label would keep whichever theme was
    // active at the moment a pill was first hovered and stop following ApplyConfiguredTheme.
    // ###########################################################################################
    private void ResetSchematicsHoverLabelToDefaultAppearance()
    {
        // Already in the default appearance - nothing to undo. This also keeps the rebinding below
        // off the per-frame pointer-move path.
        if (this.thisHoverLabelPaintedWorklogEntryId == 0)
        {
            return;
        }

        this.thisHoverLabelPaintedWorklogEntryId = 0;

        this.SchematicsHoverLabelIdText.IsVisible = false;
        this.SchematicsHoverLabelIconBorder.IsVisible = false;

        this.SchematicsHoverLabelBorder.Bind(
            Border.BackgroundProperty,
            this.GetResourceObservable("Schematics_ComponentHover_Bg"));

        this.SchematicsHoverLabelBorder.Bind(
            Border.BorderBrushProperty,
            this.GetResourceObservable("Schematics_ComponentHover_Border"));

        this.SchematicsHoverLabelText.Bind(
            TextBlock.ForegroundProperty,
            this.GetResourceObservable("Schematics_ComponentHover_Fg"));
    }

    // ###########################################################################################
    // Clears hover label and resets schematic cursor.
    // ###########################################################################################
    public void HideSchematicsHoverUi()
    {
        this.SetHoveredComponentBoardLabel(null);
        this.SetHoveredKiCadNet(null);
        this.thisHoveredKiCadPadNumber = null;
        this.ResetSchematicsHoverLabelToDefaultAppearance();
        this.SchematicsHoverLabelBorder.IsVisible = false;
        this.SchematicsHoverLabelText.Text = string.Empty;
        this.SchematicsHoverPadBorder.IsVisible = false;
        this.SchematicsHoverPadText.Text = string.Empty;
        this.SchematicsContainer.Cursor = Cursor.Default;

        // The worklog area's resize markers are hover state too, so they go with everything else.
        // Without this they stayed painted after the pointer left the schematic entirely - and rode
        // along on a stale entry while panning - advertising a grab affordance for an area the
        // pointer was nowhere near.
        this.ClearWorklogEntryResizeHover();

        if (this.MainWindow != null)
            this.MainWindow.isHoveringComponent = false;
    }

    // ###########################################################################################
    // Updates hover label and cursor from current pointer position.
    // The new KiCad trace calibration mode takes over hover UI so alignment work stays visually
    // clean and only the calibration box interaction feedback is shown.
    // ###########################################################################################
    private void UpdateSchematicsHoverUi(Point pointerInContainer)
    {
        // Hovering a marked worklog area shows its resize markers and a directional cursor. Checked
        // before the component/net branches below because the area is drawn ON TOP of them: the
        // pointer is visibly over the area, so offering the component underneath instead would
        // contradict what the user sees - the same reason the worklog pill wins the hover label.
        //
        // Returns false the moment the pointer leaves every area, which also clears the markers, so
        // normal hover behaviour resumes without this branch having to undo anything.
        if (this.UpdateWorklogEntryResizeHover(pointerInContainer))
        {
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            this.SetHoveredComponentBoardLabel(null);
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;

            this.ResetSchematicsHoverLabelToDefaultAppearance();
            this.SchematicsHoverLabelText.Text =
                "KiCad calibration mode - drag inside box to move, drag edges to resize, drag across to flip";
            this.SchematicsHoverLabelBorder.IsVisible = true;
            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = false;
            }

            this.UpdateKiCadTraceCalibrationCursor(pointerInContainer);
            return;
        }

        if (this.IsPointerInsideKiCadNetConnectionsPanel(pointerInContainer))
        {
            this.ClearTransientHoverForKiCadNetConnectionsPanel();
            return;
        }

        if (this.thisIsLabelEditorMode)
        {
            this.SetHoveredComponentBoardLabel(null);

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = false;
            }

            this.ResetSchematicsHoverLabelToDefaultAppearance();
            this.SchematicsHoverLabelBorder.IsVisible = false;
            this.SchematicsHoverLabelText.Text = string.Empty;
            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;
            this.UpdateLabelEditorCursor(pointerInContainer);
            return;
        }

        // A saved worklog entry's marked area is frequently drawn right on top of the component it
        // concerns, so its pill takes priority over the component hover label underneath it -
        // checked first, before the component highlight hit-test below.
        if (this.TryGetHoveredWorklogEntry(pointerInContainer, out var hoveredWorklogEntry, out var hoveredWorklogColor))
        {
            // Clears all three hover sources, as the calibration branch above does. Without the
            // KiCad half the pad box kept showing a pin from the copper underneath the pill,
            // beside a label that had already switched to describing the worklog entry.
            this.SetHoveredComponentBoardLabel(null);
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;

            // Only repaint when the hovered entry changed - this runs on every pointer-move frame.
            if (this.thisHoverLabelPaintedWorklogEntryId != hoveredWorklogEntry.Id)
            {
                this.thisHoverLabelPaintedWorklogEntryId = hoveredWorklogEntry.Id;

                // Drop the theme bindings ResetSchematicsHoverLabelToDefaultAppearance restores,
                // so they cannot overwrite the entry colors assigned just below.
                this.SchematicsHoverLabelBorder.ClearValue(Border.BackgroundProperty);
                this.SchematicsHoverLabelBorder.ClearValue(Border.BorderBrushProperty);
                this.SchematicsHoverLabelText.ClearValue(TextBlock.ForegroundProperty);

                this.SchematicsHoverLabelIdText.Text = $"#{hoveredWorklogEntry.Id}";
                this.SchematicsHoverLabelIdText.Foreground = Brushes.White;
                this.SchematicsHoverLabelIdText.IsVisible = true;
                this.SchematicsHoverLabelIconBorder.Background = new SolidColorBrush(this.ResolveWorklogStateColor(hoveredWorklogEntry.State));
                this.SchematicsHoverLabelIconBorder.IsVisible = true;
                this.SchematicsHoverLabelText.Text = hoveredWorklogEntry.Title;
                this.SchematicsHoverLabelBorder.Background = new SolidColorBrush(hoveredWorklogColor);
                this.SchematicsHoverLabelBorder.BorderBrush = Brushes.White;
                this.SchematicsHoverLabelText.Foreground = Brushes.White;
                this.SchematicsHoverLabelBorder.IsVisible = true;
            }

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = false;
            }

            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;

            this.SchematicsContainer.Cursor = this.thisHandCursor;
            return;
        }

        this.ResetSchematicsHoverLabelToDefaultAppearance();

        bool hoveringComponent = this.TryGetHoveredBoardLabel(pointerInContainer, out var hoveredBoardLabel, out var displayText);

        if (hoveringComponent)
        {
            this.SetHoveredComponentBoardLabel(hoveredBoardLabel);
            this.SchematicsHoverLabelText.Text = displayText;
            this.SchematicsHoverLabelBorder.IsVisible = true;

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = true;
            }
        }
        else
        {
            this.SetHoveredComponentBoardLabel(null);

            if (this.MainWindow != null)
            {
                this.MainWindow.isHoveringComponent = false;
            }

            this.SchematicsHoverLabelBorder.IsVisible = false;
            this.SchematicsHoverLabelText.Text = string.Empty;
        }

        this.RefreshKiCadHoverPadUi();

        if (hoveringComponent || this.SchematicsHoverPadBorder.IsVisible)
        {
            this.SchematicsContainer.Cursor = this.thisHandCursor;
        }
        else
        {
            this.SchematicsContainer.Cursor = Cursor.Default;
        }
    }

    // ###########################################################################################
    // Resolves hovered board label and the exact text shown in component selector.
    // Uses all components for the current board/region so schematic hit-testing remains available
    // even when the left-side category or search filters hide that component from the list.
    // ###########################################################################################
    private bool TryGetHoveredBoardLabel(Point pointerInContainer, out string boardLabel, out string displayText)
    {
        boardLabel = string.Empty;
        displayText = string.Empty;

        if (this.currentFullResBitmap == null || this.MainWindow == null || this.MainWindow.CurrentBoardData == null)
        {
            return false;
        }

        var selectedThumb = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;
        if (selectedThumb == null)
        {
            return false;
        }

        if (!this.highlightRectsBySchematicAndLabel.TryGetValue(selectedThumb.Name, out var byLabel))
        {
            return false;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            return false;
        }

        var localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        var contentRect = this.GetImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0 || !contentRect.Contains(localPoint))
        {
            return false;
        }

        double px = ((localPoint.X - contentRect.X) / contentRect.Width) * this.currentFullResBitmap.PixelSize.Width;
        double py = ((localPoint.Y - contentRect.Y) / contentRect.Height) * this.currentFullResBitmap.PixelSize.Height;
        var pixelPoint = new Point(px, py);

        string activeRegion = this.MainWindow.LocalRegion?.Trim() ?? string.Empty;
        var seenLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var component in this.MainWindow.CurrentBoardData.Components)
        {
            string componentBoardLabel = component.BoardLabel?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(componentBoardLabel))
            {
                continue;
            }

            string componentRegion = component.Region?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(componentRegion) &&
                !string.Equals(componentRegion, activeRegion, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (!seenLabels.Add(componentBoardLabel))
            {
                continue;
            }

            if (!byLabel.TryGetValue(componentBoardLabel, out var rects))
            {
                continue;
            }

            if (!rects.Any(r => r.Contains(pixelPoint)))
            {
                continue;
            }

            var parts = new List<string>(3);

            if (!string.IsNullOrWhiteSpace(componentBoardLabel))
            {
                parts.Add(componentBoardLabel);
            }

            if (!string.IsNullOrWhiteSpace(component.FriendlyName))
            {
                parts.Add(component.FriendlyName.Trim());
            }

            if (!string.IsNullOrWhiteSpace(component.TechnicalNameOrValue))
            {
                parts.Add(component.TechnicalNameOrValue.Trim());
            }

            boardLabel = componentBoardLabel;
            displayText = parts.Count > 0 ? string.Join(" | ", parts) : componentBoardLabel;
            return true;
        }

        return false;
    }

    // ###########################################################################################
    // Selects first component row matching board label and scrolls it into view.
    // Falls back to a schematic-only selection when the component is hidden by current filters.
    // ###########################################################################################
    private void SelectComponentByBoardLabel(string boardLabel)
    {
        if (this.MainWindow == null)
        {
            return;
        }

        var items = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<ComponentListItem>().ToList() ?? new List<ComponentListItem>();
        int index = items.FindIndex(item => string.Equals(item.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase));

        if (index >= 0)
        {
            this.thisSchematicsOnlySelectedBoardLabels.Remove(boardLabel);
            this.MainWindow.ComponentFilterListBox.Selection.Select(index);
            this.MainWindow.ComponentFilterListBox.ScrollIntoView(items[index]);
            return;
        }

        if (this.thisSchematicsOnlySelectedBoardLabels.Add(boardLabel))
        {
            this.RefreshHighlightsFromCurrentComponentSelection();
        }
    }

    // ###########################################################################################
    // Deselects all component rows that match the given board label.
    // Also removes any schematic-only hidden selection for the same board label.
    // ###########################################################################################
    private void DeselectComponentByBoardLabel(string boardLabel)
    {
        bool removedHiddenSelection = this.thisSchematicsOnlySelectedBoardLabels.Remove(boardLabel);

        if (this.MainWindow == null)
        {
            if (removedHiddenSelection)
            {
                this.RefreshHighlightsFromCurrentComponentSelection();
            }

            return;
        }

        var items = this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<ComponentListItem>().ToList() ?? new List<ComponentListItem>();
        bool removedVisibleSelection = false;

        for (int i = 0; i < items.Count; i++)
        {
            if (string.Equals(items[i].BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase))
            {
                this.MainWindow.ComponentFilterListBox.Selection.Deselect(i);
                removedVisibleSelection = true;
            }
        }

        if (removedHiddenSelection && !removedVisibleSelection)
        {
            this.RefreshHighlightsFromCurrentComponentSelection();
        }
    }

    // ###########################################################################################
    // Toggles selection for a component board label.
    // ###########################################################################################
    private void ToggleComponentSelectionByBoardLabel(string boardLabel)
    {
        if (this.IsComponentBoardLabelSelected(boardLabel))
        {
            this.DeselectComponentByBoardLabel(boardLabel);
        }
        else
        {
            this.SelectComponentByBoardLabel(boardLabel);
        }
    }

    // ###########################################################################################
    // Updates the currently hovered component board label and rebuilds the hover-only overlay.
    // Only forces an immediate KiCad overlay refresh when the effective hover-driven KiCad state
    // actually changes.
    // ###########################################################################################
    private void SetHoveredComponentBoardLabel(string? boardLabel)
    {
        string? normalizedBoardLabel = string.IsNullOrWhiteSpace(boardLabel)
            ? null
            : boardLabel.Trim();

        string? previousBoardLabel = string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel)
            ? null
            : this.thisHoveredComponentBoardLabel.Trim();

        if (string.Equals(previousBoardLabel, normalizedBoardLabel, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        bool shouldRefreshKiCadOverlay = false;

        if (this.IsBoardMarkPin1OnSelectedComponentEnabled())
        {
            shouldRefreshKiCadOverlay = true;
        }

        if (!shouldRefreshKiCadOverlay && this.IsBoardShowTracesOnSelectedComponentEnabled())
        {
            var previousNets = this.BuildKiCadNormalizedNetNamesForSingleReference(previousBoardLabel);
            var nextNets = this.BuildKiCadNormalizedNetNamesForSingleReference(normalizedBoardLabel);

            if (!ViewportMath.SetEqualsOrdinalIgnoreCase(previousNets, nextNets))
            {
                shouldRefreshKiCadOverlay = true;
            }
        }

        this.thisHoveredComponentBoardLabel = normalizedBoardLabel;
        this.RefreshHoveredComponentHighlightOverlay();

        if (shouldRefreshKiCadOverlay)
        {
            this.RefreshKiCadOverlay(forceImmediate: true);
        }
    }

    // ###########################################################################################
    // Rebuilds the hover-only highlight overlay for the active schematic component.
    // Suppresses hover overlay when the same component is already selected.
    // ###########################################################################################
    private void RefreshHoveredComponentHighlightOverlay()
    {
        this.SchematicsHoverHighlightsOverlay.BitmapPixelSize = this.currentFullResBitmap?.PixelSize ?? new PixelSize(0, 0);
        this.SchematicsHoverHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        if (this.currentFullResBitmap == null || string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel))
        {
            this.SchematicsHoverHighlightsOverlay.HighlightIndex = null;
            this.SchematicsHoverHighlightsOverlay.InvalidateVisual();
            return;
        }

        string schematicName = this.GetCurrentSchematicName();
        if (string.IsNullOrWhiteSpace(schematicName) ||
            !this.highlightRectsBySchematicAndLabel.TryGetValue(schematicName, out var byLabel) ||
            !byLabel.TryGetValue(this.thisHoveredComponentBoardLabel, out var rects) ||
            rects.Count == 0)
        {
            this.SchematicsHoverHighlightsOverlay.HighlightIndex = null;
            this.SchematicsHoverHighlightsOverlay.InvalidateVisual();
            return;
        }

        bool isAlreadySelected = this.IsComponentBoardLabelSelected(this.thisHoveredComponentBoardLabel);

        if (isAlreadySelected)
        {
            this.SchematicsHoverHighlightsOverlay.HighlightIndex = null;
            this.SchematicsHoverHighlightsOverlay.InvalidateVisual();
            return;
        }

        Color highlightColor = Colors.IndianRed;
        double highlightOpacity = 0.20;

        if (this.schematicByName.TryGetValue(schematicName, out var schematic))
        {
            highlightColor = RectGeometry.ParseColorOrDefault(schematic.SchematicHighlightColor, Colors.IndianRed);
            highlightOpacity = RectGeometry.ParseOpacityOrDefault(schematic.SchematicHighlightOpacity, 0.20);
        }

        this.SchematicsHoverHighlightsOverlay.HighlightColor = highlightColor;
        this.SchematicsHoverHighlightsOverlay.HighlightOpacity = highlightOpacity;
        this.SchematicsHoverHighlightsOverlay.HighlightIndex = new HighlightSpatialIndex(rects);
        this.SchematicsHoverHighlightsOverlay.InvalidateVisual();
    }

    // ###########################################################################################
    // Clears all cached component-label visual pools so stale controls are not retained when the
    // viewer is reset or when label rendering switches between standard and editor modes.
    // ###########################################################################################
    private void ResetComponentLabelVisualCaches()
    {
        this.thisEditorLabelContainers.Clear();
        this.thisEditorLabelTextBlocks.Clear();
        this.thisEditorLabelScaleTransforms.Clear();
        this.thisLastEditorLabelVisualSignature = string.Empty;

        this.thisStandardLabelContainers.Clear();
        this.thisStandardLabelTextBlocks.Clear();
        this.thisStandardLabelScaleTransforms.Clear();
        this.thisLastStandardLabelVisualSignature = string.Empty;
    }

    // ###########################################################################################
    // Clears the cached standard component-label visual pool so stale controls are not retained
    // when normal label rendering is disabled or the editor takes ownership of the labels canvas.
    // ###########################################################################################
    private void ResetStandardComponentLabelVisualCache()
    {
        this.thisStandardLabelContainers.Clear();
        this.thisStandardLabelTextBlocks.Clear();
        this.thisStandardLabelScaleTransforms.Clear();
        this.thisLastStandardLabelVisualSignature = string.Empty;
    }

    // ###########################################################################################
    // Ensures the reusable standard component-label visual pool contains at least the requested
    // number of controls. Labels are created once and then reused on later refreshes.
    // ###########################################################################################
    private void EnsureStandardComponentLabelVisualPoolSize(int requiredCount)
    {
        while (this.thisStandardLabelContainers.Count < requiredCount)
        {
            var textBlock = new TextBlock
            {
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                TextAlignment = TextAlignment.Center
            };
            textBlock.Bind(TextBlock.ForegroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Fg"));

            var innerBorder = new Border
            {
                BorderThickness = new Avalonia.Thickness(1),
                CornerRadius = new Avalonia.CornerRadius(4),
                Padding = new Avalonia.Thickness(6, 4),
                Child = textBlock
            };
            innerBorder.Bind(Border.BackgroundProperty, this.GetResourceObservable("Schematics_ComponentLabel_Bg"));
            innerBorder.Bind(Border.BorderBrushProperty, this.GetResourceObservable("Schematics_ComponentLabel_Border"));

            var scaleTransform = new ScaleTransform(1.0, 1.0);
            var transformGroup = new TransformGroup();
            transformGroup.Children.Add(scaleTransform);

            var container = new Border
            {
                IsHitTestVisible = false,
                IsVisible = false,
                Child = innerBorder,
                RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative),
                RenderTransform = transformGroup
            };

            this.thisStandardLabelContainers.Add(container);
            this.thisStandardLabelTextBlocks.Add(textBlock);
            this.thisStandardLabelScaleTransforms.Add(scaleTransform);
            this.SchematicsLabelsCanvas.Children.Add(container);
        }
    }

    // ###########################################################################################
    // Builds a lightweight signature for the currently visible standard labels so text updates
    // only occur when the visible label set actually changes.
    // ###########################################################################################
    private static string BuildStandardComponentLabelVisualSignature(IReadOnlyList<(string Text, double LocalX, double LocalY)> labels)
    {
        if (labels.Count == 0)
        {
            return string.Empty;
        }

        var parts = new string[labels.Count];

        for (int i = 0; i < labels.Count; i++)
        {
            parts[i] = labels[i].Text ?? string.Empty;
        }

        return string.Join("\u001F", parts);
    }

    // ###########################################################################################
    // Updates the reusable standard component-label controls without clearing and rebuilding the
    // entire canvas on every refresh. This removes the same allocation churn pattern that the
    // editor labels originally suffered from.
    // ###########################################################################################
    private void UpdateStandardComponentLabels(
        IReadOnlyList<(string Text, double LocalX, double LocalY)> labels,
        double inverseScale)
    {
        if (this.thisStandardLabelContainers.Count == 0 && this.SchematicsLabelsCanvas.Children.Count > 0)
        {
            this.SchematicsLabelsCanvas.Children.Clear();
        }

        this.EnsureStandardComponentLabelVisualPoolSize(labels.Count);

        string newSignature = BuildStandardComponentLabelVisualSignature(labels);
        bool textChanged = !string.Equals(
            this.thisLastStandardLabelVisualSignature,
            newSignature,
            StringComparison.Ordinal);

        if (textChanged)
        {
            for (int i = 0; i < labels.Count; i++)
            {
                this.thisStandardLabelTextBlocks[i].Text = labels[i].Text;
            }

            this.thisLastStandardLabelVisualSignature = newSignature;
        }

        for (int i = 0; i < labels.Count; i++)
        {
            var label = labels[i];
            var container = this.thisStandardLabelContainers[i];
            var scaleTransform = this.thisStandardLabelScaleTransforms[i];

            scaleTransform.ScaleX = inverseScale;
            scaleTransform.ScaleY = inverseScale;

            bool needsMeasure =
                textChanged ||
                !container.IsVisible ||
                container.DesiredSize.Width <= 0 ||
                container.DesiredSize.Height <= 0;

            container.IsVisible = true;

            if (needsMeasure)
            {
                container.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            }

            Size desiredSize = container.DesiredSize;

            Canvas.SetLeft(container, label.LocalX - (desiredSize.Width / 2.0));
            Canvas.SetTop(container, label.LocalY - (desiredSize.Height / 2.0));
        }

        for (int i = labels.Count; i < this.thisStandardLabelContainers.Count; i++)
        {
            this.thisStandardLabelContainers[i].IsVisible = false;
        }
    }

    // ###########################################################################################
    // Clears any schematic-only selections that were created for components hidden by the current
    // category or search filters.
    // ###########################################################################################
    internal void ClearSchematicsOnlySelectedComponents()
    {
        this.thisSchematicsOnlySelectedBoardLabels.Clear();
    }

    // ###########################################################################################
    // Returns true when the board label is selected either through the visible component list or
    // through the schematic-only hidden-selection cache.
    // ###########################################################################################
    private bool IsComponentBoardLabelSelected(string boardLabel)
    {
        if (string.IsNullOrWhiteSpace(boardLabel))
        {
            return false;
        }

        if (this.thisSchematicsOnlySelectedBoardLabels.Contains(boardLabel))
        {
            return true;
        }

        return this.MainWindow?.ComponentFilterListBox.SelectedItems?
            .Cast<ComponentListItem>()
            .Any(item => string.Equals(item.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase)) ?? false;
    }

    // ###########################################################################################
    // Removes schematic-only selections for any labels that are currently visible in the filtered
    // component list so the visible list remains the active source of truth for those rows.
    // ###########################################################################################
    private void ClearVisibleBoardLabelsFromSchematicsOnlySelection()
    {
        if (this.MainWindow == null)
        {
            return;
        }

        foreach (var item in this.MainWindow.ComponentFilterListBox.ItemsSource?.Cast<ComponentListItem>() ?? Enumerable.Empty<ComponentListItem>())
        {
            if (!string.IsNullOrWhiteSpace(item.BoardLabel))
            {
                this.thisSchematicsOnlySelectedBoardLabels.Remove(item.BoardLabel);
            }
        }
    }

    // ###########################################################################################
    // Rebuilds highlights from the visible component-list selection while preserving any extra
    // schematic-only selections for components hidden by category or search filters.
    // ###########################################################################################
    private void RefreshHighlightsFromCurrentComponentSelection()
    {
        var boardLabels = this.MainWindow?.ComponentFilterListBox.SelectedItems?
            .Cast<ComponentListItem>()
            .Select(item => item.BoardLabel)
            .Where(label => !string.IsNullOrWhiteSpace(label))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList() ?? new List<string>();

        this.UpdateHighlightsForComponents(boardLabels);
    }
}