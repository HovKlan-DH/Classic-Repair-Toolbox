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
// The schematic thumbnail list: loading and sorting, selection changes, thumbnail bitmap
// generation, and the drag-and-drop reordering interaction.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Thumbnails
    internal ObservableCollection<SchematicThumbnail> currentThumbnails = new();

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

    private double thisThumbnailLastPointerYInList = double.NaN;

    private double thisThumbnailDragGhostFixedX;

    private bool thisSuppressThumbnailSelectionChanged;

    // Lets a headless test put the tab into the mid-drag state SelectSchematicByName has to refuse
    // to act in, without driving a whole pointer drag-reorder through the list. Set by the drag
    // handlers in the running app, never by anything else - the same test-seam pattern the
    // LabelEditor and Worklog parts use.
    internal bool SuppressThumbnailSelectionChangedForTests
    {
        get => this.thisSuppressThumbnailSelectionChanged;
        set => this.thisSuppressThumbnailSelectionChanged = value;
    }

    private PointerPressedEventArgs? thisThumbnailDragStartEventArgs;

    private string thisLastThumbnailHighlightSignature = string.Empty;

    // ###########################################################################################
    // Loads the full-resolution image for the selected thumbnail and sets up the highlight overlay.
    // ###########################################################################################
    private async void OnSchematicsThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (this.thisSuppressThumbnailSelectionChanged)
            return;

        this.CancelWorklogEntryMode();

        this.UpdateOverlayLabels();

        this.RebuildImportantSignalsPanel();
        this.UpdateKiCadNetConnectionsPanel(Array.Empty<string>());

        this.fullResLoadCts?.Cancel();
        this.fullResLoadCts = new CancellationTokenSource();
        var cts = this.fullResLoadCts;

        var selected = this.SchematicsThumbnailList.SelectedItem as SchematicThumbnail;

        this.thisHoveredComponentBoardLabel = null;

        this.ResetKiCadHoverHitTestThrottle();

        this.SchematicsImage.Source = null;
        this.SchematicsMissingImageText.IsVisible = false; // Hide while loading
        this.schematicsMatrix = Matrix.Identity;
        ((MatrixTransform)this.SchematicsImage.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;
        ((MatrixTransform)this.SchematicsHoverHighlightsOverlay.RenderTransform!).Matrix = this.schematicsMatrix;

        this.SchematicsHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        this.SchematicsHoverHighlightsOverlay.HighlightIndex = null;
        this.SchematicsHoverHighlightsOverlay.BitmapPixelSize = new PixelSize(0, 0);
        this.SchematicsHoverHighlightsOverlay.ViewMatrix = this.schematicsMatrix;

        // Save the newly selected schematic for this board
        var boardKey = this.MainWindow?.GetCurrentBoardKey();
        if (!string.IsNullOrEmpty(boardKey) && selected != null)
        {
            UserSettings.SetLastSchematicForBoard(boardKey, selected.Name);
        }

        this.RestoreBoardSettings(boardKey ?? string.Empty);

        if (selected == null || string.IsNullOrEmpty(selected.ImageFilePath))
            return;

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

        if (cts.Token.IsCancellationRequested || !ReferenceEquals(cts, this.fullResLoadCts))
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
            this.SchematicsHoverHighlightsOverlay.BitmapPixelSize = bitmap.PixelSize;

            if (this.highlightIndexBySchematic.TryGetValue(selected.Name, out var index) &&
                this.schematicByName.TryGetValue(selected.Name, out var schematic))
            {
                this.SchematicsHighlightsOverlay.HighlightIndex = index;
                this.SchematicsHighlightsOverlay.HighlightColor = RectGeometry.ParseColorOrDefault(schematic.SchematicHighlightColor, Colors.IndianRed);
                this.SchematicsHighlightsOverlay.HighlightOpacity = RectGeometry.ParseOpacityOrDefault(schematic.SchematicHighlightOpacity, 0.20);
            }
        }
        else
        {
            this.SchematicsMissingImageText.IsVisible = true;
        }

        this.SchematicsHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHighlightsOverlay.InvalidateVisual();
        this.SchematicsHoverHighlightsOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsHoverHighlightsOverlay.InvalidateVisual();

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
            this.RefreshHoveredComponentHighlightOverlay();
            this.RefreshKiCadOverlay();

            // RefreshWorklogEntriesListOverlay bails out while currentFullResBitmap is still null,
            // so a "Show worklogs" checkbox that starts checked (at app launch, or after switching
            // boards) needs re-applying once this schematic's bitmap has actually finished loading -
            // otherwise the overlay silently never appears until the checkbox is toggled off and on.
            this.RefreshWorklogEntriesListOverlay();
        });
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
            HighlightColor = RectGeometry.ParseColorOrDefault(schematic.ThumbnailHighlightColor, Colors.IndianRed),
            HighlightOpacity = RectGeometry.ParseOpacityOrDefault(schematic.ThumbnailHighlightOpacity, 0.20) * Math.Clamp(opacityMultiplier, 0.0, 1.0),
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
    // Returns true when explicit KiCad trace selections should participate in thumbnail dimming.
    // Hover-only KiCad nets are excluded so thumbnails do not flicker while moving the pointer.
    // Component-derived net names are also excluded because component presence must be validated by
    // the actual schematic image/component data, not by shared net names on other pages.
    // ###########################################################################################
    private bool HasKiCadSelectionForThumbnailDimming()
    {
        return this.BuildKiCadThumbnailDimmingNetNames().Count > 0;
    }

    // ###########################################################################################
    // Loads thumbnails, applies saved user order, and removes stale saved entries automatically.
    // ###########################################################################################
    public void LoadSortedThumbnails(string boardKey, List<SchematicThumbnail> rawList)
    {
        this.RestoreBoardSettings(boardKey);

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
    // Selects a schematic by name in the thumbnail list, if one by that name is currently loaded -
    // called from TabWorkbooks (SelectWorkbook/SelectSchematic) so that picking a workbook or a
    // schematic on the Workbooks tab makes the matching schematic image the one waiting on the
    // Schematics tab, without switching tabs to show it. Setting SelectedItem runs the exact same
    // OnSchematicsThumbnailSelectionChanged path a click on the thumbnail would, so the full-res
    // image, overlays and saved "last schematic for this board" all follow normally.
    //
    // A no-op when the name is not found (board not loaded yet, or the workbook references a
    // schematic that was renamed/removed) or already selected - same "nothing to do" guard
    // OnSchematicsThumbnailSelectionChanged's own callers rely on elsewhere.
    // ###########################################################################################
    public void SelectSchematicByName(string schematicName)
    {
        if (string.IsNullOrWhiteSpace(schematicName))
            return;

        // A thumbnail drag-reorder is in progress (the flag is set on drag start and cleared only
        // on release), during which OnSchematicsThumbnailSelectionChanged early-returns. Assigning
        // SelectedItem here would move the list's highlight while the image, overlays and
        // currentFullResBitmap stayed on the OLD schematic - and the drag's own release handler
        // reassigns SelectedItem to the dragged thumbnail anyway, so the assignment would be lost.
        // Dropping the request is right: the user is reordering thumbnails, not asking for a
        // different one, and the caller is a cross-tab convenience rather than a command.
        if (this.thisSuppressThumbnailSelectionChanged)
            return;

        if (this.SchematicsThumbnailList.SelectedItem is SchematicThumbnail current &&
            string.Equals(current.Name, schematicName, StringComparison.OrdinalIgnoreCase))
            return;

        var match = this.currentThumbnails.FirstOrDefault(t =>
            string.Equals(t.Name, schematicName, StringComparison.OrdinalIgnoreCase));

        if (match != null)
        {
            this.SchematicsThumbnailList.SelectedItem = match;
        }
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

        this.thisThumbnailDragStartEventArgs = e;
        this.thisThumbnailDragStartPoint = e.GetPosition(this);
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
        this.ShowThumbnailDropPlaceholder(this.thisDraggedThumbnailOriginalIndex);
        this.ShowThumbnailDragGhost(this.thisDraggedThumbnail, pointerInList);

        if (this.thisThumbnailDragStartEventArgs == null)
        {
            this.RestoreDraggedThumbnail();
            this.HideThumbnailDragGhost();
            this.ClearThumbnailDropPlaceholder();
            this.thisDraggedThumbnail = null;
            this.thisDraggedThumbnailOriginalIndex = -1;
            this.thisDraggedThumbnailWasSelected = false;
            this.thisThumbnailLastPointerYInList = double.NaN;
            this.thisThumbnailDragStartEventArgs = null;
            return;
        }

        var dragData = new DataTransfer();

        var effect = await DragDrop.DoDragDropAsync(
            this.thisThumbnailDragStartEventArgs,
            dragData,
            DragDropEffects.Move);

        if (effect != DragDropEffects.Move && this.thisDraggedThumbnail != null)
        {
            this.RestoreDraggedThumbnail();
        }

        this.HideThumbnailDragGhost();
        this.ClearThumbnailDropPlaceholder();
        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
        this.thisDraggedThumbnailWasSelected = false;
        this.thisThumbnailLastPointerYInList = double.NaN;
        this.thisThumbnailDragStartEventArgs = null;

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
        this.thisThumbnailDragStartEventArgs = null;
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
            return;
        }

        if (existingIndex >= 0)
        {
            this.currentThumbnails.RemoveAt(existingIndex);
        }

        insertIndex = Math.Clamp(insertIndex, 0, this.currentThumbnails.Count);
        this.currentThumbnails.Insert(insertIndex, this.thisThumbnailDropPlaceholder);
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
        this.thisThumbnailLastPointerYInList = double.NaN;

        if (this.thisDraggedThumbnailWasSelected)
        {
            this.SchematicsThumbnailList.SelectedItem = this.thisDraggedThumbnail;
        }

        this.thisSuppressThumbnailSelectionChanged = false;
        this.thisThumbnailDragStartEventArgs = null;
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
        if (this.thisDraggedThumbnail == null)
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
        if (this.thisDraggedThumbnail == null)
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

        if (this.thisDraggedThumbnailWasSelected)
        {
            this.SchematicsThumbnailList.SelectedItem = this.thisDraggedThumbnail;
        }

        this.HideThumbnailDragGhost();
        this.SaveCurrentThumbnailOrder();

        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
        this.thisDraggedThumbnailWasSelected = false;
        this.thisThumbnailLastPointerYInList = double.NaN;
        this.thisSuppressThumbnailSelectionChanged = false;
        this.thisThumbnailDragStartEventArgs = null;

        e.Handled = true;
    }

    // ###########################################################################################
    // Builds a stable signature for thumbnail highlight state so thumbnails only rebuild when the
    // actual rendered state changes. Component blink phase is included because thumbnail highlight
    // overlays are baked into the rendered bitmap.
    // ###########################################################################################
    private string BuildThumbnailHighlightSignature(bool hasComponentSelection, bool hasKiCadSelection)
    {
        // Keyed by the actual set of highlighted board labels rather than by which schematics have
        // at least one highlight, since two different label sets can highlight the same schematics.
        string componentPart = hasComponentSelection
            ? this.thisHighlightedBoardLabelsSignature
            : string.Empty;

        string componentBlinkPart = hasComponentSelection
            ? this.thisCurrentHighlightBlinkFactor.ToString("0.###", CultureInfo.InvariantCulture)
            : string.Empty;

        var explicitKiCadNets = this.BuildKiCadThumbnailDimmingNetNames();

        string explicitKiCadNetPart = hasKiCadSelection
            ? string.Join(
                "\u001E",
                explicitKiCadNets.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            : string.Empty;

        string importantSignalPart = this.thisSelectedImportantSignalDisplayNames.Count > 0
            ? string.Join(
                "\u001E",
                this.thisSelectedImportantSignalDisplayNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            : string.Empty;

        string lockedNetPart = hasKiCadSelection
            ? string.Join(
                "\u001E",
                this.thisLockedKiCadNetNames.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))
            : string.Empty;

        string labelEditorModePart = this.thisIsLabelEditorMode ? "LabelEditor" : string.Empty;
        string labelEditorSchematicPart = this.thisIsLabelEditorMode
            ? this.GetCurrentSchematicName()
            : string.Empty;

        return string.Join(
            "\u001F",
            componentPart,
            componentBlinkPart,
            explicitKiCadNetPart,
            importantSignalPart,
            lockedNetPart,
            labelEditorModePart,
            labelEditorSchematicPart);
    }

    // ###########################################################################################
    // Builds the explicit KiCad net-name set that is allowed to participate in thumbnail dimming.
    // Component-derived net names are intentionally excluded so a selected component does not make
    // unrelated schematic pages look relevant just because they share one of the same net names.
    // ###########################################################################################
    private HashSet<string> BuildKiCadThumbnailDimmingNetNames()
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        result.UnionWith(this.BuildSelectedImportantSignalNetNames().Where(netName => !string.IsNullOrWhiteSpace(netName)));

        result.UnionWith(this.thisLockedKiCadNetNames.Where(netName => !string.IsNullOrWhiteSpace(netName)));

        return result;
    }
}