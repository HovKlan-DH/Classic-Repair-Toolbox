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
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// The schematic thumbnails: loading and sorting, the selection and the label editor's lock every
// thumbnail panel follows, selection changes, thumbnail bitmap generation, and saving the order a
// drag leaves. The panel the user sees - in this tab, in the detached window and beside the
// fullscreen view - and its drag-to-reorder are SchematicsThumbnailGallery.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Thumbnails
    internal ObservableCollection<SchematicThumbnail> currentThumbnails = new();

    // ###########################################################################################
    // The selected schematic, as every thumbnail gallery shows it - in this tab, the detached
    // window and beside the fullscreen view (code review, 2026-10-15). Setting it raises
    // SelectedThumbnailChanged, which OnSchematicsThumbnailSelectionChanged (from Initialize) and
    // every gallery on screen follow. A gallery asks for another one through
    // SelectThumbnailFromGallery, and never passes on what its own ListBox does by itself: each
    // gallery once copied its ListBox's selection here and back, and every "nothing selected" its
    // ListBox produced by itself (a tile lifted out for a drag, a Ctrl+click) blanked the view.
    //
    // A field, not a list: it was held in a never-shown ListBox bound to the thumbnails, so lifting
    // the selected tile out for a drag emptied it - for the whole drag the tab had no schematic
    // (its name and component labels were cleared by anything that redrew them) and the drop had
    // to select it again. What that list did on purpose, OnCurrentThumbnailsChanged does.
    // ###########################################################################################
    private SchematicThumbnail? thisSelectedThumbnail;

    internal SchematicThumbnail? SelectedThumbnail
    {
        get => this.thisSelectedThumbnail;
        set
        {
            if (ReferenceEquals(this.thisSelectedThumbnail, value))
                return;

            this.thisSelectedThumbnail = value;
            this.SelectedThumbnailChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    internal event EventHandler? SelectedThumbnailChanged;

    // ###########################################################################################
    // Keeps the selection among the thumbnails: emptying them for a board load, or removing the
    // selected one, leaves nothing selected - which clears the schematic view. A drag is the
    // exception: it lifts its tile out and puts it back, and the schematic on screen stays.
    // ###########################################################################################
    private void OnCurrentThumbnailsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (this.thisSelectedThumbnail == null || this.currentThumbnails.Contains(this.thisSelectedThumbnail))
            return;

        if (this.SuppressThumbnailSelectionChanged && e.Action != NotifyCollectionChangedAction.Reset)
            return;

        this.SelectedThumbnail = null;
    }

    // The label editor's lock on picking another schematic from the thumbnails, wherever they are
    // shown. Set by UpdateLabelEditorLockState; every gallery follows it, and drops a drag on it.
    internal bool IsThumbnailNavigationLocked { get; private set; }

    internal event EventHandler? ThumbnailNavigationLockChanged;

    private void SetThumbnailNavigationLocked(bool isLocked)
    {
        if (this.IsThumbnailNavigationLocked == isLocked)
            return;

        this.IsThumbnailNavigationLocked = isLocked;
        this.ThumbnailNavigationLockChanged?.Invoke(this, EventArgs.Empty);
    }

    // Test-only, like SuppressThumbnailSelectionChangedForTests below: the lock without a whole
    // label editor session behind it.
    internal bool ThumbnailNavigationLockedForTests
    {
        get => this.IsThumbnailNavigationLocked;
        set => this.SetThumbnailNavigationLocked(value);
    }

    // ###########################################################################################
    // A thumbnail picked in a gallery - clicked, tapped, or reached with the arrow keys. Refused
    // while the label editor holds the lock, while a reorder holds the guard (the view would stay
    // on the old schematic - see SelectSchematicByName), for the drop placeholder, which names no
    // schematic, and for a thumbnail no longer on the board: a press held over a board load ends
    // on one, and picking it left the new board with nothing selected.
    // ###########################################################################################
    internal void SelectThumbnailFromGallery(SchematicThumbnail thumbnail)
    {
        if (this.IsThumbnailNavigationLocked || this.SuppressThumbnailSelectionChanged || thumbnail.IsDropPlaceholder)
            return;

        if (!this.currentThumbnails.Contains(thumbnail))
            return;

        this.SelectedThumbnail = thumbnail;
    }

    // ###########################################################################################
    // Ends a drag, or a press that may become one, in every gallery - this tab's, the detached
    // window's and the fullscreen stand-in's - without saving an order. Called before the
    // thumbnails are emptied for a board load (code review, 2026-10-15): only this tab's own
    // column's drag used to be ended, and a drag in the fullscreen stand-in kept the guard over the
    // load, so the new board's schematic never loaded.
    // ###########################################################################################
    internal event EventHandler? ThumbnailDragsCancelRequested;

    internal void CancelThumbnailDrags()
    {
        this.ThumbnailDragsCancelRequested?.Invoke(this, EventArgs.Empty);
    }

    // The shared "a reorder is in progress, refuse to change schematic" guard. Set by whichever
    // gallery is dragging (SchematicsThumbnailGallery.BeginTileDrag/EndTileDrag) - this tab's own,
    // the detached window's or the fullscreen stand-in's share one flag rather than each carrying
    // their own, so SelectSchematicByName has a single thing to check whichever surface the drag
    // is happening on.
    internal bool SuppressThumbnailSelectionChanged { get; set; }

    // Test-only alias, so a headless test can put the tab into that mid-drag state without driving
    // a whole pointer drag-reorder through the list. Kept SEPARATE from the property above rather
    // than being the only name for it: production code setting a "...ForTests" member made the
    // seam's own "never set by anything else" contract false, and any later cleanup that trusted
    // it would have silently broken the detached drag's cross-tab guard.
    internal bool SuppressThumbnailSelectionChangedForTests
    {
        get => this.SuppressThumbnailSelectionChanged;
        set => this.SuppressThumbnailSelectionChanged = value;
    }

    private string thisLastThumbnailHighlightSignature = string.Empty;

    // ###########################################################################################
    // Loads the full-resolution image for the selected thumbnail and sets up the highlight overlay.
    // ###########################################################################################
    private async void OnSchematicsThumbnailSelectionChanged(object? sender, EventArgs e)
    {
        // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
        // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
        try
        {
            if (this.SuppressThumbnailSelectionChanged)
                return;

            this.CancelWorklogEntryMode();

            // The last picture's traces let go AT ONCE - a point being dragged dropped, never saved -
            // and the new one's read with its picture (TabSchematics.ViewTools.cs, code review, 2026-10-29).
            this.LetGoOfTraces();

            // A picture tool waiting for this picture starts once its image is read (TabSchematics.PictureTools.cs).
            this.thisImageSettledFor = null;

            this.UpdateOverlayLabels();

            this.RebuildImportantSignalsPanel();
            this.UpdateKiCadNetConnectionsPanel(Array.Empty<string>());

            this.fullResLoadCts?.Cancel();
            this.fullResLoadCts = new CancellationTokenSource();
            var cts = this.fullResLoadCts;

            var selected = this.SelectedThumbnail;

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
            {
                // No picture to show: no drawing on it either (code review, 2026-10-29).
                this.ShowViewToolsForPicture(false);
                this.ImageSettled(selected?.Name, shown: false);
                return;
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

            if (cts.Token.IsCancellationRequested || !ReferenceEquals(cts, this.fullResLoadCts))
            {
                bitmap?.Dispose();
                return;
            }

            this.currentFullResBitmap?.Dispose();
            this.currentFullResBitmap = bitmap;
            this.SchematicsImage.Source = bitmap;

            // "Draw trace" and the one-time tip follow the picture (TabSchematics.ViewTools.cs).
            this.ShowViewToolsForPicture(bitmap != null);

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
            this.ReadTracesFor(boardKey, selected.Name);

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

                // Laid out and drawn: a picture tool waiting for this picture can start on it.
                if (ReferenceEquals(this.SelectedThumbnail, selected))
                    this.ImageSettled(selected.Name, shown: bitmap != null);
            });
        }
        catch (Exception ex)
        {
            Logger.Critical($"OnSchematicsThumbnailSelectionChanged failed - [{ex}]");
        }
    }

    // ###########################################################################################
    // Renders a thumbnail's highlight rectangles ALONE, on transparent, at the base thumbnail's
    // pixel size - the layer every thumbnail template draws over the base image
    // (SchematicThumbnail.HighlightImage). On the UI thread: it lays out a control.
    //
    // It used to composite the base image too, with the blink factor baked into the highlight
    // opacity, so each blink tick rendered every highlighted thumbnail again (code review,
    // 2026-10-13). Drawn over the same image at the same size, it looks as the composite did; the
    // blink is the layer's opacity now (ApplyBlinkPhase).
    // ###########################################################################################
    public static RenderTargetBitmap CreateThumbnailHighlightLayer(
        IImage baseThumbnail, PixelSize originalPixelSize,
        HighlightSpatialIndex index, BoardSchematicEntry schematic)
    {
        int tw = 1, th = 1;
        if (baseThumbnail is Bitmap bmp)
        {
            tw = bmp.PixelSize.Width;
            th = bmp.PixelSize.Height;
        }

        var overlay = new SchematicHighlightsOverlay
        {
            HighlightIndex = index,
            BitmapPixelSize = originalPixelSize,
            ViewMatrix = Matrix.Identity,
            HighlightColor = RectGeometry.ParseColorOrDefault(schematic.ThumbnailHighlightColor, Colors.IndianRed),
            HighlightOpacity = RectGeometry.ParseOpacityOrDefault(schematic.ThumbnailHighlightOpacity, 0.20),
            IsHitTestVisible = false
        };

        var root = new Grid();
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

        // No drag may run on across the old tiles being replaced (see CancelThumbnailDrags).
        this.CancelThumbnailDrags();

        this.currentThumbnails.Clear();
        foreach (var thumbnail in orderedList)
        {
            this.currentThumbnails.Add(thumbnail);
        }
    }

    // ###########################################################################################
    // Selects a schematic by name among the thumbnails, if one by that name is currently loaded -
    // called from TabWorkbooks (SelectWorkbook/SelectSchematic) so that picking a workbook or a
    // schematic on the Workbooks tab makes the matching schematic image the one waiting on the
    // Schematics tab, without switching tabs to show it. Setting SelectedThumbnail runs the exact
    // same OnSchematicsThumbnailSelectionChanged path a click on the thumbnail would, so the
    // full-res image, overlays and saved "last schematic for this board" all follow normally.
    //
    // A no-op when the name is not found (board not loaded yet, or the workbook references a
    // schematic that was renamed/removed) or already selected - same "nothing to do" guard
    // OnSchematicsThumbnailSelectionChanged's own callers rely on elsewhere.
    // ###########################################################################################
    public void SelectSchematicByName(string schematicName)
    {
        if (string.IsNullOrWhiteSpace(schematicName))
            return;

        // A thumbnail drag-reorder is in progress (the flag is set on drag start and cleared when
        // the drag ends), during which OnSchematicsThumbnailSelectionChanged early-returns.
        // Selecting here would move every gallery's highlight while the image, overlays and
        // currentFullResBitmap stayed on the OLD schematic. Dropping the request is right: the user
        // is reordering thumbnails, not asking for a different one, and the caller is a cross-tab
        // convenience rather than a command.
        if (this.SuppressThumbnailSelectionChanged)
            return;

        if (this.SelectedThumbnail is SchematicThumbnail current &&
            string.Equals(current.Name, schematicName, StringComparison.OrdinalIgnoreCase))
            return;

        var match = this.currentThumbnails.FirstOrDefault(t =>
            string.Equals(t.Name, schematicName, StringComparison.OrdinalIgnoreCase));

        if (match != null)
        {
            this.SelectedThumbnail = match;
        }
    }

    // ###########################################################################################
    // Saves the current thumbnail order for the active board, excluding the placeholder item.
    //
    // internal rather than private so SchematicsThumbnailGallery (same assembly) persists every
    // gallery's reordering through this one UserSettings.SetSchematicsOrder path - the order is the
    // same whether thumbnails are in the tab, detached or beside the fullscreen view.
    // ###########################################################################################
    // The board key normally comes from the main window's hardware/board selection, which no test
    // constructs. This is the seam that lets a headless test prove the ORDER this method persists -
    // placeholders and blank names filtered out - rather than only that the call does not throw.
    // The same idea, and the same wording, as TabWorkbooks.BoardKeyOverrideForTests. Null in the
    // running app, always.
    internal string? BoardKeyOverrideForTests { get; set; }

    internal void SaveCurrentThumbnailOrder()
    {
        var boardKey = this.BoardKeyOverrideForTests ?? this.MainWindow?.GetCurrentBoardKey();
        if (string.IsNullOrWhiteSpace(boardKey))
            return;

        var orderedNames = this.currentThumbnails
            .Where(x => !x.IsDropPlaceholder && !string.IsNullOrWhiteSpace(x.Name))
            .Select(x => x.Name)
            .ToList();

        UserSettings.SetSchematicsOrder(boardKey, orderedNames);
    }

    // ###########################################################################################
    // Builds a stable signature for thumbnail highlight state so thumbnails only rebuild when the
    // actual rendered state changes. The blink phase is NOT part of it (code review, 2026-10-13):
    // it is the highlight layer's opacity (ApplyBlinkPhase), so a blink tick rebuilds nothing.
    // ###########################################################################################
    private string BuildThumbnailHighlightSignature(bool hasComponentSelection, bool hasKiCadSelection)
    {
        // Keyed by the actual set of highlighted board labels rather than by which schematics have
        // at least one highlight, since two different label sets can highlight the same schematics.
        string componentPart = hasComponentSelection
            ? this.thisHighlightedBoardLabelsSignature
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