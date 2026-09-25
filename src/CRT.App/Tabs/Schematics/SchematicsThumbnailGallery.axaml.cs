using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Handlers.Geometry;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// The auto-fit thumbnail gallery shown inside SchematicsThumbnailsWindow when thumbnails are
// detached. Bound to the SAME currentThumbnails collection the inline SchematicsThumbnailList
// uses (TabSchematics.Thumbnails.cs), so nothing about board switches, worklog pill overlays or
// selected-image loading needs to be duplicated - only reflowed into a 2D grid instead of a
// single column.
//
// Selection sync mirrors SchematicsFullscreenPlaceholder.axaml.cs exactly: a click here forwards
// to the hosted (now-hidden) SchematicsThumbnailList, which fires the existing, unmodified
// OnSchematicsThumbnailSelectionChanged - all the real work (full-res image load, overlays,
// SetLastSchematicForBoard) stays there. A change on the hosted list from anywhere else (e.g.
// SelectSchematicByName, called from Workbooks) is mirrored back here the same way.
//
// Drag-to-reorder follows the inline list's model: the dragged thumbnail is LIFTED OUT of the
// shared collection and a placeholder row takes its slot, which then moves to whichever grid cell
// is nearest the pointer. An earlier version moved the item around live instead, which left no
// visible gap at all - reported as a reorder that looked like nothing was happening. Dropping puts
// the thumbnail back where the placeholder ended up and calls TabSchematics.SaveCurrentThumbnailOrder(),
// so persistence (UserSettings.SetSchematicsOrder) is identical whether thumbnails are inline or
// detached. No separate drag ghost: ThumbnailDragGhost is inline-list-specific, and the grid
// reflows from the collection on every layout pass anyway.
//
// EVERY path that ends a drag must go through EndTileDrag, which is what releases the tab's shared
// SuppressThumbnailSelectionChanged guard. That flag stops ALL schematic selection app-wide while
// set, and an earlier version cleared it only on the pointer-release path - so a drag ending any
// other way (capture lost, window closed) left selection permanently dead until restart.
// ###########################################################################################
public partial class SchematicsThumbnailGallery : UserControl
{
    private ObservableCollection<SchematicThumbnail>? thisThumbnails;
    private ListBox? thisHostedThumbnailList;
    private TabSchematics? thisOwner;
    private bool thisSuppressSelectionSync;

    private bool thisIsDraggingTile;
    private Point thisDragStartPoint;
    private SchematicThumbnail? thisDraggedThumbnail;
    private SchematicThumbnail? thisDropPlaceholder;
    private bool thisDraggedThumbnailWasSelected;

    public SchematicsThumbnailGallery()
    {
        this.InitializeComponent();
    }

    // ###########################################################################################
    // Wires the shared thumbnails collection, the hosted (hidden) ListBox for selection
    // forwarding, and the owning tab for persisting reorder and suppressing cross-tab selection
    // during a drag - the same guard SelectSchematicByName already checks for the inline list's
    // own drag, shared rather than duplicated.
    // ###########################################################################################
    public void Initialize(ObservableCollection<SchematicThumbnail> thumbnails, ListBox hostedThumbnailList, TabSchematics owner)
    {
        this.thisThumbnails = thumbnails;
        this.thisHostedThumbnailList = hostedThumbnailList;
        this.thisOwner = owner;

        this.GalleryListBox.ItemsSource = thumbnails;
        this.SelectedThumbnail = hostedThumbnailList.SelectedItem as SchematicThumbnail;

        this.GalleryListBox.SelectionChanged += this.OnGallerySelectionChanged;
        hostedThumbnailList.SelectionChanged += this.OnHostedThumbnailSelectionChanged;

        // On THIS control, not the tiles: the pressed tile is removed from the collection the
        // moment a drag begins, taking its container - and any handler or capture on it - with it.
        // handledEventsToo because the ListBox marks pointer events handled for its own selection.
        this.AddHandler(PointerMovedEvent, this.OnGalleryPointerMoved, handledEventsToo: true);
        this.AddHandler(PointerReleasedEvent, this.OnGalleryPointerReleased, handledEventsToo: true);
        this.AddHandler(PointerCaptureLostEvent, this.OnGalleryPointerCaptureLost, handledEventsToo: true);
    }

    public SchematicThumbnail? SelectedThumbnail
    {
        get => this.GalleryListBox.SelectedItem as SchematicThumbnail;
        set => this.GalleryListBox.SelectedItem = value;
    }

    // ###########################################################################################
    // Forwards a gallery selection to the hosted list, which does the real work via its own
    // unmodified OnSchematicsThumbnailSelectionChanged handler.
    // ###########################################################################################
    private void OnGallerySelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (this.thisSuppressSelectionSync || this.thisHostedThumbnailList == null)
            return;

        // The drop placeholder is a real row in the shared collection while a drag is running, so
        // it can be landed on by a click or the arrow keys. It names no schematic, and forwarding
        // it would blank the schematic view.
        if (this.GalleryListBox.SelectedItem is SchematicThumbnail { IsDropPlaceholder: true })
            return;

        this.thisSuppressSelectionSync = true;
        this.thisHostedThumbnailList.SelectedItem = this.GalleryListBox.SelectedItem;
        this.thisSuppressSelectionSync = false;
    }

    // ###########################################################################################
    // Mirrors a selection change made elsewhere (e.g. SelectSchematicByName from Workbooks) back
    // into the gallery's own highlighted tile.
    // ###########################################################################################
    private void OnHostedThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (this.thisSuppressSelectionSync || this.thisHostedThumbnailList == null)
            return;

        this.thisSuppressSelectionSync = true;
        this.GalleryListBox.SelectedItem = this.thisHostedThumbnailList.SelectedItem;
        this.thisSuppressSelectionSync = false;
    }

    // ###########################################################################################
    // Starts tracking a tile for a possible 2D drag-reorder.
    //
    // The pointer is captured by the GALLERY, never by the tile that was pressed. Beginning a drag
    // removes that tile's item from the collection, so the ListBox destroys the container holding
    // the capture - and with it every further PointerMoved. That is why the placeholder appeared
    // but could never be moved: the drag was over the instant it began. For the same reason the
    // move/release handlers are subscribed on this control rather than on the item template.
    // ###########################################################################################
    private void OnGalleryTilePointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            return;

        if (sender is not Control control || control.DataContext is not SchematicThumbnail thumbnail)
            return;

        if (thumbnail.IsDropPlaceholder)
            return;

        this.thisDragStartPoint = e.GetPosition(this);
        this.thisDraggedThumbnail = thumbnail;
        this.thisIsDraggingTile = false;

        e.Pointer.Capture(this);
    }

    // ###########################################################################################
    // Once past a small movement threshold, starts a 2D reorder: the dragged thumbnail is lifted
    // out of the collection and a placeholder takes its slot, which then follows the pointer to
    // whichever cell is nearest. Every other tile reflows around it, since ThumbnailGalleryPanel
    // recomputes the whole grid from the collection on each layout pass.
    // ###########################################################################################
    private void OnGalleryPointerMoved(object? sender, PointerEventArgs e)
    {
        if (this.thisDraggedThumbnail == null || this.thisThumbnails == null)
            return;

        var point = e.GetPosition(this);

        if (!this.thisIsDraggingTile)
        {
            var diff = this.thisDragStartPoint - point;
            if (Math.Abs(diff.X) <= 6 && Math.Abs(diff.Y) <= 6)
                return;

            this.BeginTileDrag();
        }

        if (this.thisDropPlaceholder == null)
            return;

        if (this.GalleryListBox.ItemsPanelRoot is not ThumbnailGalleryPanel panel)
            return;

        var pointerInPanel = e.GetPosition(panel);
        var layout = panel.CurrentLayout;

        // layout.Positions is the authoritative cell count, not the collection's: a lift-and-insert
        // pair can leave the panel's last computed layout and the collection disagreeing by one
        // within a single pointer-move, and FindNearestCellIndex clamps its scan to the smaller of
        // the two - silently excluding the final cell from the search, so the bottom-right cell
        // could not be dropped onto.
        int nearestIndex = ThumbnailGalleryGeometry.FindNearestCellIndex(
            layout.Positions, layout.CellSize, pointerInPanel, layout.Positions.Count);

        if (nearestIndex < 0)
            return;

        int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);
        if (placeholderIndex < 0 || placeholderIndex == nearestIndex)
            return;

        this.thisThumbnails.Move(placeholderIndex, nearestIndex);
    }

    // ###########################################################################################
    // Lifts the dragged thumbnail out of the collection and drops a placeholder into its slot, so
    // the gap follows the pointer the way the inline thumbnail list's own reorder shows it.
    //
    // The dragged tile is genuinely REMOVED rather than moved around live: leaving it in place
    // meant there was no gap to see, which is what made a detached-window reorder look like
    // nothing was happening.
    // ###########################################################################################
    private void BeginTileDrag()
    {
        if (this.thisThumbnails == null || this.thisDraggedThumbnail == null)
            return;

        int index = this.thisThumbnails.IndexOf(this.thisDraggedThumbnail);
        if (index < 0)
            return;

        this.thisIsDraggingTile = true;
        this.thisDraggedThumbnailWasSelected =
            ReferenceEquals(this.GalleryListBox.SelectedItem, this.thisDraggedThumbnail);

        // The tab's shared guard, so a cross-tab SelectSchematicByName cannot land mid-reorder -
        // the same flag the inline list's own drag sets. ALWAYS released again in EndTileDrag.
        if (this.thisOwner != null)
            this.thisOwner.SuppressThumbnailSelectionChanged = true;

        this.thisDropPlaceholder = new SchematicThumbnail { IsDropPlaceholder = true };

        this.thisThumbnails.RemoveAt(index);
        this.thisThumbnails.Insert(index, this.thisDropPlaceholder);
    }

    // ###########################################################################################
    // Puts the dragged thumbnail back where the placeholder ended up and clears every piece of
    // drag state. Safe to call more than once, and from any of the paths that can end a drag -
    // pointer release, capture loss, or the window being detached mid-drag.
    //
    // Releasing the shared suppress flag here rather than only on the release path is the whole
    // point: a drag that ended any other way used to leave it set, and with it set NOTHING in the
    // app could change the selected schematic again until restart.
    // ###########################################################################################
    private void EndTileDrag(bool commitOrder)
    {
        if (this.thisThumbnails != null && this.thisDropPlaceholder != null)
        {
            int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);

            if (placeholderIndex >= 0)
            {
                this.thisThumbnails.RemoveAt(placeholderIndex);

                if (this.thisDraggedThumbnail != null)
                {
                    this.thisThumbnails.Insert(placeholderIndex, this.thisDraggedThumbnail);

                    // Reselect through the gallery so the highlight follows the tile to its new
                    // slot, and so the hosted list keeps pointing at the same schematic.
                    if (this.thisDraggedThumbnailWasSelected)
                    {
                        this.GalleryListBox.SelectedItem = this.thisDraggedThumbnail;
                    }
                }
            }
        }

        this.thisDropPlaceholder = null;

        if (this.thisOwner != null)
            this.thisOwner.SuppressThumbnailSelectionChanged = false;

        if (commitOrder && this.thisIsDraggingTile)
            this.thisOwner?.SaveCurrentThumbnailOrder();

        this.thisIsDraggingTile = false;
        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailWasSelected = false;
    }

    // ###########################################################################################
    // Finalizes the drag: persists the new order via the same path the inline list uses, or - if
    // no real drag happened - treats the release as a plain selection click.
    // ###########################################################################################
    private void OnGalleryPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        // Subscribed on the whole gallery, so this also sees releases that never started on a tile.
        if (this.thisDraggedThumbnail == null)
            return;

        e.Pointer.Capture(null);

        if (this.thisIsDraggingTile)
        {
            this.EndTileDrag(commitOrder: true);
            return;
        }

        // No real drag: an ordinary click, so just select the tile.
        var clicked = this.thisDraggedThumbnail;
        this.EndTileDrag(commitOrder: false);

        if (clicked != null)
        {
            this.SelectedThumbnail = clicked;
        }
    }

    // ###########################################################################################
    // A drag can also end without a PointerReleased arriving at all - another control taking the
    // capture, or the window closing mid-drag. Without this the shared suppress flag stayed set and
    // the whole app lost the ability to change schematic.
    // ###########################################################################################
    private void OnGalleryPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        if (this.thisDraggedThumbnail == null)
            return;

        this.EndTileDrag(commitOrder: this.thisIsDraggingTile);
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);

        // Never leave the tab's shared selection guard set behind a closing window.
        this.EndTileDrag(commitOrder: false);

        this.Teardown();
    }

    // ###########################################################################################
    // Unsubscribes from the HOSTED list, which is the one thing here that outlives this control.
    // SchematicsThumbnailList lives for the application's lifetime inside TabSchematics, so a
    // subscription left on it keeps this whole gallery - its ListBox, its containers and every
    // decoded thumbnail bitmap they reference - reachable forever. Toggling the setting off and on
    // ten times left ten dead galleries behind, each still running OnHostedThumbnailSelectionChanged
    // on every schematic change and writing SelectedItem on detached controls.
    //
    // Idempotent, so the detach that happens on the way to a close can run it and a later explicit
    // call costs nothing.
    // ###########################################################################################
    internal void Teardown()
    {
        if (this.thisHostedThumbnailList != null)
        {
            this.thisHostedThumbnailList.SelectionChanged -= this.OnHostedThumbnailSelectionChanged;
            this.thisHostedThumbnailList = null;
        }

        this.thisOwner = null;
    }

    // ###########################################################################################
    // Test seams. The reorder is pointer-driven, and a headless test cannot produce a real pointer
    // capture against tiles that a never-shown window has not laid out - so the tests drive the
    // same three steps the pointer handlers call, rather than a parallel implementation of them.
    // ###########################################################################################
    internal void BeginTileDragForTests(SchematicThumbnail thumbnail)
    {
        this.thisDraggedThumbnail = thumbnail;
        this.BeginTileDrag();
    }

    internal void MoveDragPlaceholderForTests(int targetIndex)
    {
        if (this.thisThumbnails == null || this.thisDropPlaceholder == null)
            return;

        int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);
        if (placeholderIndex < 0 || placeholderIndex == targetIndex)
            return;

        this.thisThumbnails.Move(placeholderIndex, targetIndex);
    }

    internal void EndTileDragForTests(bool commitOrder) => this.EndTileDrag(commitOrder);

    internal bool IsDraggingForTests => this.thisIsDraggingTile;

    // Runs the REAL pointer-moved reorder step against a supplied pointer position, so a test
    // covers the handler and its guards rather than a parallel copy of the move logic.
    internal void DragPlaceholderToPointForTests(Point pointInPanel)
    {
        if (this.thisThumbnails == null || this.thisDropPlaceholder == null)
            return;

        if (this.GalleryListBox.ItemsPanelRoot is not ThumbnailGalleryPanel panel)
            return;

        var layout = panel.CurrentLayout;

        int nearestIndex = ThumbnailGalleryGeometry.FindNearestCellIndex(
            layout.Positions, layout.CellSize, pointInPanel, layout.Positions.Count);

        if (nearestIndex < 0)
            return;

        int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);
        if (placeholderIndex < 0 || placeholderIndex == nearestIndex)
            return;

        this.thisThumbnails.Move(placeholderIndex, nearestIndex);
    }
}
