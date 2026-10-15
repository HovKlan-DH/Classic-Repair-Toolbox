using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.Geometry;
using System;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// The auto-fit thumbnail gallery: every thumbnail shown at once, as large as the space allows,
// in as many columns as make them largest - never scrolled. It is the ONE thumbnail panel CRT
// has, shown in three places: the Schematics tab's own thumbnail column (owner request,
// 2026-10-15: "it should behave exactly as the detached window"), the detached
// SchematicsThumbnailsWindow, and the tab's stand-in while the schematic is fullscreen
// (SchematicsFullscreenPlaceholder). Each is bound to the SAME currentThumbnails collection
// (TabSchematics.Thumbnails.cs), so nothing about board switches, worklog pill overlays or
// selected-image loading needs to be duplicated.
//
// Selection belongs to the TAB, never to a gallery (code review, 2026-10-15). TabSchematics holds
// it (SelectedThumbnail, raising SelectedThumbnailChanged) and the label editor's lock on it
// (IsThumbnailNavigationLocked); all the real work of a new selection - full-res image load,
// overlays, SetLastSchematicForBoard - stays in its OnSchematicsThumbnailSelectionChanged. A
// gallery only SHOWS that selection, and asks the tab for another one in exactly two ways: a left
// click or tap on a tile, acted on at the RELEASE (OnGalleryPointerReleased), and the arrow keys
// in its list (ThumbnailGalleryPanel says how they move through the grid). What its own ListBox
// does by itself is never passed on - "nothing selected" when the highlighted tile leaves the
// collection, a Ctrl+click toggling it off. Each gallery once copied its ListBox's selection to
// the tab and back, and every such "nothing" blanked the schematic view.
//
// Every press on a tile is marked handled, so the ListBoxItem never acts on it: it selected on a
// mouse press, so every drag switched the schematic - and so did the right button that opens the
// panel's menu. One press is followed at a time, and only that pointer's left-button release ends
// it: a second finger is not a new press.
//
// A gallery listens to the tab only while it is in the visual tree (FollowOwner), and catches up
// whenever it comes back: the tab lives as long as the application, and a subscription on it keeps
// a gallery alive. The tab's own gallery leaves the tree on EVERY tab switch (a TabControl detaches
// its content) and comes back; a fullscreen stand-in made while another tab was shown is never in
// the tree at all, and so never listens.
//
// Drag-to-reorder: the dragged thumbnail is LIFTED OUT of the shared collection and a placeholder
// row takes its slot, which then moves to whichever grid cell is nearest the pointer. An earlier
// version moved the item around live instead, which left no visible gap at all - reported as a
// reorder that looked like nothing was happening. Dropping puts the thumbnail back where the
// placeholder ended up and calls TabSchematics.SaveCurrentThumbnailOrder(), so persistence
// (UserSettings.SetSchematicsOrder) is the same wherever the gallery is shown. Anything else that
// ends a drag - the pointer taken by something else, the tab cancelling it, the gallery leaving
// the tree - puts the tile back where it started and saves nothing. No drag ghost: the grid
// reflows from the collection on every layout pass anyway.
//
// EVERY path that ends a drag must go through EndTileDrag, which is what releases the tab's shared
// SuppressThumbnailSelectionChanged guard. That flag stops ALL schematic selection app-wide while
// set, and an earlier version cleared it only on the pointer-release path - so a drag ending any
// other way (capture lost, window closed) left selection permanently dead until restart. A gallery
// releases the guard only when it set it (thisHoldsSelectionGuard): up to three galleries are
// alive at once, and one with no drag of its own must not lift another's.
// ###########################################################################################
public partial class SchematicsThumbnailGallery : UserControl
{
    private TabSchematics? thisOwner;
    private ObservableCollection<SchematicThumbnail>? thisThumbnails;
    private bool thisIsShowingOwnerSelection;
    private bool thisIsOwnerSelectionShowPosted;
    private bool thisIsFollowingOwner;

    private IPointer? thisPressedPointer;
    private bool thisIsDraggingTile;
    private Point thisDragStartPoint;
    private SchematicThumbnail? thisDraggedThumbnail;
    private int thisDraggedThumbnailOriginalIndex = -1;
    private SchematicThumbnail? thisDropPlaceholder;
    private bool thisHoldsSelectionGuard;

    public SchematicsThumbnailGallery()
    {
        this.InitializeComponent();
    }

    // ###########################################################################################
    // Wires the gallery to its tab - the tab's thumbnails, its selection and lock, its order saving
    // and the drag guard - and shows the selection and lock as they are now. Listening for changes
    // starts once the gallery is in the visual tree (OnAttachedToVisualTree), or here if it already is.
    // ###########################################################################################
    public void Initialize(TabSchematics owner)
    {
        this.thisOwner = owner;
        this.thisThumbnails = owner.currentThumbnails;

        this.GalleryListBox.ItemsSource = owner.currentThumbnails;
        this.GalleryListBox.SelectionChanged += this.OnGallerySelectionChanged;

        this.ShowOwnerState();

        if (this.IsAttachedToVisualTree())
            this.FollowOwner();

        // On THIS control, not the tiles: the pressed tile is removed from the collection the
        // moment a drag begins, taking its container - and any handler or capture on it - with it.
        // handledEventsToo because the tile's own press handler marks the press handled.
        this.AddHandler(PointerMovedEvent, this.OnGalleryPointerMoved, handledEventsToo: true);
        this.AddHandler(PointerReleasedEvent, this.OnGalleryPointerReleased, handledEventsToo: true);
        this.AddHandler(PointerCaptureLostEvent, this.OnGalleryPointerCaptureLost, handledEventsToo: true);
    }

    // The tile this gallery highlights: the tab's selection, while the gallery is in the visual tree.
    public SchematicThumbnail? SelectedThumbnail => this.GalleryListBox.SelectedItem as SchematicThumbnail;

    // ###########################################################################################
    // The gallery's own list moved its highlight. Only the arrow keys do that on purpose - a click
    // selects on the release instead - so a real tile is asked of the tab, and whatever the list
    // shows afterwards that the tab does not hold is put back: "nothing" from a tile leaving the
    // collection or a Ctrl+Space toggle, the drop placeholder (it names no schematic), or a pick the
    // tab refused (locked, or mid-drag). Out of the visual tree the list is ignored altogether; the
    // gallery catches up when it is back.
    // ###########################################################################################
    private void OnGallerySelectionChanged(object? sender, SelectionChangedEventArgs e)
    {
        if (this.thisIsShowingOwnerSelection || !this.thisIsFollowingOwner || this.thisOwner == null)
            return;

        if (this.GalleryListBox.SelectedItem is SchematicThumbnail { IsDropPlaceholder: false } picked)
            this.thisOwner.SelectThumbnailFromGallery(picked);

        if (!ReferenceEquals(this.GalleryListBox.SelectedItem, this.thisOwner.SelectedThumbnail))
            this.ShowOwnerSelectionSoon();
    }

    // ###########################################################################################
    // The tiles changed. A tile that comes back - a dragged one dropped, in this gallery or another
    // - is not highlighted again by the list on its own, so the tab's selection is shown again.
    // ###########################################################################################
    private void OnThumbnailsChanged(object? sender, NotifyCollectionChangedEventArgs e)
    {
        if (!ReferenceEquals(this.GalleryListBox.SelectedItem, this.thisOwner?.SelectedThumbnail))
            this.ShowOwnerSelectionSoon();
    }

    // ###########################################################################################
    // Shows the tab's selection once the list has caught up: this is reached in the middle of a
    // collection change, before the list itself has seen it. Several in a row (a board's tiles
    // added one by one) are shown once.
    // ###########################################################################################
    private void ShowOwnerSelectionSoon()
    {
        if (this.thisIsOwnerSelectionShowPosted)
            return;

        this.thisIsOwnerSelectionShowPosted = true;

        Dispatcher.UIThread.Post(() =>
        {
            this.thisIsOwnerSelectionShowPosted = false;
            this.ShowOwnerSelection();
        });
    }

    // ###########################################################################################
    // Highlights the tab's selection - or nothing while it is not among the tiles, which is the
    // case while it is the tile a drag has lifted out.
    // ###########################################################################################
    private void ShowOwnerSelection()
    {
        if (this.thisOwner == null || this.thisThumbnails == null)
            return;

        var selected = this.thisOwner.SelectedThumbnail;

        this.thisIsShowingOwnerSelection = true;

        try
        {
            this.GalleryListBox.SelectedItem = selected != null && this.thisThumbnails.Contains(selected) ? selected : null;
        }
        finally
        {
            this.thisIsShowingOwnerSelection = false;
        }
    }

    private void ShowOwnerState()
    {
        if (this.thisOwner == null)
            return;

        this.ShowOwnerSelection();
        this.GalleryListBox.IsEnabled = !this.thisOwner.IsThumbnailNavigationLocked;
    }

    private void OnOwnerSelectedThumbnailChanged(object? sender, EventArgs e) => this.ShowOwnerSelection();

    // ###########################################################################################
    // The label editor's lock reaches every gallery through the tab. Locking also ends a drag in
    // progress here.
    // ###########################################################################################
    private void OnOwnerThumbnailNavigationLockChanged(object? sender, EventArgs e)
    {
        if (this.thisOwner?.IsThumbnailNavigationLocked == true)
            this.CancelTileDrag();

        this.ShowOwnerState();
    }

    private void OnOwnerThumbnailDragsCancelRequested(object? sender, EventArgs e) => this.CancelTileDrag();

    // ###########################################################################################
    // Listens to the tab and its tiles, and catches up on whatever changed while not listening -
    // the selection and the lock. A no-op while already listening, or once torn down.
    // ###########################################################################################
    private void FollowOwner()
    {
        if (this.thisIsFollowingOwner || this.thisOwner == null)
            return;

        this.thisIsFollowingOwner = true;
        this.thisOwner.SelectedThumbnailChanged += this.OnOwnerSelectedThumbnailChanged;
        this.thisOwner.ThumbnailNavigationLockChanged += this.OnOwnerThumbnailNavigationLockChanged;
        this.thisOwner.ThumbnailDragsCancelRequested += this.OnOwnerThumbnailDragsCancelRequested;
        this.thisOwner.currentThumbnails.CollectionChanged += this.OnThumbnailsChanged;

        this.ShowOwnerState();
    }

    private void StopFollowingOwner()
    {
        if (!this.thisIsFollowingOwner || this.thisOwner == null)
            return;

        this.thisIsFollowingOwner = false;
        this.thisOwner.SelectedThumbnailChanged -= this.OnOwnerSelectedThumbnailChanged;
        this.thisOwner.ThumbnailNavigationLockChanged -= this.OnOwnerThumbnailNavigationLockChanged;
        this.thisOwner.ThumbnailDragsCancelRequested -= this.OnOwnerThumbnailDragsCancelRequested;
        this.thisOwner.currentThumbnails.CollectionChanged -= this.OnThumbnailsChanged;
    }

    // ###########################################################################################
    // Starts tracking a tile for a possible 2D drag-reorder.
    //
    // Every press on a tile is handled, whatever the button, so the ListBoxItem neither selects on
    // it (a mouse) nor arms its own selection for the release (touch and pen). The panel's
    // right-click menu listens with handledEventsToo, so it still opens.
    //
    // The pointer is captured by the GALLERY, never by the tile that was pressed. Beginning a drag
    // removes that tile's item from the collection, so the ListBox destroys the container holding
    // the capture - and with it every further PointerMoved. That is why the placeholder appeared
    // but could never be moved: the drag was over the instant it began. For the same reason the
    // move/release handlers are subscribed on this control rather than on the item template.
    // ###########################################################################################
    private void OnGalleryTilePointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (sender is not Control control || control.DataContext is not SchematicThumbnail thumbnail)
            return;

        e.Handled = true;

        if (thumbnail.IsDropPlaceholder || this.thisPressedPointer != null)
            return;

        if (e.GetCurrentPoint(this).Properties.PointerUpdateKind != PointerUpdateKind.LeftButtonPressed)
            return;

        this.thisPressedPointer = e.Pointer;
        this.thisDragStartPoint = e.GetPosition(this);
        this.thisDraggedThumbnail = thumbnail;
        this.thisIsDraggingTile = false;

        e.Pointer.Capture(this);
    }

    // ###########################################################################################
    // Every release that reaches a tile is handled too, for the same reason: the ListBoxItem
    // selects on a touch or pen release even when the press was handled. The press this gallery
    // follows is captured and released on the gallery itself, so what arrives here is any other
    // pointer - a second finger, whose release chose its tile over the first finger's tap. The
    // gallery's own release handler still sees it (handledEventsToo) and leaves it alone.
    // ###########################################################################################
    private void OnGalleryTilePointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        e.Handled = true;
    }

    // ###########################################################################################
    // Once past a small movement threshold, starts a 2D reorder: the dragged thumbnail is lifted
    // out of the collection and a placeholder takes its slot, which then follows the pointer to
    // whichever cell is nearest. Every other tile reflows around it, since ThumbnailGalleryPanel
    // recomputes the whole grid from the collection on each layout pass.
    // ###########################################################################################
    private void OnGalleryPointerMoved(object? sender, PointerEventArgs e)
    {
        if (this.thisDraggedThumbnail == null || !ReferenceEquals(e.Pointer, this.thisPressedPointer))
            return;

        var point = e.GetPosition(this);

        if (!this.thisIsDraggingTile)
        {
            var diff = this.thisDragStartPoint - point;
            if (Math.Abs(diff.X) <= 6 && Math.Abs(diff.Y) <= 6)
                return;

            this.BeginTileDrag();
        }

        if (this.GalleryListBox.ItemsPanelRoot is ThumbnailGalleryPanel panel)
            this.MoveDropPlaceholderNearest(e.GetPosition(panel));
    }

    // ###########################################################################################
    // Moves the drop placeholder to the grid cell nearest a point in the panel - the one step of a
    // drag both the pointer handler and the test seam take.
    // ###########################################################################################
    private void MoveDropPlaceholderNearest(Point pointInPanel)
    {
        if (this.thisThumbnails == null || this.thisDropPlaceholder == null)
            return;

        if (this.GalleryListBox.ItemsPanelRoot is not ThumbnailGalleryPanel panel)
            return;

        var layout = panel.CurrentLayout;

        // layout.Positions is the authoritative cell count, not the collection's: a lift-and-insert
        // pair can leave the panel's last computed layout and the collection disagreeing by one
        // within a single pointer-move, and FindNearestCellIndex clamps its scan to the smaller of
        // the two - silently excluding the final cell from the search, so the bottom-right cell
        // could not be dropped onto.
        int nearestIndex = ThumbnailGalleryGeometry.FindNearestCellIndex(
            layout.Positions, layout.CellSize, pointInPanel, layout.Positions.Count);

        if (nearestIndex < 0)
            return;

        int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);
        if (placeholderIndex < 0 || placeholderIndex == nearestIndex)
            return;

        this.thisThumbnails.Move(placeholderIndex, nearestIndex);
    }

    // ###########################################################################################
    // Lifts the dragged thumbnail out of the collection and drops a placeholder into its slot, so
    // there is a gap to see, and it follows the pointer.
    //
    // The dragged tile is genuinely REMOVED rather than moved around live: leaving it in place
    // meant there was no gap to see, which is what made a detached-window reorder look like
    // nothing was happening. The tab's selection is not the tiles', so it stays as it is.
    // ###########################################################################################
    private void BeginTileDrag()
    {
        if (this.thisThumbnails == null || this.thisDraggedThumbnail == null)
            return;

        int index = this.thisThumbnails.IndexOf(this.thisDraggedThumbnail);
        if (index < 0)
            return;

        this.thisIsDraggingTile = true;
        this.thisDraggedThumbnailOriginalIndex = index;

        // The tab's shared guard, so a cross-tab SelectSchematicByName cannot land mid-reorder.
        // ALWAYS released again in EndTileDrag.
        if (this.thisOwner != null)
        {
            this.thisOwner.SuppressThumbnailSelectionChanged = true;
            this.thisHoldsSelectionGuard = true;
        }

        this.thisDropPlaceholder = new SchematicThumbnail { IsDropPlaceholder = true };

        this.thisThumbnails.RemoveAt(index);
        this.thisThumbnails.Insert(index, this.thisDropPlaceholder);
    }

    // ###########################################################################################
    // Puts the dragged thumbnail back and clears every piece of drag and press state. Safe to call
    // more than once, and from any of the paths that can end a drag - pointer release, capture
    // loss, the window being detached mid-drag, the tab cancelling it, or Teardown.
    //
    // A drop (commitOrder) lands where the placeholder ended up and saves that order. Anything else
    // puts the tile back where it STARTED: nothing saves the order then, and leaving the tile where
    // the gap had got to showed an order the next board load quietly undid (code review, 2026-10-15).
    // An order is saved only when the tile was really put back: a drop after the thumbnails were
    // emptied for a board load once saved that empty list as the board's order.
    //
    // Releasing the shared suppress flag here rather than only on the release path is the whole
    // point: a drag that ended any other way used to leave it set, and with it set NOTHING in the
    // app could change the selected schematic again until restart.
    // ###########################################################################################
    private void EndTileDrag(bool commitOrder)
    {
        bool isPutBack = false;

        if (this.thisThumbnails != null && this.thisDropPlaceholder != null)
        {
            int placeholderIndex = this.thisThumbnails.IndexOf(this.thisDropPlaceholder);

            if (placeholderIndex >= 0)
            {
                this.thisThumbnails.RemoveAt(placeholderIndex);

                if (this.thisDraggedThumbnail != null)
                {
                    int landingIndex = commitOrder
                        ? placeholderIndex
                        : Math.Clamp(this.thisDraggedThumbnailOriginalIndex, 0, this.thisThumbnails.Count);

                    this.thisThumbnails.Insert(landingIndex, this.thisDraggedThumbnail);
                    isPutBack = true;
                }
            }
        }

        this.thisDropPlaceholder = null;

        if (this.thisHoldsSelectionGuard && this.thisOwner != null)
            this.thisOwner.SuppressThumbnailSelectionChanged = false;

        this.thisHoldsSelectionGuard = false;

        if (commitOrder && isPutBack)
            this.thisOwner?.SaveCurrentThumbnailOrder();

        this.thisPressedPointer = null;
        this.thisIsDraggingTile = false;
        this.thisDraggedThumbnail = null;
        this.thisDraggedThumbnailOriginalIndex = -1;
    }

    // ###########################################################################################
    // Ends the press: a drag is dropped and its order saved through the same
    // TabSchematics.SaveCurrentThumbnailOrder every gallery uses; no real drag is an ordinary click
    // or tap, which asks the tab for that schematic. Only the left-button release of the pointer
    // that pressed counts - a second finger let go is not this press ending.
    //
    // The capture is let go only AFTER the press has been dealt with. Letting go raises
    // PointerCaptureLost at once, which would end the drag as a cancel - letting go first ended the
    // click before it was read, so a touch or pen tap selected nothing (code review, 2026-10-15).
    // ###########################################################################################
    private void OnGalleryPointerReleased(object? sender, PointerReleasedEventArgs e)
    {
        // Subscribed on the whole gallery, so this also sees releases that never started on a tile.
        var pressed = this.thisDraggedThumbnail;
        if (pressed == null || !ReferenceEquals(e.Pointer, this.thisPressedPointer))
            return;

        if (e.InitialPressMouseButton != MouseButton.Left)
            return;

        bool wasDragging = this.thisIsDraggingTile;

        this.EndTileDrag(commitOrder: wasDragging);
        e.Pointer.Capture(null);

        if (!wasDragging)
            this.thisOwner?.SelectThumbnailFromGallery(pressed);
    }

    // ###########################################################################################
    // The pointer taken away mid-press - another control capturing it, the window losing it to
    // Alt+Tab or the OS. That is not a drop: the tile goes back where it started and no order is
    // saved, as for every other end that is not a release (code review, 2026-10-15). Without this
    // the shared suppress flag stayed set and the whole app lost the ability to change schematic.
    // ###########################################################################################
    private void OnGalleryPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
    {
        if (this.thisDraggedThumbnail == null || !ReferenceEquals(e.Pointer, this.thisPressedPointer))
            return;

        this.EndTileDrag(commitOrder: false);
    }

    // ###########################################################################################
    // Back in the visual tree - the tab shown again after another tab, or moved back from the
    // fullscreen window: listen to the tab again, and show what was selected meanwhile.
    // ###########################################################################################
    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);

        this.FollowOwner();
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnDetachedFromVisualTree(e);

        // Never leave the tab's shared selection guard set behind a closing window or a tab switch.
        this.EndTileDrag(commitOrder: false);

        this.StopFollowingOwner();
    }

    // ###########################################################################################
    // Ends a drag or a held press here without saving the order - for the tab's own reasons to
    // drop one (a board load, the column being collapsed, the label editor's lock). A no-op when
    // this gallery is not pressed or dragging, so it never lifts another gallery's guard.
    // ###########################################################################################
    internal void CancelTileDrag()
    {
        if (this.thisDraggedThumbnail == null && this.thisDropPlaceholder == null)
            return;

        this.EndTileDrag(commitOrder: false);
    }

    // ###########################################################################################
    // Lets go of the tab for good - for a gallery that will never come back, the detached window's
    // once closed. A drag still running is ended first: only a gallery that still has its tab can
    // release the guard, so letting go mid-drag would have left schematic selection dead everywhere
    // (code review, 2026-10-15). Leaving the visual tree already does both; this makes sure nothing
    // can make the gallery listen again. Idempotent.
    // ###########################################################################################
    internal void Teardown()
    {
        this.CancelTileDrag();
        this.StopFollowingOwner();

        this.thisOwner = null;
    }

    // ###########################################################################################
    // Test seams. These drive the same steps the pointer handlers call, rather than a parallel
    // implementation of them; the real pointer path itself is covered by the tests that press and
    // drag the tab's column with headless input (SchematicsThumbnailGalleryTests).
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

    // Runs the REAL pointer-moved reorder step against a supplied pointer position.
    internal void DragPlaceholderToPointForTests(Point pointInPanel) => this.MoveDropPlaceholderNearest(pointInPanel);
}
