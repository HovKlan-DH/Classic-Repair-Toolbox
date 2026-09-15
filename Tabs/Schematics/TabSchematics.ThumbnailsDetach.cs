using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;

namespace CRT;

// ###########################################################################################
// Detach-to-window mode for the thumbnail gallery: collapsing the inline thumbnail column and
// restoring it, mirroring EnterFullscreenMode/ExitFullscreenMode (TabSchematics.axaml.cs) almost
// exactly - the same three-column collapse/restore mechanic, just triggered by the "Detach
// thumbnails to its own window" Configuration setting instead of F11.
//
// The two modes keep SEPARATE restore fields rather than sharing one set, because they COMPOSE:
// either can be entered while the other is already active. Whichever mode is entered second takes
// the real pre-collapse widths off the first (which is holding them, since the live columns are
// already collapsed) and hands back the collapsed ones in exchange, so exactly one set of fields
// holds the user's actual splitter position at any moment and whichever mode is exited LAST is the
// one that restores it. One shared set would be overwritten with collapsed values by the second
// entry, losing that position for good.
//
// The detached window itself and its own 2D-aware gallery/drag-reorder live in
// SchematicsThumbnailGallery, SchematicsThumbnailsWindow and ThumbnailGalleryPanel. This file
// owns only the inline-column side of the toggle.
//
// This file also owns the small right-click menu over the inline thumbnail panel itself
// (SchematicsThumbnailsContextMenuBorder) - a quicker way to detach than going to Configuration,
// which stays the primary/discoverable way to do it and to turn it back off from the tab side.
// It is a separate control from SchematicsLabelEditorMenuBorder: that one is contributor-mode/
// label-editor/calibration specific and positioned over the schematic image, not the thumbnail
// column, and this action has nothing to do with any of those modes.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private bool thisIsThumbnailsDetached;

    // Exposed so Main can skip re-applying the per-board schematics/thumbnail splitter ratio
    // (OnBoardSelectionChanged) while thumbnails are detached - that code runs on every board
    // change and would otherwise silently re-expand column 2 straight back out from under this
    // mode, leaving an empty widened column since the inline list stays hidden.
    internal bool IsThumbnailsDetached => this.thisIsThumbnailsDetached;

    private GridLength thisDetachRestoreLeftColumnWidth = new(1, GridUnitType.Star);

    private GridLength thisDetachRestoreSplitterColumnWidth = new(4, GridUnitType.Pixel);

    private GridLength thisDetachRestoreRightColumnWidth = new(1, GridUnitType.Star);

    private double thisDetachRestoreRightColumnMinWidth = 100.0;

    // ###########################################################################################
    // Collapses the inline thumbnail column so the detached window's gallery becomes the only
    // place the thumbnails are shown. A no-op while already detached, and while F11 fullscreen is
    // active - fullscreen already hides thumbnails by moving the whole control into its own
    // window, so the two modes must not both try to own the same column.
    // ###########################################################################################
    public void EnterThumbnailsDetachedMode()
    {
        if (this.thisIsThumbnailsDetached)
            return;

        this.thisIsThumbnailsDetached = true;

        // Fullscreen has already collapsed this column, so the LIVE column widths are its collapsed
        // ones and the user's real pre-fullscreen widths are sitting in fullscreen's own restore
        // fields. Take them from THERE, and hand fullscreen the collapsed values in exchange so
        // leaving fullscreen keeps the strip collapsed for this mode rather than re-expanding it.
        //
        // Reading the live columns here instead - or, as an earlier version did, returning before
        // capturing anything at all - left this mode's restore fields at their bare field
        // initializers (1* and MinWidth 100), so later unticking the setting threw away whatever
        // width the user had dragged the splitter to and snapped the column to half the tab.
        if (this.thisIsFullscreenMode)
        {
            this.thisDetachRestoreLeftColumnWidth = this.thisRestoreLeftColumnWidth;
            this.thisDetachRestoreSplitterColumnWidth = this.thisRestoreSplitterColumnWidth;
            this.thisDetachRestoreRightColumnWidth = this.thisRestoreRightColumnWidth;
            this.thisDetachRestoreRightColumnMinWidth = this.thisRestoreRightColumnMinWidth;

            this.thisRestoreLeftColumnWidth = new GridLength(1, GridUnitType.Star);
            this.thisRestoreSplitterColumnWidth = new GridLength(0, GridUnitType.Pixel);
            this.thisRestoreRightColumnWidth = new GridLength(0, GridUnitType.Pixel);
            this.thisRestoreRightColumnMinWidth = 0;

            return;
        }

        this.thisDetachRestoreLeftColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[0].Width;
        this.thisDetachRestoreSplitterColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[1].Width;
        this.thisDetachRestoreRightColumnWidth = this.SchematicsInnerGrid.ColumnDefinitions[2].Width;
        this.thisDetachRestoreRightColumnMinWidth = this.SchematicsInnerGrid.ColumnDefinitions[2].MinWidth;

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
    // Restores the inline thumbnail column to its pre-detach width. A no-op if not detached.
    // ###########################################################################################
    public void ExitThumbnailsDetachedMode()
    {
        if (!this.thisIsThumbnailsDetached)
            return;

        this.thisIsThumbnailsDetached = false;

        // Mirror of EnterThumbnailsDetachedMode's fullscreen branch: fullscreen owns the live
        // columns, so the only thing to do is hand the pre-detach widths BACK to it. Without this
        // the values it would restore on the way out are the collapsed ones it was given when this
        // mode started, and leaving fullscreen would show an empty full-width column.
        // ExitFullscreenMode then sees the flag false and brings the strip back itself.
        if (this.thisIsFullscreenMode)
        {
            this.thisRestoreLeftColumnWidth = this.thisDetachRestoreLeftColumnWidth;
            this.thisRestoreSplitterColumnWidth = this.thisDetachRestoreSplitterColumnWidth;
            this.thisRestoreRightColumnWidth = this.thisDetachRestoreRightColumnWidth;
            this.thisRestoreRightColumnMinWidth = this.thisDetachRestoreRightColumnMinWidth;

            return;
        }

        this.SchematicsInnerGrid.ColumnDefinitions[0].Width = this.thisDetachRestoreLeftColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[1].Width = this.thisDetachRestoreSplitterColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[2].Width = this.thisDetachRestoreRightColumnWidth;
        this.SchematicsInnerGrid.ColumnDefinitions[2].MinWidth = this.thisDetachRestoreRightColumnMinWidth;

        this.SchematicsSplitter.IsVisible = true;
        this.SchematicsThumbnailList.IsVisible = true;

        this.RefreshAfterHostChanged();
    }

    // ###########################################################################################
    // Wires the thumbnail panel's right-click menu. Called once from Initialize.
    //
    // Attached to SchematicsThumbnailList rather than to individual tiles: a tile's own
    // PointerPressed only reacts to the left button (drag-reorder/select), so a right-click over
    // a tile still reaches the list underneath it. That covers the whole panel - "somewhere in the
    // thumbnails panel", not just the gaps between images.
    // ###########################################################################################
    private void InitializeThumbnailsContextMenu()
    {
        this.SchematicsThumbnailList.AddHandler(
            InputElement.PointerPressedEvent,
            this.OnThumbnailsPanelPointerPressedForContextMenu,
            handledEventsToo: true);

        this.AddHandler(
            InputElement.PointerPressedEvent,
            this.OnSchematicsPointerPressedDismissThumbnailsContextMenu,
            RoutingStrategies.Tunnel);

        this.AddHandler(
            InputElement.KeyDownEvent,
            this.OnSchematicsKeyDownDismissThumbnailsContextMenu,
            RoutingStrategies.Tunnel);
    }

    // ###########################################################################################
    // Escape dismisses the menu, ahead of OnSchematicsKeyDown's own mode-specific Escape handling
    // (tunnel runs first) - kept as a small handler of its own rather than threaded through that
    // large mode dispatch, since this menu is not a mode.
    // ###########################################################################################
    private void OnSchematicsKeyDownDismissThumbnailsContextMenu(object? sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape && this.SchematicsThumbnailsContextMenuBorder.IsVisible)
        {
            this.SchematicsThumbnailsContextMenuBorder.IsVisible = false;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Opens the menu at the pointer on a stationary right-click over the thumbnail panel. Never
    // while thumbnails are already detached (nothing to offer there - the panel is empty/hidden)
    // or mid worklog-area-drag, matching CanShowSchematicsActionsMenu's own worklog exclusion.
    // ###########################################################################################
    private void OnThumbnailsPanelPointerPressedForContextMenu(object? sender, PointerPressedEventArgs e)
    {
        if (!e.GetCurrentPoint(this.SchematicsThumbnailList).Properties.IsRightButtonPressed)
            return;

        if (this.thisIsThumbnailsDetached || this.thisIsWorklogEntryMode)
            return;

        // Positioned in the WHOLE GRID's space, not the thumbnail column's: the border spans all
        // three columns so it can grow over the schematic image, which is the only way it fits on
        // one line when the splitter leaves the column ~170px wide.
        Point pointInGrid = e.GetPosition(this.SchematicsInnerGrid);

        double gridWidth = this.SchematicsInnerGrid.Bounds.Width;
        double gridHeight = this.SchematicsInnerGrid.Bounds.Height;

        // Measured against the FULL grid, so the button's text is not re-wrapped into the narrow
        // column - measuring against the column is what produced one character per line.
        this.SchematicsThumbnailsContextMenuBorder.Measure(new Size(gridWidth, gridHeight));
        Size menuSize = this.SchematicsThumbnailsContextMenuBorder.DesiredSize;

        // DesiredSize can still be short of what the border actually arranges to (its Width sets
        // the CONTENT box, with padding and border added on top), and being wrong here puts the
        // menu's right edge off the tab. Take whichever is larger, and fall back to the declared
        // width plus chrome on the very first open, before it has ever been arranged.
        double chrome =
            this.SchematicsThumbnailsContextMenuBorder.Padding.Left +
            this.SchematicsThumbnailsContextMenuBorder.Padding.Right +
            this.SchematicsThumbnailsContextMenuBorder.BorderThickness.Left +
            this.SchematicsThumbnailsContextMenuBorder.BorderThickness.Right;

        double declaredWidth = double.IsNaN(this.SchematicsThumbnailsContextMenuBorder.Width)
            ? 0.0
            : this.SchematicsThumbnailsContextMenuBorder.Width + chrome;

        double menuWidth = Math.Max(
            Math.Max(menuSize.Width, declaredWidth),
            this.SchematicsThumbnailsContextMenuBorder.Bounds.Width);

        double menuHeight = Math.Max(menuSize.Height, this.SchematicsThumbnailsContextMenuBorder.Bounds.Height);

        // Opens to the LEFT of the cursor, as asked: the right-click happens in the thumbnail
        // column at the right-hand edge of the tab, so growing rightward always runs out of room.
        // Clamped to 0 so a click near the left edge still shows the whole menu.
        double x = Math.Clamp(
            pointInGrid.X - menuWidth,
            0.0,
            Math.Max(0.0, gridWidth - menuWidth));

        double y = Math.Clamp(
            pointInGrid.Y,
            0.0,
            Math.Max(0.0, gridHeight - menuHeight));

        this.SchematicsThumbnailsContextMenuBorder.Margin = new Thickness(x, y, 0, 0);
        this.SchematicsThumbnailsContextMenuBorder.IsVisible = true;

        e.Handled = true;
    }

    // ###########################################################################################
    // Light-dismiss: any OTHER press anywhere in the tab closes the menu, tunnelled so it runs
    // before a click reaches whatever is underneath (a thumbnail, the schematic image). A press on
    // the menu itself is left alone so its own button click still fires normally.
    // ###########################################################################################
    private void OnSchematicsPointerPressedDismissThumbnailsContextMenu(object? sender, PointerPressedEventArgs e)
    {
        if (!this.SchematicsThumbnailsContextMenuBorder.IsVisible)
            return;

        bool isInsideMenu = e.Source is Control source && this.SchematicsThumbnailsContextMenuBorder.IsVisualAncestorOf(source);

        if (isInsideMenu)
            return;

        this.SchematicsThumbnailsContextMenuBorder.IsVisible = false;
    }

    // ###########################################################################################
    // "Detach thumbnails into their own window" from the panel's own right-click menu - a quicker route
    // to the SAME Main.SetThumbnailsDetached entry point the Configuration checkbox and the
    // detached window's own close button use, rather than a second copy of the persist/apply/resync
    // sequence. Reimplementing those steps here is how the two routes came to differ by one.
    // ###########################################################################################
    private void OnDetachThumbnailsContextMenuButtonClick(object? sender, RoutedEventArgs e)
    {
        this.SchematicsThumbnailsContextMenuBorder.IsVisible = false;

        if (this.MainWindow != null)
        {
            this.MainWindow.SetThumbnailsDetached(true);
        }
        else
        {
            // No main window to apply the layout through - the tab is standing alone (only ever the
            // case in a headless test). The setting itself is still the user's choice and is
            // written either way, exactly as the Configuration checkbox does in the same situation.
            UserSettings.DetachSchematicsThumbnails = true;
        }
    }
}
