using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// Worklog "Add worklog" area-marking mode: drag-draw a rectangle on the board to mark where a
// fault is, then open the FULL worklog editor on it (WorklogEntryEditorWindow, via
// InitializeForNewEntry), pre-populated with every component whose highlight rectangle
// intersects the drawn area. Cancelling there writes nothing and the drawn area goes with the
// mode; its Save is what creates the entry, through WorklogManager.AddEntryRecord.
//
// There used to be a small "New fault" quick card between the two, asking for a title,
// description, category, state and the component checklist - all of them fields the full editor
// also has, so reaching anything else (links, work done, comments, photos, files) meant saving
// and immediately reopening the same entry. It was removed rather than kept as a shortcut:
// OpenNewWorklogEntryEditor is now the whole of that step.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private bool thisIsWorklogEntryMode;

    private bool thisIsDrawingWorklogEntryRectangle;

    private Point thisWorklogEntryDrawStartPixelPoint;

    private Rect? thisWorklogEntryDraftRectangle;

    private Rect? thisWorklogEntryFinalRectangle;

    private int thisWorklogEntryWorkbookId;

    // ###########################################################################################
    // The id AddEntry will hand the entry currently being drawn, resolved once when entry mode
    // starts and reused for the on-board badge and the card header.
    //
    // Not looked up per draw: WorklogManager.PeekNextEntryId reads and re-parses entries.json, and
    // RefreshWorklogEntryBadge runs from UpdateSchematicsTransform on every zoom/pan frame - so
    // asking it there did synchronous disk I/O at frame rate for a number that cannot change while
    // the mode is open. The sibling RescaleWorklogEntriesListBadges avoids the same trap by design.
    // ###########################################################################################
    private int thisWorklogEntryNextId = 1;

    private Border? thisWorklogEntryBadgeBorder;

    private ScaleTransform? thisWorklogEntryBadgeScaleTransform;

    // ###########################################################################################
    // The category a still-unsaved entry's marked area and "#N" badge are drawn in.
    //
    // Always the "Note" default now: the drawn area is only on screen between the mouse release and
    // the full editor opening over it, and the editor - where the category is actually chosen -
    // carries its own copy of that choice. It stays a field rather than becoming a literal because
    // the two call sites read as "the category of the entry being drawn", which is the thing that
    // may yet gain a way to be set before the editor opens.
    // ###########################################################################################
    private const string WorklogCategoryNote = "Note";

    private const string WorklogStateClosed = "Closed";

    private readonly string thisWorklogEntrySelectedCategory = WorklogCategoryNote;

    // ###########################################################################################
    // Whether the top-bar "Show worklogs" checkbox is currently checked. Mirrors the checkbox
    // rather than owning the preference: its starting value comes from
    // UserSettings.WorklogShowEntriesChecked (which defaults to true), and Main.RefreshWorklogBar
    // re-seeds it whenever the workbook on screen changes.
    // ###########################################################################################
    private bool thisIsShowingWorklogEntriesList;

    private int thisWorklogEntriesListWorkbookId;

    // ###########################################################################################
    // One "#N" badge + state pill placed on SchematicsWorklogEntriesBadgeCanvas for the "Show
    // worklogs" list view - one row per saved entry currently rendered for this schematic. Rebuilt
    // from scratch (control tree torn down and recreated) only when the underlying entry set can
    // actually change: the checkbox being toggled, a schematic switch, or a new/edited entry being
    // saved. A zoom/pan tick instead calls RescaleWorklogEntriesListBadges, which repositions and
    // rescales these same controls in place - see that method for why a full rebuild on every
    // pointer-move frame was wasteful (re-reading entries.json and recreating every control).
    // ###########################################################################################
    private readonly List<(Border Badge, Rect PixelRect, WorklogEntryRecord Entry, Color Color)> thisWorklogEntriesListBadges = new();

    // ###########################################################################################
    // The pills of entries whose "Show marked area" is unticked. They have no marked area to sit
    // on, so they are stacked in the top-right corner of the schematic panel instead of being
    // hidden - an entry with no area AND no pill would be invisible on the board and unreachable
    // without opening the worklog list.
    //
    // Kept separate from thisWorklogEntriesListBadges because they live on a different canvas and
    // obey different rules: these sit in the container's own coordinate space and do not move with
    // zoom or pan, so a rescale tick repositions them for the viewport rather than for the board.
    // ###########################################################################################
    private readonly List<(Border Badge, WorklogEntryRecord Entry)> thisWorklogParkedBadges = new();

    // Gap from the panel edges, and between stacked pills.
    private const double WorklogParkedBadgeMargin = 10.0;

    private const double WorklogParkedBadgeSpacing = 6.0;

    // ###########################################################################################
    // Whether "Add worklog" area-marking mode is currently waiting for an area to be drawn.
    //
    // Exists so Main's WINDOW-level Escape handler can tell whether to cancel the mode or leave the
    // key alone for whatever else wants it. The mode is entered by a button in the top bar above
    // the tabs, which keeps keyboard focus, so the tab's own OnSchematicsKeyDown never sees the
    // keypress - see Main.OnMainKeyDownCloseSinglePopup, which is where Escape actually lands.
    //
    // Public and read-only, the same shape and for the same reason as IsLabelEditorActive.
    // ###########################################################################################
    public bool IsWorklogEntryModeActive => this.thisIsWorklogEntryMode;

    // ###########################################################################################
    // Enters worklog entry-drawing mode for the given active workbook. Refuses to start while the
    // label editor or KiCad calibration mode is active, or while no schematic image is loaded, so
    // the caller can tell whether the mode actually started.
    // ###########################################################################################
    public bool BeginWorklogEntryMode(int workbookId)
    {
        if (this.currentFullResBitmap == null || this.thisIsLabelEditorMode || this.thisIsKiCadTraceCalibrationMode)
        {
            return false;
        }

        this.thisIsWorklogEntryMode = true;
        this.thisWorklogEntryWorkbookId = workbookId;
        this.thisWorklogEntryNextId = WorklogManager.PeekNextEntryId(workbookId);
        this.thisIsDrawingWorklogEntryRectangle = false;
        this.thisWorklogEntryDraftRectangle = null;
        this.thisWorklogEntryFinalRectangle = null;

        this.RefreshWorklogEntryOverlay();

        this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Cross);
        this.SchematicsContainer.Focus();
        this.Focus();

        // The crosshair alone does not say what to do with it.
        this.MainWindow?.ShowModeHint(Main.WorklogAreaModeHint);

        Logger.Info($"Worklog entry drawing mode enabled for workbook [#{workbookId}] on schematic [{this.GetCurrentSchematicName()}]");
        return true;
    }

    // ###########################################################################################
    // Exits worklog entry-drawing mode from any point in the flow (top-bar Cancel, Escape, or the
    // card's own close/Cancel/Save), clearing the drawn area and telling the main window to reset
    // the "Add worklog" / "Cancel" buttons. Safe to call when the mode is not active.
    // ###########################################################################################
    public void CancelWorklogEntryMode()
    {
        if (!this.thisIsWorklogEntryMode)
        {
            return;
        }

        this.thisIsWorklogEntryMode = false;
        this.thisIsDrawingWorklogEntryRectangle = false;
        this.thisWorklogEntryDraftRectangle = null;
        this.thisWorklogEntryFinalRectangle = null;

        this.RefreshWorklogEntryOverlay();

        this.SchematicsContainer.Cursor = Cursor.Default;
        this.SchematicsContainer.Focus();

        this.MainWindow?.HideModeHint();

        this.MainWindow?.ResetWorklogEntryModeButtons();

        Logger.Info("Worklog entry drawing mode canceled");
    }

    // ###########################################################################################
    // Starts drawing the entry area rectangle from the current bitmap pixel position.
    // ###########################################################################################
    private void StartDrawingWorklogEntryRectangle(Point startPixelPoint)
    {
        this.thisIsDrawingWorklogEntryRectangle = true;
        this.thisWorklogEntryDrawStartPixelPoint = startPixelPoint;

        this.thisWorklogEntryDraftRectangle = new Rect(startPixelPoint.X, startPixelPoint.Y, 0, 0);
        this.thisWorklogEntryFinalRectangle = null;

        this.RefreshWorklogEntryOverlay();
    }

    // ###########################################################################################
    // Updates the draft entry area rectangle while the mouse is being dragged.
    // ###########################################################################################
    private void UpdateDrawingWorklogEntryRectangle(Point currentPixelPoint)
    {
        if (!this.thisIsDrawingWorklogEntryRectangle)
        {
            return;
        }

        this.thisWorklogEntryDraftRectangle = RectGeometry.CreateNormalizedRect(
            this.thisWorklogEntryDrawStartPixelPoint,
            currentPixelPoint);

        this.RefreshWorklogEntryOverlay();
    }

    // ###########################################################################################
    // Finishes drawing the entry area rectangle and opens the FULL worklog editor on it. A
    // rectangle too small to be a deliberate drag (an accidental click) is discarded instead, the
    // same threshold the label editor uses for the same reason.
    //
    // This used to open a small "New fault" quick card anchored beside the drawn area, which asked
    // for a title, description, category, state and the component checklist - every one of them a
    // field the full editor also has - and anything further (links, work done, comments, photos,
    // files) then meant saving and reopening the same entry in the full editor. The card was
    // removed outright rather than kept as a shortcut, so there is one place a worklog entry is
    // written and one set of fields to keep in step.
    // ###########################################################################################
    private void CompleteDrawingWorklogEntryRectangle(Point releasePixelPoint)
    {
        if (!this.TryFinishWorklogEntryDrawing(releasePixelPoint, out var finalRect))
        {
            return;
        }

        this.OpenNewWorklogEntryEditor(finalRect);
    }

    // ###########################################################################################
    // Everything the drag's end does EXCEPT opening the editor: normalise the drawn rectangle,
    // leave drawing state, apply the too-small rule, and record the accepted area so the overlay
    // draws it.
    //
    // Split out from CompleteDrawingWorklogEntryRectangle so a headless test can exercise the
    // accept/reject decision - OpenNewWorklogEntryEditor calls ShowDialog, which needs a real
    // owner Window and cannot run in a test. The alternative was for the test seam to restate
    // this decision, which would let the two drift; here there is one copy and the seam calls it.
    //
    // Returns true when the drag was a deliberate area, with it in acceptedRectangle.
    // ###########################################################################################
    private bool TryFinishWorklogEntryDrawing(Point releasePixelPoint, out Rect acceptedRectangle)
    {
        acceptedRectangle = default;

        if (!this.thisIsDrawingWorklogEntryRectangle)
        {
            return false;
        }

        var finalRect = RectGeometry.CreateNormalizedRect(this.thisWorklogEntryDrawStartPixelPoint, releasePixelPoint);

        this.thisIsDrawingWorklogEntryRectangle = false;
        this.thisWorklogEntryDraftRectangle = null;

        if (LabelEditorGeometry.IsLabelEditorRectangleTooSmall(finalRect))
        {
            this.RefreshWorklogEntryOverlay();
            return false;
        }

        this.thisWorklogEntryFinalRectangle = finalRect;

        // Draws the marked area under the modal, so the user can see which area they are describing
        // while they fill the editor in.
        this.RefreshWorklogEntryOverlay();

        acceptedRectangle = finalRect;
        return true;
    }

    // ###########################################################################################
    // Opens the full worklog editor on the just-drawn area, leaving entry mode as it opens - see
    // the CancelWorklogEntryMode call below for why that happens before the await rather than after.
    //
    // The editor holds the new entry entirely in memory until its Save (see
    // WorklogEntryEditorWindow.InitializeForNewEntry), so cancelling here leaves the workbook
    // untouched and the drawn area simply disappears with the mode.
    //
    // async void because it is reached from a pointer-release handler; everything after the await
    // is therefore guarded, the same way OnWorklogEntryPillPointerPressed's continuation is - an
    // exception in an async void continuation is unobservable and takes the process down.
    // ###########################################################################################
    private async void OpenNewWorklogEntryEditor(Rect areaPixelRectangle)
    {
        int workbookId = this.thisWorklogEntryWorkbookId;
        string schematicName = this.GetCurrentSchematicName();

        if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
        {
            this.CancelWorklogEntryMode();
            return;
        }

        var editor = new WorklogEntryEditorWindow();
        editor.InitializeForNewEntry(workbookId, schematicName, areaPixelRectangle, this.currentFullResBitmap);

        // The same component checklist the quick card built, computed the same way and handed to the
        // editor the same way a saved entry's is - see BuildWorklogEntryComponentScope. Every row
        // starts TICKED for a new entry, which is what the card did: the user drew the area around
        // these components, so having them all in scope is the sensible starting point, and
        // unticking is quicker than ticking a list of eight.
        var componentsInScope = this.BuildNewWorklogEntryComponentScope(schematicName, areaPixelRectangle);
        if (componentsInScope != null)
        {
            editor.InitializeComponentScope(componentsInScope, tickAll: true);
        }

        // Entry mode ends BEFORE the modal is awaited, not after it returns.
        //
        // ShowDialog does not block the dispatcher, so leaving the mode live for the whole life of
        // the dialog left thisWorklogEntryWorkbookId and the drawn rectangle mutable underneath it:
        // any refresh reaching Main.ActivateWorkbookAcrossBoards while the editor was open calls
        // CancelWorklogEntryMode itself and clears them, and a second BeginWorklogEntryMode would
        // reassign them. Everything the editor needs is already captured above - the workbook id,
        // the schematic name, the area and the component scope - so there is nothing left for the
        // mode to hold, and ending it here means the state the editor was built from cannot be
        // changed out from under it.
        //
        // It also matches what the user sees: the drawn area's job is done once the editor is up.
        this.CancelWorklogEntryMode();

        try
        {
            await editor.ShowDialog(ownerWindow);
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to show the worklog entry editor: [{ex.Message}]");
            return;
        }

        try
        {
            if (!editor.WasSaved)
            {
                return;
            }

            this.MainWindow?.RefreshWorklogBar();
            this.RefreshWorklogEntriesListOverlay();
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to refresh worklog overlays after adding an entry: [{ex.Message}]");
        }
    }

    // ###########################################################################################
    // The component scope for an area that has no saved entry yet - the same intersection the
    // quick card computed, expressed against the shared WorklogEntryScope helper the saved-entry
    // path already uses so the two cannot drift.
    //
    // Returns null when there is no highlight-rect cache for this schematic at all, which is the
    // "scope unknown" case the editor treats differently from "scope is empty" - see
    // WorklogEntryScope.BuildComponentsInScope.
    // ###########################################################################################
    private List<(string BoardLabel, string DisplayName)>? BuildNewWorklogEntryComponentScope(
        string schematicName,
        Rect areaPixelRectangle) =>
        WorklogEntryScope.BuildComponentsInScope(
            this.MainWindow?.CurrentBoardData,
            this.highlightRectsBySchematicAndLabel,
            new WorklogEntryRecord
            {
                SchematicName = schematicName,
                AreaX = areaPixelRectangle.X,
                AreaY = areaPixelRectangle.Y,
                AreaWidth = areaPixelRectangle.Width,
                AreaHeight = areaPixelRectangle.Height
            });

    // ###########################################################################################
    // Pushes the current draft/final rectangle into the shared overlay control. The final
    // rectangle is rendered as both "selected" and "hovered" so ComponentLabelEditorOverlay draws
    // its corner+side marker handles around it, matching the mockup's marked-area look.
    // ###########################################################################################
    private void RefreshWorklogEntryOverlay()
    {
        Color categoryColor = this.GetSelectedWorklogEntryCategoryColor();

        if (!this.thisIsWorklogEntryMode || this.currentFullResBitmap == null)
        {
            this.SchematicsWorklogEntryOverlay.ApplyState(
                rectangles: Array.Empty<Rect>(),
                selectedIndex: -1,
                selectedIndices: Array.Empty<int>(),
                selectionBounds: null,
                hoveredIndex: -1,
                draftRectangle: null,
                snapGuides: Array.Empty<(Point Start, Point End)>(),
                bitmapPixelSize: this.currentFullResBitmap?.PixelSize ?? new PixelSize(0, 0),
                viewMatrix: this.schematicsMatrix,
                highlightColor: categoryColor,
                highlightOpacity: 0.12,
                isVisible: false);
            this.RefreshWorklogEntryBadge(null);
            return;
        }

        var rectangles = this.thisWorklogEntryFinalRectangle.HasValue
            ? new[] { this.thisWorklogEntryFinalRectangle.Value }
            : Array.Empty<Rect>();

        this.SchematicsWorklogEntryOverlay.ApplyState(
            rectangles: rectangles,
            selectedIndex: rectangles.Length > 0 ? 0 : -1,
            selectedIndices: Array.Empty<int>(),
            selectionBounds: null,
            hoveredIndex: rectangles.Length > 0 ? 0 : -1,
            draftRectangle: this.thisWorklogEntryDraftRectangle,
            snapGuides: Array.Empty<(Point Start, Point End)>(),
            bitmapPixelSize: this.currentFullResBitmap.PixelSize,
            viewMatrix: this.schematicsMatrix,
            highlightColor: categoryColor,
            highlightOpacity: 0.12,
            isVisible: true);

        this.RefreshWorklogEntryBadge(this.thisWorklogEntryFinalRectangle);
    }

    // ###########################################################################################
    // Resolves the currently selected category's color - see ResolveWorklogCategoryColor for the
    // shared lookup (falls back to IndianRed if the resource cannot be resolved for any reason).
    // ###########################################################################################
    private Color GetSelectedWorklogEntryCategoryColor() => this.ResolveWorklogCategoryColor(this.thisWorklogEntrySelectedCategory);

    // ###########################################################################################
    // Shows a small "#N" badge at the top-left of the freshly marked entry area, so the id the
    // entry is about to get is visible on the board while the editor is being filled in - the same
    // corner a SAVED entry's badge uses, so nothing moves when the entry is written.
    // Kept a constant screen size across zoom the same way the standard component labels do: a
    // ScaleTransform set to the inverse of the current view scale.
    // ###########################################################################################
    private void RefreshWorklogEntryBadge(Rect? finalPixelRectangle)
    {
        if (!finalPixelRectangle.HasValue || this.currentFullResBitmap == null)
        {
            if (this.thisWorklogEntryBadgeBorder != null)
            {
                this.thisWorklogEntryBadgeBorder.IsVisible = false;
            }
            return;
        }

        if (this.thisWorklogEntryBadgeBorder == null)
        {
            var textBlock = new TextBlock
            {
                FontSize = 11,
                FontWeight = FontWeight.Bold,
                Foreground = Brushes.White
            };

            this.thisWorklogEntryBadgeBorder = new Border
            {
                CornerRadius = new CornerRadius(10),
                Padding = new Thickness(8, 3),
                IsHitTestVisible = false,
                Child = textBlock,
                RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative)
            };

            this.thisWorklogEntryBadgeScaleTransform = new ScaleTransform(1.0, 1.0);
            this.thisWorklogEntryBadgeBorder.RenderTransform = this.thisWorklogEntryBadgeScaleTransform;

            this.SchematicsWorklogEntryBadgeCanvas.Children.Add(this.thisWorklogEntryBadgeBorder);
        }

        this.thisWorklogEntryBadgeBorder.Background = new SolidColorBrush(this.GetSelectedWorklogEntryCategoryColor());

        var contentRect = this.GetImageContentRect();
        var localRect = RectGeometry.PixelToLocalRect(finalPixelRectangle.Value, contentRect, this.currentFullResBitmap.PixelSize);

        double scale = this.schematicsMatrix.M11;
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;

        ((TextBlock)this.thisWorklogEntryBadgeBorder.Child!).Text = $"#{this.thisWorklogEntryNextId}";
        this.thisWorklogEntryBadgeScaleTransform!.ScaleX = inverseScale;
        this.thisWorklogEntryBadgeScaleTransform.ScaleY = inverseScale;

        this.thisWorklogEntryBadgeBorder.IsVisible = true;
        this.thisWorklogEntryBadgeBorder.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));

        Size unscaledSize = this.thisWorklogEntryBadgeBorder.DesiredSize;

        // Always the top-left corner, the same corner a SAVED entry's badge uses (see
        // PositionWorklogEntriesListBadge), so the badge does not jump when the entry is written
        // and the drawn area becomes a saved one. It used to be whichever corner the quick "New
        // fault" card was NOT anchored to; with that card gone there is nothing left to dodge.
        var anchorPoint = new Point(localRect.Left, localRect.Top);

        // Centred on its chosen corner, the same way the saved entries' badges are - see
        // PositionWorklogEntriesListBadge. Placing it at anchor - scaledSize/2 only lands correctly
        // at inverseScale 0.5; BadgeGeometry carries the derivation.
        var centreOffset = BadgeGeometry.GetCenterScaledCentreOffset(unscaledSize);

        Canvas.SetLeft(this.thisWorklogEntryBadgeBorder, anchorPoint.X + centreOffset.X);
        Canvas.SetTop(this.thisWorklogEntryBadgeBorder, anchorPoint.Y + centreOffset.Y);
    }

    // ###########################################################################################
    // Builds the component scope for a SAVED entry, for the full editor's copy of the checklist.
    //
    // The computation itself is WorklogEntryScope.BuildComponentsInScope, shared with the Workbooks
    // tab, which opens the same editor modal from its own pills - see that method for the null-vs-
    // empty rule, which matters and must not be collapsed. This wrapper only supplies the board data
    // and the highlight-rect cache; the two tabs read those from different places.
    // ###########################################################################################
    private List<(string BoardLabel, string DisplayName)>? BuildWorklogEntryComponentScope(WorklogEntryRecord entry) =>
        WorklogEntryScope.BuildComponentsInScope(
            this.MainWindow?.CurrentBoardData,
            this.highlightRectsBySchematicAndLabel,
            entry);

    // ###########################################################################################
    // Turns the "Show worklogs" list view on or off for the given workbook - called by the
    // top-bar checkbox. Turning it on rebuilds the overlay for the schematic currently on screen;
    // turning it off clears both the colored-area overlay and its badges/pills.
    // ###########################################################################################
    public void SetShowWorklogEntriesList(bool isShowing, int workbookId)
    {
        this.thisIsShowingWorklogEntriesList = isShowing;
        this.thisWorklogEntriesListWorkbookId = workbookId;

        this.RefreshWorklogEntriesListOverlay();
    }

    // ###########################################################################################
    // Redraws the worklog overlay and the thumbnail pills for the workbook ALREADY being shown,
    // without changing which workbook that is.
    //
    // Called from Main.RefreshWorklogBar on every worklog change. SetShowWorklogEntriesList above
    // covers the case where the shown workbook CHANGES, but an entry edited inside the workbook
    // already on screen changes neither of that method's two arguments - so before this existed,
    // editing an entry's area, or ticking "Show marked area" on it, left both the schematic overlay
    // and the thumbnail pills drawing the entry as it was before the edit until something else
    // happened to rebuild them. Reported as the marker not updating.
    //
    // Cheap when nothing is being shown: RefreshWorklogEntriesListOverlay's first branch clears and
    // returns without touching the disk.
    // ###########################################################################################
    public void RefreshWorklogEntriesListForCurrentWorkbook() =>
        this.RefreshWorklogEntriesListOverlay();

    // ###########################################################################################
    // Rebuilds the "Show worklogs" list view for the schematic currently on screen: every saved
    // entry whose SchematicName matches it, each drawn in its own category color, with a "#N"
    // badge and state pill anchored at its area's top-left. Entries are deliberately scoped to
    // the current schematic name (not the whole workbook) - a board can have entries recorded
    // against different schematics (e.g. top vs. bottom PCB side), and only the ones for what is
    // actually on screen should show.
    //
    // Also rebuilds every thumbnail's own state-pill overlay (ThumbnailWorklogPillsOverlay) from
    // the same workbook, one call covering both since they always change together - toggling
    // "Show worklogs" or switching workbook/board affects the main view and the thumbnails in the
    // same instant.
    // ###########################################################################################
    private void RefreshWorklogEntriesListOverlay()
    {
        foreach (var (badge, _, _, _) in this.thisWorklogEntriesListBadges)
        {
            this.SchematicsWorklogEntriesBadgeCanvas.Children.Remove(badge);
        }
        this.thisWorklogEntriesListBadges.Clear();

        this.ClearWorklogParkedBadges();

        // The hover label caches which entry it is painted for and skips repainting while that id
        // is unchanged. An entry edited in the full editor keeps its id, so without this the label
        // would keep showing the pre-edit title and state icon until the pointer moved to a
        // different pill. Resetting the appearance rather than only clearing the cached id matters:
        // the id records what is currently painted, so zeroing it while the label still carried an
        // entry's colors would make the next reset a no-op and strand those colors on the label.
        this.ResetSchematicsHoverLabelToDefaultAppearance();

        if (!this.thisIsShowingWorklogEntriesList || this.currentFullResBitmap == null)
        {
            this.SchematicsWorklogEntriesOverlay.IsVisible = false;
            this.SchematicsWorklogEntriesOverlay.Entries = Array.Empty<WorklogEntriesOverlay.Entry>();
            this.ClearThumbnailWorklogPills();
            return;
        }

        string schematicName = this.GetCurrentSchematicName();
        var allEntries = WorklogManager.GetEntries(this.thisWorklogEntriesListWorkbookId);
        // OrdinalIgnoreCase, matching how schematic names are keyed everywhere else in the app
        // (schematicByName, highlightIndexBySchematic, highlightRectsBySchematicAndLabel and the
        // Workbooks tab's own grouping are all OrdinalIgnoreCase). With Ordinal here, an entry saved
        // against "sheet 1" on a board whose schematic is named "Sheet 1" - a hand edit of
        // entries.json, or contributed board data - showed on the Workbooks tab and vanished on this
        // one: the same "the two views are not identical" class of bug already reported once.
        var entries = allEntries
            .Where(entry => string.Equals(entry.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        var contentRect = this.GetImageContentRect();
        double scale = this.schematicsMatrix.M11;
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;

        var overlayEntries = new List<WorklogEntriesOverlay.Entry>(entries.Count);

        foreach (var entry in entries)
        {
            var pixelRect = new Rect(entry.AreaX, entry.AreaY, entry.AreaWidth, entry.AreaHeight);
            Color color = this.ResolveWorklogCategoryColor(entry.Category);

            // "Show marked area" unticked: no coloured rectangle and no anchored badge. The pill is
            // parked in the corner instead - see LayOutWorklogParkedBadges.
            if (!entry.ShowMarkedArea)
            {
                this.CreateWorklogParkedBadge(entry, color);
                continue;
            }

            overlayEntries.Add(new WorklogEntriesOverlay.Entry(pixelRect, color, entry.Id));

            this.CreateWorklogEntriesListBadge(entry, color, pixelRect, contentRect, inverseScale);
        }

        this.LayOutWorklogParkedBadges();

        this.SchematicsWorklogEntriesOverlay.BitmapPixelSize = this.currentFullResBitmap.PixelSize;
        this.SchematicsWorklogEntriesOverlay.ViewMatrix = this.schematicsMatrix;
        this.SchematicsWorklogEntriesOverlay.Entries = overlayEntries;
        this.SchematicsWorklogEntriesOverlay.IsVisible = true;
        this.SchematicsWorklogEntriesOverlay.InvalidateVisual();

        this.RefreshThumbnailWorklogPills(allEntries);
    }

    // ###########################################################################################
    // Resizing a saved worklog entry's marked area directly on the schematic.
    //
    // Hovering a marked area shows the same corner/side markers the component label editor puts on
    // a selected highlight, and dragging one resizes the area; dragging its interior moves it. The
    // new bounds are written straight back to the entry's workbook on release, so the change is
    // saved without opening the editor.
    //
    // Only available while "Show worklogs" is on - that is when the areas are drawn, and an
    // invisible drag target would be a trap. It is also skipped while the label editor or the
    // new-entry card is active, so those modes keep exclusive use of the pointer.
    // ###########################################################################################
    private int thisWorklogResizeEntryId = -1;

    private LabelEditorDragMode thisWorklogResizeDragMode = LabelEditorDragMode.None;

    private Point thisWorklogResizeStartPixelPoint;

    private Rect thisWorklogResizeOriginalRect;

    private bool thisIsResizingWorklogEntry;

    // Small enough that a deliberately tiny area is still allowed, large enough that an area
    // cannot be shrunk to something impossible to grab again.
    private const double MinimumWorklogAreaSize = 8.0;

    // thisIsWorklogEntryMode keeps resize hit-testing out of the way while a NEW area is being
    // drawn - otherwise the drag paints resize markers and a directional cursor over the crosshair
    // the drawing mode set. (It used to also check the quick card's visibility; that card is gone,
    // and the full editor that replaced it is a modal window, so nothing on this canvas can be
    // hovered while it is up.)
    private bool IsWorklogEntryResizeAvailable =>
        this.thisIsShowingWorklogEntriesList &&
        this.currentFullResBitmap != null &&
        !this.thisIsWorklogEntryMode &&
        !this.IsLabelEditorActive;

    // ###########################################################################################
    // Finds the marked area under the pointer and which of its handles, if any, is being touched.
    //
    // Entries are tested in reverse draw order so the one drawn last - the one on top - wins when
    // two overlap, matching what the user sees. The handle hit rectangles come from
    // LabelEditorGeometry, so they sit in exactly the same places as the label editor's.
    // ###########################################################################################
    private bool TryGetWorklogEntryHandleAt(
        Point pointerInContainer,
        out int entryId,
        out LabelEditorDragMode dragMode,
        out Rect pixelRect)
    {
        entryId = -1;
        dragMode = LabelEditorDragMode.None;
        pixelRect = default;

        if (!this.IsWorklogEntryResizeAvailable)
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
        if (contentRect.Width <= 0 || contentRect.Height <= 0)
        {
            return false;
        }

        double scale = Math.Max(0.0001, this.schematicsMatrix.M11);
        var entries = this.SchematicsWorklogEntriesOverlay.Entries;

        // EVERY area's handles are tested before ANY area's interior.
        //
        // Testing each entry completely in turn made handles win only within a single entry: a
        // small area sitting inside a larger one drawn later had its corner handles swallowed by
        // the larger area's interior, so it could never be grabbed. Two passes make "a handle
        // always beats an interior" true across entries, which is what the user sees - the handle
        // markers are drawn on top.
        //
        // Both passes run in reverse draw order, so among equals the area drawn last (on top) wins.
        for (int i = entries.Count - 1; i >= 0; i--)
        {
            var entry = entries[i];
            var localRect = RectGeometry.PixelToLocalRect(entry.PixelRect, contentRect, this.currentFullResBitmap!.PixelSize);

            foreach (var (hitRect, handleMode) in LabelEditorGeometry.BuildLabelEditorHandleHitRects(localRect, scale))
            {
                if (hitRect.Contains(localPoint))
                {
                    entryId = entry.EntryId;
                    dragMode = handleMode;
                    pixelRect = entry.PixelRect;
                    return true;
                }
            }
        }

        for (int i = entries.Count - 1; i >= 0; i--)
        {
            var entry = entries[i];
            var localRect = RectGeometry.PixelToLocalRect(entry.PixelRect, contentRect, this.currentFullResBitmap!.PixelSize);

            // Inside the area but not on a handle: a move rather than a resize. Whether the
            // interior actually belongs to the area is decided by the callers, which give a
            // component underneath precedence - see UpdateWorklogEntryResizeHover.
            if (localRect.Contains(localPoint))
            {
                entryId = entry.EntryId;
                dragMode = LabelEditorDragMode.Move;
                pixelRect = entry.PixelRect;
                return true;
            }
        }

        return false;
    }

    // ###########################################################################################
    // Updates the hovered area's markers and the pointer cursor. Returns true when the pointer is
    // over an area, so the caller can leave the schematic's normal hover handling alone.
    // ###########################################################################################
    private bool UpdateWorklogEntryResizeHover(Point pointerInContainer)
    {
        if (this.thisIsResizingWorklogEntry)
        {
            return true;
        }

        if (!this.TryGetWorklogEntryHandleAt(pointerInContainer, out int entryId, out var dragMode, out _))
        {
            this.ClearWorklogEntryResizeHover();
            return false;
        }

        // A component underneath wins the INTERIOR of a marked area.
        //
        // A worklog area is often drawn around the very component it is about, so if the interior
        // always belonged to the area, that component could never be hovered or selected again -
        // the area would swallow it entirely. The edges do not have that problem: they are a thin
        // band, and the user is deliberately aiming at them.
        //
        // So the handles always win (there is nothing else to mean at the edge of an area being
        // resized), and the interior defers whenever a component is actually there. Where no
        // component is underneath, the interior still moves the area - so dragging one around
        // empty board space keeps working.
        if (dragMode == LabelEditorDragMode.Move &&
            this.TryGetHoveredBoardLabel(pointerInContainer, out _, out _))
        {
            this.ClearWorklogEntryResizeHover();
            return false;
        }

        this.SchematicsWorklogEntriesOverlay.HoveredEntryId = entryId;
        this.SetWorklogResizeCursor(dragMode);
        return true;
    }

    // ###########################################################################################
    // Sets the directional cursor, reusing one Cursor instance per drag mode.
    //
    // This runs on every pointer-move frame while an area is hovered, so allocating a fresh Cursor
    // each time meant hundreds of platform handles a second for a value that only takes nine
    // distinct forms. The cache also lets the setter skip the assignment when nothing changed,
    // which keeps it from fighting whatever else last wrote the cursor.
    // ###########################################################################################
    private readonly Dictionary<StandardCursorType, Cursor> thisWorklogResizeCursors = new();

    private StandardCursorType? thisAppliedWorklogResizeCursor;

    private void SetWorklogResizeCursor(LabelEditorDragMode dragMode)
    {
        var cursorType = ResolveWorklogResizeCursor(dragMode);

        if (this.thisAppliedWorklogResizeCursor == cursorType)
        {
            return;
        }

        if (!this.thisWorklogResizeCursors.TryGetValue(cursorType, out var cursor))
        {
            cursor = new Cursor(cursorType);
            this.thisWorklogResizeCursors[cursorType] = cursor;
        }

        this.SchematicsContainer.Cursor = cursor;
        this.thisAppliedWorklogResizeCursor = cursorType;
    }

    // Drops the markers and the directional cursor. Guarded so it does not fight whatever else is
    // setting the cursor when no area was hovered in the first place.
    private void ClearWorklogEntryResizeHover()
    {
        this.SchematicsWorklogEntriesOverlay.HoveredEntryId = -1;

        // Only reset the cursor if THIS feature is what last set it. The old guard was on the
        // hovered id, not on ownership, so leaving an area onto a polyline or a KiCad hover reset
        // the cursor those modes had just set, one frame later.
        if (this.thisAppliedWorklogResizeCursor != null)
        {
            this.SchematicsContainer.Cursor = Cursor.Default;
            this.thisAppliedWorklogResizeCursor = null;
        }
    }

    // The standard directional cursors, so the pointer says which way the edge will move.
    private static StandardCursorType ResolveWorklogResizeCursor(LabelEditorDragMode dragMode) => dragMode switch
    {
        LabelEditorDragMode.ResizeTopLeft => StandardCursorType.TopLeftCorner,
        LabelEditorDragMode.ResizeTopRight => StandardCursorType.TopRightCorner,
        LabelEditorDragMode.ResizeBottomLeft => StandardCursorType.BottomLeftCorner,
        LabelEditorDragMode.ResizeBottomRight => StandardCursorType.BottomRightCorner,
        LabelEditorDragMode.ResizeTop => StandardCursorType.SizeNorthSouth,
        LabelEditorDragMode.ResizeBottom => StandardCursorType.SizeNorthSouth,
        LabelEditorDragMode.ResizeLeft => StandardCursorType.SizeWestEast,
        LabelEditorDragMode.ResizeRight => StandardCursorType.SizeWestEast,
        _ => StandardCursorType.SizeAll,
    };

    // ###########################################################################################
    // Begins a resize or move. Returns true when one started, so the caller can stop the press
    // reaching pan/selection.
    // ###########################################################################################
    private bool TryBeginWorklogEntryResize(Point pointerInContainer)
    {
        if (!this.TryGetWorklogEntryHandleAt(pointerInContainer, out int entryId, out var dragMode, out var pixelRect))
        {
            return false;
        }

        // Same precedence as the hover above: a press on the interior belongs to the component
        // underneath if there is one, so a component covered by a marked area stays selectable.
        if (dragMode == LabelEditorDragMode.Move &&
            this.TryGetHoveredBoardLabel(pointerInContainer, out _, out _))
        {
            return false;
        }

        // Unbounded so the drag origin is expressed in the same space every later move uses.
        if (!this.TryGetSchematicsImagePixelPointUnbounded(pointerInContainer, out var pixelPoint))
        {
            return false;
        }

        this.thisWorklogResizeEntryId = entryId;
        this.thisWorklogResizeDragMode = dragMode;
        this.thisWorklogResizeStartPixelPoint = pixelPoint;
        this.thisWorklogResizeOriginalRect = pixelRect;
        this.thisIsResizingWorklogEntry = true;

        return true;
    }

    // ###########################################################################################
    // Applies the in-progress drag to the overlay only. Nothing is written to disk until release,
    // so an abandoned drag costs no saves and an interrupted one cannot leave a half-applied area.
    // ###########################################################################################
    private void UpdateWorklogEntryResize(Point pointerInContainer)
    {
        if (!this.thisIsResizingWorklogEntry || this.currentFullResBitmap == null)
        {
            return;
        }

        // Unbounded: a drag that leaves the image must keep tracking the pointer, with the RESULT
        // clamped below. The bounded variant returns false past the edge, which froze the area at
        // its last in-bounds size and made ClampRectToBounds unreachable on the very path it names.
        if (!this.TryGetSchematicsImagePixelPointUnbounded(pointerInContainer, out var pixelPoint))
        {
            return;
        }

        double dx = pixelPoint.X - this.thisWorklogResizeStartPixelPoint.X;
        double dy = pixelPoint.Y - this.thisWorklogResizeStartPixelPoint.Y;

        var resized = LabelEditorGeometry.ResizeRect(
            this.thisWorklogResizeOriginalRect, this.thisWorklogResizeDragMode, dx, dy, MinimumWorklogAreaSize);

        var bitmapSize = new Size(this.currentFullResBitmap.PixelSize.Width, this.currentFullResBitmap.PixelSize.Height);

        // A Move slides the whole area back inside, keeping its size; a RESIZE trims only the edge
        // that strayed out, so the edge the user is not dragging stays anchored. Using the move
        // clamp for both pushed the opposite edge outward whenever a resize hit the board boundary.
        resized = this.thisWorklogResizeDragMode == LabelEditorDragMode.Move
            ? LabelEditorGeometry.ClampRectToBounds(resized, bitmapSize)
            : LabelEditorGeometry.ClampResizedRectToBounds(resized, bitmapSize, MinimumWorklogAreaSize);

        var updated = new List<WorklogEntriesOverlay.Entry>(this.SchematicsWorklogEntriesOverlay.Entries.Count);

        foreach (var entry in this.SchematicsWorklogEntriesOverlay.Entries)
        {
            updated.Add(entry.EntryId == this.thisWorklogResizeEntryId
                ? entry with { PixelRect = resized }
                : entry);
        }

        this.SchematicsWorklogEntriesOverlay.Entries = updated;
        this.RepositionWorklogEntriesListBadge(this.thisWorklogResizeEntryId, resized);
    }

    // ###########################################################################################
    // Ends the drag and persists the new bounds.
    //
    // The overlay is rebuilt from disk afterwards rather than trusted: if the save failed, the area
    // must snap back to what is actually stored rather than showing bounds nothing recorded.
    // ###########################################################################################
    private void CompleteWorklogEntryResize()
    {
        if (!this.thisIsResizingWorklogEntry)
        {
            return;
        }

        int entryId = this.thisWorklogResizeEntryId;
        var originalRect = this.thisWorklogResizeOriginalRect;

        this.thisIsResizingWorklogEntry = false;
        this.thisWorklogResizeEntryId = -1;
        this.thisWorklogResizeDragMode = LabelEditorDragMode.None;

        Rect? finalRect = null;
        foreach (var entry in this.SchematicsWorklogEntriesOverlay.Entries)
        {
            if (entry.EntryId == entryId)
            {
                finalRect = entry.PixelRect;
                break;
            }
        }

        if (finalRect == null || entryId < 0)
        {
            return;
        }

        // Unchanged bounds mean the user clicked rather than dragged - no save, so a stray click on
        // an area does not rewrite its workbook.
        //
        // The overlay is still rebuilt: a drag that wandered and came back (or was pinned by the
        // minimum or the board edge) has already written intermediate rects into the overlay and
        // moved the badge, and nothing else would undo them. Skipping the refresh here left what is
        // drawn disagreeing with what is on disk until some unrelated redraw happened.
        if (AreWorklogRectsEquivalent(finalRect.Value, originalRect))
        {
            this.RefreshWorklogEntriesListOverlay();
            return;
        }

        var record = WorklogManager.GetEntries(this.thisWorklogEntriesListWorkbookId)
            .FirstOrDefault(x => x.Id == entryId);

        if (record == null)
        {
            this.RefreshWorklogEntriesListOverlay();
            return;
        }

        record.AreaX = finalRect.Value.X;
        record.AreaY = finalRect.Value.Y;
        record.AreaWidth = finalRect.Value.Width;
        record.AreaHeight = finalRect.Value.Height;

        // A resize can only ever REMOVE components from the selection.
        //
        // Shrinking an area off a component the user had marked in scope leaves that component
        // recorded against a fault whose area no longer covers it, so it is dropped. Growing an
        // area over new components does NOT tick them: being inside the rectangle is not the same
        // as the user deciding they are relevant, and auto-ticking would quietly add components
        // nobody chose - more of them the wider the area is dragged. Adding stays a deliberate act
        // in the full editor's checklist. See ComponentListBuilder.NarrowSelectionToScope.
        //
        // Skipped when the scope cannot be determined (no board data, or no highlight rectangles
        // for this schematic): an unknown scope is not an empty one, and treating it as such would
        // wipe the whole selection.
        var scopeAfterResize = this.BuildWorklogEntryComponentScope(record);
        if (scopeAfterResize != null)
        {
            // Both lists narrowed together, so the completed list is measured against the scope
            // that REMAINS rather than against the raw area - see NarrowEntryToScope for why those
            // are not the same thing.
            (record.ComponentLabels, record.CompletedComponentLabels) = ComponentListBuilder.NarrowEntryToScope(
                record.ComponentLabels,
                record.CompletedComponentLabels,
                scopeAfterResize.Select(c => c.BoardLabel).ToList());
        }

        if (!WorklogManager.UpdateEntry(this.thisWorklogEntriesListWorkbookId, record))
        {
            Logger.Warning($"Failed to save resized worklog entry area [#{entryId}]");
        }

        // Rebuilt either way - on success to pick up anything the save normalised, on failure to
        // snap the area back to what is genuinely on disk.
        this.RefreshWorklogEntriesListOverlay();
        this.MainWindow?.RefreshWorklogBar();
    }

    // Sub-pixel differences come from the pointer maths, not from the user, so they do not count
    // as a change worth saving.
    private static bool AreWorklogRectsEquivalent(Rect a, Rect b) =>
        Math.Abs(a.X - b.X) < 0.5 &&
        Math.Abs(a.Y - b.Y) < 0.5 &&
        Math.Abs(a.Width - b.Width) < 0.5 &&
        Math.Abs(a.Height - b.Height) < 0.5;

    // ###########################################################################################
    // Moves an entry's "#N" badge to follow its area during a drag, so the badge does not sit at
    // the old corner until the drag ends.
    // ###########################################################################################
    private void RepositionWorklogEntriesListBadge(int entryId, Rect pixelRect)
    {
        if (this.currentFullResBitmap == null)
        {
            return;
        }

        var contentRect = this.GetImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0)
        {
            return;
        }

        double scale = this.schematicsMatrix.M11;
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;

        for (int i = 0; i < this.thisWorklogEntriesListBadges.Count; i++)
        {
            var (badge, _, badgeEntry, color) = this.thisWorklogEntriesListBadges[i];

            if (badgeEntry.Id != entryId)
            {
                continue;
            }

            // Positioned through the shared method, NOT by writing Canvas.SetLeft directly. That
            // shortcut dropped both the centre offset and the viewport nudge, so the badge jumped
            // by half its own size the moment a drag began and ignored the view edge for the whole
            // gesture - the exact drift BadgeGeometry was extracted to stop, reintroduced at the one
            // call site that bypassed it.
            this.PositionWorklogEntriesListBadge(badge, pixelRect, contentRect, inverseScale);

            // The cache carries each badge's anchor rect, and RescaleWorklogEntriesListBadges
            // repositions from it on every zoom/pan tick. Left stale, a wheel-zoom mid-drag would
            // snap the badge back to the pre-resize corner and strand it there until release.
            this.thisWorklogEntriesListBadges[i] = (badge, pixelRect, badgeEntry, color);
            break;
        }
    }

    // ###########################################################################################
    // Tears down every worklog visual on the schematic surface. Called by ResetSchematicsViewer on
    // a board switch, alongside the KiCad/label/polyline teardown it already did.
    //
    // Without this the badges stayed as live children of SchematicsWorklogEntriesBadgeCanvas after
    // the board changed - still carrying their entry ids and click handlers, drawn over the new
    // board, opening the editor for an entry belonging to the previous one. RefreshWorklogBar does
    // not necessarily clean up, because it only re-syncs when the workbook id actually differs.
    // ###########################################################################################
    public void ResetWorklogOverlays()
    {
        foreach (var (badge, _, _, _) in this.thisWorklogEntriesListBadges)
        {
            this.SchematicsWorklogEntriesBadgeCanvas.Children.Remove(badge);
        }
        this.thisWorklogEntriesListBadges.Clear();

        // Parked pills live on their own canvas, so tearing down the anchored ones leaves these
        // behind - stranded pills for a board that is no longer on screen.
        this.ClearWorklogParkedBadges();

        this.SchematicsWorklogEntriesOverlay.IsVisible = false;
        this.SchematicsWorklogEntriesOverlay.Entries = Array.Empty<WorklogEntriesOverlay.Entry>();

        if (this.thisWorklogEntryBadgeBorder != null)
        {
            this.thisWorklogEntryBadgeBorder.IsVisible = false;
        }

        this.ClearThumbnailWorklogPills();
    }

    // ###########################################################################################
    // Clears every thumbnail's state-pill overlay - used when "Show worklogs" is off or there is
    // no workbook to show.
    // ###########################################################################################
    private void ClearThumbnailWorklogPills()
    {
        foreach (var thumbnail in this.currentThumbnails)
        {
            if (thumbnail.WorklogPills.Count > 0)
            {
                thumbnail.WorklogPills = Array.Empty<ThumbnailWorklogPillsOverlay.Pill>();
            }
        }
    }

    // ###########################################################################################
    // Assigns each thumbnail its own "#N" pills - one per saved entry whose SchematicName matches
    // that thumbnail, colored by category - see ThumbnailWorklogPillsOverlay for why status is
    // deliberately left off. A board's entries can span several schematics, so this groups once
    // across all of them rather than only the schematic currently on screen.
    //
    // The "Show marked area" branch is carried through as IsParked, so the thumbnail draws the
    // entry the SAME way the main view does: ticked gets a pill centred on its marked area,
    // unticked gets one parked in the image's own top-right corner. Passing the marker position
    // regardless was a reported bug - an entry whose area is hidden showed no area on either view,
    // but the thumbnail still pointed a pill at where that area would have been.
    // ###########################################################################################
    private void RefreshThumbnailWorklogPills(List<WorklogEntryRecord> allEntries)
    {
        var entriesBySchematic = allEntries
            .GroupBy(entry => entry.SchematicName, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);

        foreach (var thumbnail in this.currentThumbnails)
        {
            if (!entriesBySchematic.TryGetValue(thumbnail.Name, out var thumbnailEntries))
            {
                if (thumbnail.WorklogPills.Count > 0)
                {
                    thumbnail.WorklogPills = Array.Empty<ThumbnailWorklogPillsOverlay.Pill>();
                }
                continue;
            }

            var pills = new List<ThumbnailWorklogPillsOverlay.Pill>(thumbnailEntries.Count);
            foreach (var entry in thumbnailEntries)
            {
                var center = new Point(
                    entry.AreaX + (entry.AreaWidth / 2.0),
                    entry.AreaY + (entry.AreaHeight / 2.0));

                pills.Add(new ThumbnailWorklogPillsOverlay.Pill(
                    center,
                    this.ResolveWorklogCategoryColor(entry.Category),
                    entry.Id,
                    IsParked: !entry.ShowMarkedArea));
            }

            thumbnail.WorklogPills = pills;
        }
    }

    // ###########################################################################################
    // Resolves a category's theme color by name - the same Worklog_Category_* resources the "New
    // fault" card's chips use (GetSelectedWorklogEntryCategoryColor covers only the category
    // currently being edited, not an arbitrary saved one, hence this separate overload).
    // ###########################################################################################
    private Color ResolveWorklogCategoryColor(string category)
    {
        var brush = this.ResolveThemeBrush($"Worklog_Category_{category}", new SolidColorBrush(Colors.IndianRed));
        return brush is ISolidColorBrush solidBrush ? solidBrush.Color : Colors.IndianRed;
    }

    // ###########################################################################################
    // Resolves a saved entry's state pill color - the same Worklog_Status_Open/
    // Worklog_Status_Closed theme resources the worklog bar's own workbook-status pill uses, so an
    // entry's Open/Closed and a workbook's Open/Closed always render identically. They are two
    // different axes wearing one palette on purpose: both read "Open"/"Closed" to the user, and
    // showing them in different colours would imply a distinction that does not exist.
    //
    // Anything unrecognised falls through to Open: state is a free-form string in entries.json, so
    // a hand-edited or future value must still render rather than throw.
    // ###########################################################################################
    private Color ResolveWorklogStateColor(string state)
    {
        string stateColorResourceKey = state switch
        {
            WorklogStateClosed => "Worklog_Status_Closed",
            _ => "Worklog_Status_Open",
        };

        var brush = this.ResolveThemeBrush(stateColorResourceKey, new SolidColorBrush(Colors.IndianRed));
        return brush is ISolidColorBrush solidBrush ? solidBrush.Color : Colors.IndianRed;
    }

    // ###########################################################################################
    // Builds one "#N" badge + state pill for the "Show worklogs" list view, anchored at the top-
    // left of the entry's marked area, and adds it to SchematicsWorklogEntriesBadgeCanvas. Kept a
    // constant screen size across zoom via a ScaleTransform, the same technique
    // RefreshWorklogEntryBadge uses for the single in-progress draft badge.
    // ###########################################################################################
    private void CreateWorklogEntriesListBadge(WorklogEntryRecord entry, Color color, Rect pixelRect, Rect contentRect, double inverseScale)
    {
        var badge = this.CreateWorklogBadgeControl(entry, color, inverseScale);

        this.SchematicsWorklogEntriesBadgeCanvas.Children.Add(badge);
        this.thisWorklogEntriesListBadges.Add((badge, pixelRect, entry, color));

        this.PositionWorklogEntriesListBadge(badge, pixelRect, contentRect, inverseScale);
    }

    // ###########################################################################################
    // Builds the pill for an entry whose marked area is hidden and adds it to the parked canvas.
    // Position is left to LayOutWorklogParkedBadges, which needs the whole set to stack them.
    //
    // Scale 1, unlike the anchored badges: those sit on a canvas carrying the view matrix and use
    // an inverse scale to cancel it out, so they hold a constant screen size while the board zooms.
    // The parked canvas has no such transform, so the same compensation would shrink or magnify
    // these pills for no reason.
    // ###########################################################################################
    private void CreateWorklogParkedBadge(WorklogEntryRecord entry, Color color)
    {
        var badge = this.CreateWorklogBadgeControl(entry, color, 1.0);

        this.SchematicsWorklogParkedBadgeCanvas.Children.Add(badge);
        this.thisWorklogParkedBadges.Add((badge, entry));
    }

    private void ClearWorklogParkedBadges()
    {
        foreach (var (badge, _) in this.thisWorklogParkedBadges)
        {
            this.SchematicsWorklogParkedBadgeCanvas.Children.Remove(badge);
        }

        this.thisWorklogParkedBadges.Clear();
    }

    // ###########################################################################################
    // Arranges the parked pills as a compact block in the schematic panel's top-right corner,
    // stepping left out of the "Netlist names" panel's way whenever that panel is open. A block
    // rather than one long column - see ParkedBadgeGeometry for the row/column progression.
    //
    // The reservation is read from the panel's ACTUAL laid-out width rather than from its MaxWidth,
    // because the panel sizes itself to its content: a short net name gives a narrow panel, and
    // reserving the maximum would leave the pills floating in empty space most of the time. Its
    // margin is added on so the pills clear the panel rather than touching it.
    //
    // Called on refresh AND on every zoom/pan tick. The pills do not move with the board, but the
    // panel can open, close or change width underneath them, and the viewport itself can resize.
    // ###########################################################################################
    private bool thisIsLayingOutParkedBadges;

    private void LayOutWorklogParkedBadges()
    {
        if (this.thisWorklogParkedBadges.Count == 0)
        {
            return;
        }

        // Re-entrancy guard. This method Measures every parked badge and writes Canvas.SetLeft/Top,
        // which invalidates the canvas's arrange; when that settles, the container's own Bounds can
        // change and fire the very PropertyChanged handler that called us. The property filter in
        // that handler screens out unrelated properties but says nothing about a pass this method
        // triggers itself, so without this a single resize could run the measure loop twice over.
        if (this.thisIsLayingOutParkedBadges)
        {
            return;
        }

        this.thisIsLayingOutParkedBadges = true;
        try
        {

            var viewportSize = this.SchematicsContainer.Bounds.Size;
            if (viewportSize.Width <= 0 || viewportSize.Height <= 0)
            {
                return;
            }

            var sizes = new List<Size>(this.thisWorklogParkedBadges.Count);
            foreach (var (badge, _) in this.thisWorklogParkedBadges)
            {
                badge.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                sizes.Add(badge.DesiredSize);
            }

            var positions = ParkedBadgeGeometry.ArrangeInTopRightBlock(
                sizes,
                viewportSize,
                WorklogParkedBadgeMargin,
                WorklogParkedBadgeSpacing,
                this.GetWorklogParkedBadgeReservedRight());

            for (int i = 0; i < this.thisWorklogParkedBadges.Count && i < positions.Count; i++)
            {
                Canvas.SetLeft(this.thisWorklogParkedBadges[i].Badge, positions[i].X);
                Canvas.SetTop(this.thisWorklogParkedBadges[i].Badge, positions[i].Y);
            }
        }
        finally
        {
            this.thisIsLayingOutParkedBadges = false;
        }
    }

    // ###########################################################################################
    // Re-stacks the parked pills when something they are positioned against moves: the "Netlist
    // names" panel appearing, disappearing or resizing, or the schematic container itself being
    // resized by the window or the splitter.
    //
    // Filtered to the two properties that actually matter. PropertyChanged fires for every styled
    // property on these controls - a re-layout on each would run the stack maths hundreds of times
    // per second during a drag for no visible difference.
    // ###########################################################################################
    private void OnWorklogParkedBadgeLayoutTriggerChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
    {
        if (e.Property != Visual.BoundsProperty && e.Property != Visual.IsVisibleProperty)
        {
            return;
        }

        this.LayOutWorklogParkedBadges();
    }

    // The width the "Netlist names" panel is currently claiming on the right-hand edge, including
    // its own margin - zero when it is not showing.
    private double GetWorklogParkedBadgeReservedRight()
    {
        if (!this.KiCadNetConnectionsPanel.IsVisible)
        {
            return 0.0;
        }

        double width = this.KiCadNetConnectionsPanel.Bounds.Width;
        if (width <= 0)
        {
            return 0.0;
        }

        // Right margin only. The panel is right-aligned, so the space it claims on that edge is its
        // width plus the gap between it and the window edge; its LEFT margin is on the board-facing
        // side - the gap the pills are meant to sit beside, not extra to reserve. Counting both put
        // the block 10px further left than needed, doubling the intended gap.
        return width + this.KiCadNetConnectionsPanel.Margin.Right;
    }

    // ###########################################################################################
    // The pill control itself - shared by the anchored badges and the parked ones so the two can
    // never drift apart visually. An entry that is merely parked must still look like the same
    // worklog it was when its area was showing.
    // ###########################################################################################
    private Border CreateWorklogBadgeControl(WorklogEntryRecord entry, Color color, double inverseScale)
    {
        // The visual comes from WorklogBadgeBuilder, shared with the Workbooks tab's board pane -
        // the two used to be line-for-line copies of one another, each conceding the duplication and
        // asserting the two "must look the same" with nothing enforcing it.
        var badge = WorklogBadgeBuilder.Build(entry, color, this.ResolveWorklogStateColor(entry.State));

        // What is genuinely this tab's own: these badges sit on a canvas carrying the view matrix, so
        // a centred inverse scale cancels it out and keeps them a constant size on screen while the
        // board zooms. The Workbooks pane never zooms and applies none of this.
        badge.RenderTransformOrigin = new RelativePoint(0.5, 0.5, RelativeUnit.Relative);
        badge.RenderTransform = new ScaleTransform(inverseScale, inverseScale);
        badge.Tag = entry.Id;
        badge.PointerPressed += this.OnWorklogEntryPillPointerPressed;

        return badge;
    }

    // ###########################################################################################
    // Positions and scales one "Show worklogs" badge for the given inverse-scale, CENTRED on the
    // top-left corner of its entry's marked area at EVERY zoom level - shared by
    // CreateWorklogEntriesListBadge (first layout) and RescaleWorklogEntriesListBadges (every later
    // zoom/pan tick).
    //
    // Centred, so the badge straddles the corner with roughly a quarter of it outside the marked
    // area. That reads as a label attached to the area rather than a box sitting inside it, and it
    // keeps the badge clear of whatever the area's top-left corner is drawn over.
    //
    // The offset comes from BadgeGeometry rather than being open-coded: Canvas.SetLeft/SetTop place
    // the badge's PRE-transform layout box, and the badge carries a centred ScaleTransform
    // (RenderTransformOrigin 0.5,0.5) that keeps it a constant size on screen while the board
    // scales. Getting that interaction wrong is what made the badge slide away from its corner as
    // the user zoomed - by nearly a badge-width at high zoom - so the maths and its reasoning live
    // together in one tested place.
    // ###########################################################################################
    private void PositionWorklogEntriesListBadge(Border badge, Rect pixelRect, Rect contentRect, double inverseScale)
    {
        ((ScaleTransform)badge.RenderTransform!).ScaleX = inverseScale;
        ((ScaleTransform)badge.RenderTransform).ScaleY = inverseScale;

        var localRect = RectGeometry.PixelToLocalRect(pixelRect, contentRect, this.currentFullResBitmap!.PixelSize);

        badge.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
        Size unscaledSize = badge.DesiredSize;

        var offset = BadgeGeometry.GetCenterScaledCentreOffset(unscaledSize);

        double left = localRect.Left + offset.X;
        double top = localRect.Top + offset.Y;

        var nudge = this.GetWorklogBadgeViewportNudge(new Point(left, top), unscaledSize, inverseScale);

        Canvas.SetLeft(badge, left + nudge.X);
        Canvas.SetTop(badge, top + nudge.Y);
    }

    // ###########################################################################################
    // How far a badge must move to stay fully inside the visible viewport, in the badge canvas's
    // own (pre-transform) coordinates.
    //
    // A badge whose area sits near the edge of the view has half of itself off-screen - its "#N"
    // unreadable and its click target unreachable - because the badges straddle their corner. This
    // pushes it back in by exactly the overhang, so it stays against the edge it belongs to.
    //
    // The conversion matters: Canvas.SetLeft works in the canvas's local space, but "visible" is a
    // property of SchematicsContainer, and the canvas carries the same zoom/pan matrix as the
    // image. So the badge's rendered rect is mapped INTO container space, clamped there against
    // the container's bounds, and the resulting adjustment mapped back out. Clamping in local space
    // instead would use the wrong units and drift with zoom - the same class of mistake that made
    // the badges slide off their corners.
    // ###########################################################################################
    private Point GetWorklogBadgeViewportNudge(Point layoutTopLeft, Size unscaledSize, double inverseScale)
    {
        var viewportSize = this.SchematicsContainer.Bounds.Size;
        if (viewportSize.Width <= 0 || viewportSize.Height <= 0)
        {
            return new Point(0, 0);
        }

        // The badge holds a constant SCREEN size, so its on-screen extent is the unscaled size -
        // the ScaleTransform and the view matrix cancel out.
        var renderedTopLeftLocal = BadgeGeometry.GetCenterScaledRenderedTopLeft(layoutTopLeft, unscaledSize, inverseScale);

        var matrix = this.schematicsMatrix;
        var topLeftInContainer = new Point(
            (renderedTopLeftLocal.X * matrix.M11) + (renderedTopLeftLocal.Y * matrix.M21) + matrix.M31,
            (renderedTopLeftLocal.X * matrix.M12) + (renderedTopLeftLocal.Y * matrix.M22) + matrix.M32);

        var renderedInContainer = new Rect(topLeftInContainer, unscaledSize);

        var nudge = BadgeGeometry.GetViewportNudge(renderedInContainer, viewportSize, WorklogBadgeViewportMargin);
        if (nudge.X == 0 && nudge.Y == 0)
        {
            return nudge;
        }

        // Back into canvas space. Only the scale matters for a delta - the translation cancels.
        double scale = matrix.M11;
        if (scale <= 0)
        {
            return new Point(0, 0);
        }

        return new Point(nudge.X / scale, nudge.Y / scale);
    }

    // A small inset so a nudged badge sits just clear of the edge rather than flush against it.
    private const double WorklogBadgeViewportMargin = 2.0;

    // ###########################################################################################
    // Hit-tests the "Show worklogs" list view's own "#N" pills (not the marked area they are
    // anchored to) at the given container point. Checked back-to-front so an overlapping pill
    // drawn later (on top) wins. Used by UpdateSchematicsHoverUi to give a worklog pill's hover
    // info priority over the component highlight underneath it - the two can and do overlap,
    // since a fault is typically marked right on top of the component it concerns, but only the
    // small pill itself should trigger the swap, not the whole marked area.
    //
    // BOTH corners are translated, not a corner plus Bounds.Size: each pill carries a centered
    // ScaleTransform (PositionWorklogEntriesListBadge keeps it a constant screen size across
    // zoom), and Bounds is the pre-transform layout size. Pairing a translated corner with an
    // untransformed size gave a rect that was both mis-sized and offset at any zoom other than
    // 100% - too big above it (hover triggering off empty space beside the pill), too small
    // below it (dead edges). Translating both corners lets the visual tree apply the pill's
    // ScaleTransform and the canvas MatrixTransform to each, so the rect matches what is drawn
    // without this method re-deriving either.
    //
    // The corners are (0,0) and (Width,Height) - badge-LOCAL coordinates, because TranslatePoint
    // reads its argument in the source control's own space. Passing Bounds.TopLeft/BottomRight
    // instead looks equivalent but is not: Bounds is expressed in the PARENT's space, so for a
    // badge at Canvas.SetLeft(120)/SetTop(80) its TopLeft is (120,80) rather than (0,0) and the
    // translation adds the pill's position a second time, landing the rect at roughly twice its
    // real offset and killing hover everywhere. GetWorklogEntryAreaBoundsInContainer looks like
    // the counter-example but is not - the rect it translates is already in the space of the
    // overlay it calls TranslatePoint on.
    // ###########################################################################################
    private bool TryGetHoveredWorklogEntry(Point pointerInContainer, out WorklogEntryRecord entry, out Color color)
    {
        entry = null!;
        color = default;

        if (!this.thisIsShowingWorklogEntriesList || this.thisWorklogEntriesListBadges.Count == 0)
        {
            return false;
        }

        for (int i = this.thisWorklogEntriesListBadges.Count - 1; i >= 0; i--)
        {
            var (badge, _, candidateEntry, candidateColor) = this.thisWorklogEntriesListBadges[i];

            Point? topLeft = badge.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
            Point? bottomRight = badge.TranslatePoint(new Point(badge.Bounds.Width, badge.Bounds.Height), this.SchematicsContainer);

            if (!topLeft.HasValue || !bottomRight.HasValue)
            {
                continue;
            }

            var badgeRect = new Rect(topLeft.Value, bottomRight.Value);
            if (badgeRect.Contains(pointerInContainer))
            {
                entry = candidateEntry;
                color = candidateColor;
                return true;
            }
        }

        return false;
    }

    // ###########################################################################################
    // Repositions/rescales the existing "Show worklogs" badges for the current view matrix without
    // touching which entries are shown - called on every zoom/pan tick (TabSchematics.Viewport.cs).
    // The canvas's own MatrixTransform already keeps the colored-area overlay and the badges'
    // container in sync cheaply; this only needs to redo the inverse-scale ScaleTransform and
    // re-anchor each badge, the same per-frame work RefreshWorklogEntryBadge does for the single
    // in-progress draft badge. No disk read and no control-tree rebuild, unlike
    // RefreshWorklogEntriesListOverlay, which this deliberately does not call.
    // ###########################################################################################
    private void RescaleWorklogEntriesListBadges()
    {
        // Parked pills are deliberately NOT laid out here. This runs on every zoom/pan frame, and
        // the parked canvas carries no transform - by construction those pills do not move when the
        // view matrix changes, so a pass here would Measure every one of them and recompute the
        // same positions, per frame, for no visible difference.
        //
        // The two things that genuinely move them - the container resizing, and the "Netlist names"
        // panel opening or changing width - are covered by the Bounds/IsVisible subscriptions set
        // up in Initialize (see OnWorklogParkedBadgeLayoutTriggerChanged).
        if (this.thisWorklogEntriesListBadges.Count == 0 || this.currentFullResBitmap == null)
        {
            return;
        }

        var contentRect = this.GetImageContentRect();
        double scale = this.schematicsMatrix.M11;
        double inverseScale = scale > 0 ? 1.0 / scale : 1.0;

        foreach (var (badge, pixelRect, _, _) in this.thisWorklogEntriesListBadges)
        {
            this.PositionWorklogEntriesListBadge(badge, pixelRect, contentRect, inverseScale);
        }
    }

    // ###########################################################################################
    // Opens the full editor for the clicked saved entry's pill. Looks the entry back up from disk
    // by id (rather than capturing the WorklogEntryRecord in the closure) so the editor always
    // opens against the latest saved data, even if something else changed it first.
    //
    // Left button only, and never while an entry area is being marked out:
    //  - the badges sit on the one overlay canvas that is hit-test visible, so handling every
    //    button made each of them a dead zone for right-button panning (and opened the editor on a
    //    right-click).
    //  - during entry-drawing mode a drag that crossed a badge opened the editor over the
    //    half-drawn rectangle, leaving the drawing state stuck with no pointer release.
    // In both cases the press must fall through untouched rather than being marked handled.
    // ###########################################################################################
    private async void OnWorklogEntryPillPointerPressed(object? sender, PointerPressedEventArgs e)
    {
        if (this.thisIsWorklogEntryMode ||
            !e.GetCurrentPoint(this.SchematicsContainer).Properties.IsLeftButtonPressed)
        {
            return;
        }

        // Without this, the press bubbles up to OnSchematicsPointerPressed and is picked up as the
        // start of a manual KiCad-trace drag, which then follows the mouse for the rest of the
        // click - even once the editor dialog above has already opened and closed.
        e.Handled = true;

        if (sender is not Border { Tag: int entryId })
        {
            return;
        }

        int workbookId = this.thisWorklogEntriesListWorkbookId;

        var entry = WorklogManager.GetEntries(workbookId).FirstOrDefault(x => x.Id == entryId);
        if (entry == null)
        {
            return;
        }

        if (TopLevel.GetTopLevel(this) is not Window ownerWindow)
        {
            return;
        }

        var editor = new WorklogEntryEditorWindow();

        // The workbook id is captured up front: the dialog below is modal but the board can still
        // change underneath it (the thumbnail load path refreshes on its own), and re-reading the
        // field afterwards would point the editor at a different workbook than the badge came from.
        editor.Initialize(workbookId, entry, this.currentFullResBitmap);

        // The editor cannot work this out for itself - it has neither the board data nor the
        // highlight rectangles - so the scope is computed here and handed over. Resolved against
        // the entry's OWN schematic rather than whichever one is on screen: the "Show worklogs"
        // list can put a badge for another schematic in view, and using the visible one would
        // offer components from a different board image entirely.
        var componentsInScope = this.BuildWorklogEntryComponentScope(entry);
        if (componentsInScope != null)
        {
            editor.InitializeComponentScope(componentsInScope);
        }

        await editor.ShowDialog(ownerWindow);

        if (!editor.WasSaved)
        {
            return;
        }

        // Anything after the await runs as an async void continuation, where an exception is
        // unobservable and takes the process down instead of surfacing - so the refresh, which
        // touches controls a board switch may have torn down meanwhile, is guarded.
        try
        {
            this.RefreshWorklogEntriesListOverlay();
            this.MainWindow?.RefreshWorklogBar();
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to refresh worklog overlays after editing entry [#{entryId}]: [{ex.Message}]");
        }
    }
}

// ###########################################################################################
// One row in the worklog editor's "Mark components in scope" / "Mark components completed"
// checklists: a component whose highlight rectangle intersects the entry's area. Public and
// top-level (not nested inside a control) so the compiled DataTemplates in
// WorklogEntryEditorWindow.axaml can bind to it.
//
// It lives in this file rather than beside that window because the SCOPE is computed here - see
// BuildWorklogEntryComponentScope and BuildNewWorklogEntryComponentScope - and the editor only
// renders the rows it is handed.
// ###########################################################################################
public sealed class WorklogEntryComponentRow : System.ComponentModel.INotifyPropertyChanged
{
    private bool _isChecked = true;

    public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;

    public bool IsChecked
    {
        get => this._isChecked;
        set
        {
            if (this._isChecked == value)
                return;

            this._isChecked = value;
            this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.IsChecked)));
        }
    }

    public string BoardLabel { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
}
