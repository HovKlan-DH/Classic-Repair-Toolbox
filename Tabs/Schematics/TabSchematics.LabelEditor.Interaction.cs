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
// Pointer and keyboard interaction inside the label editor: selecting highlights, resize
// handles, drawing new rectangles, dragging, duplicating, and the pixel/local/container
// coordinate conversions those need.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private readonly HashSet<EditableComponentHighlight> thisSelectedLabelEditorHighlights = new();

    private EditableComponentHighlight? thisSelectedLabelEditorHighlight;

    private bool thisIsDrawingLabelEditorRectangle;

    private Point thisLabelEditorDrawStartPixelPoint;

    private Rect? thisLabelEditorDraftRectangle;

    private LabelEditorDragMode thisLabelEditorDragMode;

    private Point thisLabelEditorDragStartPixelPoint;

    private Rect thisLabelEditorOriginalSelectionBounds;

    private readonly Dictionary<EditableComponentHighlight, Rect> thisLabelEditorOriginalDragRectangles = new();

    // ###########################################################################################
    // Computes the editor overlay image content rect using the exact same mapping as the main
    // schematic overlays so pointer hit testing stays aligned after reloads and mode switches.
    // ###########################################################################################
    private Rect GetLabelEditorImageContentRect()
    {
        return this.GetSchematicsContentRect();
    }

    // ###########################################################################################
    // Converts a schematic container pointer position into bitmap pixel coordinates used by the
    // label editor, returning false if the pointer is outside the visible image content area.
    // ###########################################################################################
    private bool TryGetLabelEditorPixelPoint(Point pointerInContainer, out Point pixelPoint)
    {
        pixelPoint = default;

        if (!this.thisIsLabelEditorMode || this.currentFullResBitmap == null)
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

        var contentRect = this.GetLabelEditorImageContentRect();
        if (contentRect.Width <= 0 || contentRect.Height <= 0 || !contentRect.Contains(localPoint))
        {
            return false;
        }

        double px = ((localPoint.X - contentRect.X) / contentRect.Width) * this.currentFullResBitmap.PixelSize.Width;
        double py = ((localPoint.Y - contentRect.Y) / contentRect.Height) * this.currentFullResBitmap.PixelSize.Height;

        pixelPoint = new Point(px, py);
        return true;
    }

    // ###########################################################################################
    // Tries to find the topmost editable highlight rectangle under the current pointer position.
    // Returns the working-list index for direct select/delete operations.
    // ###########################################################################################
    private bool TryGetLabelEditorHighlightAtContainerPoint(Point pointerInContainer, out int workingIndex)
    {
        workingIndex = -1;

        if (!this.TryGetLabelEditorPixelPoint(pointerInContainer, out var pixelPoint))
        {
            return false;
        }

        string schematicName = this.GetCurrentSchematicName();
        if (string.IsNullOrWhiteSpace(schematicName))
        {
            return false;
        }

        for (int i = this.thisLabelEditorWorkingHighlights.Count - 1; i >= 0; i--)
        {
            var row = this.thisLabelEditorWorkingHighlights[i];
            if (!string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var rect = new Rect(row.X, row.Y, row.Width, row.Height);
            if (!rect.Contains(pixelPoint))
            {
                continue;
            }

            workingIndex = i;
            return true;
        }

        return false;
    }

    // ###########################################################################################
    // Clears the current label-editor selection and removes visible selection markers.
    // ###########################################################################################
    private void ClearSelectedLabelEditorHighlight()
    {
        this.ClearSelectedLabelEditorHighlights();
    }

    // ###########################################################################################
    // Deletes the requested working-copy editor highlight and refreshes the overlay immediately.
    // Records the previous state so the deletion can be undone within the current editor session.
    // ###########################################################################################
    private void DeleteLabelEditorHighlight(int workingIndex)
    {
        if (workingIndex < 0 || workingIndex >= this.thisLabelEditorWorkingHighlights.Count)
        {
            return;
        }

        this.PushLabelEditorUndoState(this.CreateLabelEditorUndoState());

        var deleted = this.thisLabelEditorWorkingHighlights[workingIndex];
        this.thisLabelEditorWorkingHighlights.RemoveAt(workingIndex);
        this.thisSelectedLabelEditorHighlights.Remove(deleted);
        this.thisLabelEditorOriginalDragRectangles.Remove(deleted);

        if (ReferenceEquals(this.thisSelectedLabelEditorHighlight, deleted))
        {
            this.thisSelectedLabelEditorHighlight = this.GetFirstSelectedLabelEditorHighlightForCurrentSchematic();
        }

        this.RefreshLabelEditorOverlay();

        Logger.Debug($"Label editor rectangle deleted for board label [{deleted.BoardLabel}] on schematic [{deleted.SchematicName}]");
    }

    // ###########################################################################################
    // Returns true when the pointer is currently inside the new-label prompt bounds.
    // ###########################################################################################
    private bool IsPointerInsideNewLabelPrompt(Point containerPoint)
    {
        if (!this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return false;
        }

        Point? translatedTopLeft = this.SchematicsNewLabelPromptBorder.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
        if (!translatedTopLeft.HasValue)
        {
            return false;
        }

        var promptRect = new Rect(translatedTopLeft.Value, this.SchematicsNewLabelPromptBorder.Bounds.Size);
        return promptRect.Contains(containerPoint);
    }

    // ###########################################################################################
    // Starts drawing a new editor rectangle from the current bitmap pixel position.
    // ###########################################################################################
    private void StartDrawingLabelEditorRectangle(Point startPixelPoint)
    {
        this.thisIsDrawingLabelEditorRectangle = true;
        this.thisLabelEditorDrawStartPixelPoint = startPixelPoint;
        this.thisLabelEditorDraftRectangle = new Rect(startPixelPoint.X, startPixelPoint.Y, 0, 0);
        this.RefreshLabelEditorOverlay();
    }

    // ###########################################################################################
    // Updates the current draft editor rectangle while the mouse is being dragged.
    // Applies the same neighbor-edge snap behavior used by resize operations unless Shift is held.
    // ###########################################################################################
    private void UpdateDrawingLabelEditorRectangle(Point currentPixelPoint, KeyModifiers modifiers)
    {
        if (!this.thisIsDrawingLabelEditorRectangle)
        {
            return;
        }

        var draftRect = RectGeometry.CreateNormalizedRect(this.thisLabelEditorDrawStartPixelPoint, currentPixelPoint);
        var snapGuides = new List<(Point Start, Point End)>();

        this.ApplyNewLabelEditorRectangleSnap(
            ref draftRect,
            snapGuides,
            modifiers.HasFlag(KeyModifiers.Shift));

        this.thisLabelEditorDraftRectangle = draftRect;
        this.RefreshLabelEditorOverlay(snapGuides);
    }

    // ###########################################################################################
    // Completes the current rectangle drawing operation and opens the board-label prompt.
    // Records the pre-create state so the new rectangle can be undone after confirmation.
    // The final rectangle also uses neighbor-edge snap behavior unless Shift is held.
    // ###########################################################################################
    private void CompleteDrawingLabelEditorRectangle(
        Point releaseContainerPoint,
        Point releasePixelPoint,
        KeyModifiers modifiers)
    {
        if (!this.thisIsDrawingLabelEditorRectangle)
        {
            return;
        }

        var finalRect = RectGeometry.CreateNormalizedRect(this.thisLabelEditorDrawStartPixelPoint, releasePixelPoint);

        var snapGuides = new List<(Point Start, Point End)>();
        this.ApplyNewLabelEditorRectangleSnap(
            ref finalRect,
            snapGuides,
            modifiers.HasFlag(KeyModifiers.Shift));

        this.thisIsDrawingLabelEditorRectangle = false;
        this.thisLabelEditorDraftRectangle = null;

        if (LabelEditorGeometry.IsLabelEditorRectangleTooSmall(finalRect))
        {
            this.RefreshLabelEditorOverlay();
            return;
        }

        this.PushLabelEditorUndoState(this.CreateLabelEditorUndoState());

        var newRow = new EditableComponentHighlight
        {
            SchematicName = this.GetCurrentSchematicName(),
            BoardLabel = string.Empty,
            Category = string.Empty,
            X = finalRect.X,
            Y = finalRect.Y,
            Width = finalRect.Width,
            Height = finalRect.Height
        };

        this.thisLabelEditorWorkingHighlights.Add(newRow);
        this.SetSingleSelectedLabelEditorHighlight(newRow, refresh: false);
        this.thisPendingNewLabelEditorHighlight = newRow;

        this.RefreshLabelEditorOverlay();
        this.ShowNewLabelEditorPrompt(releaseContainerPoint);
    }

    // ###########################################################################################
    // Converts a schematic container pointer position into overlay-local coordinates used for
    // editor-handle hit testing, returning false if the pointer is outside the image content area.
    // ###########################################################################################
    private bool TryGetLabelEditorLocalPoint(Point pointerInContainer, out Point localPoint)
    {
        localPoint = default;

        if (!this.thisIsLabelEditorMode || this.currentFullResBitmap == null)
        {
            return false;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            return false;
        }

        localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        var contentRect = this.GetLabelEditorImageContentRect();
        return contentRect.Width > 0 && contentRect.Height > 0 && contentRect.Contains(localPoint);
    }

    // ###########################################################################################
    // Converts a pixel-space highlight rectangle into editor overlay local coordinates.
    // ###########################################################################################
    private Rect ConvertLabelEditorPixelRectToLocalRect(Rect pixelRect)
    {
        if (this.currentFullResBitmap == null)
        {
            return default;
        }

        var contentRect = this.GetLabelEditorImageContentRect();

        double sx = contentRect.Width / this.currentFullResBitmap.PixelSize.Width;
        double sy = contentRect.Height / this.currentFullResBitmap.PixelSize.Height;

        double x = contentRect.X + (pixelRect.X * sx);
        double y = contentRect.Y + (pixelRect.Y * sy);
        double w = pixelRect.Width * sx;
        double h = pixelRect.Height * sy;

        return new Rect(x, y, w, h);
    }

    // ###########################################################################################
    // Tries to hit one of the resize handles of any selected rectangle under the pointer.
    // Corner handles are evaluated first, and side handles only exist in the center gap between
    // corners so tiny components still allow true corner resizing in both X and Y.
    // ###########################################################################################
    private bool TryGetSelectedLabelEditorHandleAtContainerPoint(
        Point pointerInContainer,
        out int workingIndex,
        out LabelEditorDragMode dragMode)
    {
        workingIndex = -1;
        dragMode = LabelEditorDragMode.None;

        if (this.currentFullResBitmap == null)
        {
            return false;
        }

        if (!this.TryGetLabelEditorLocalPoint(pointerInContainer, out var localPoint))
        {
            return false;
        }

        double scale = Math.Max(0.0001, this.schematicsMatrix.M11);
        string schematicName = this.GetCurrentSchematicName();

        for (int i = this.thisLabelEditorWorkingHighlights.Count - 1; i >= 0; i--)
        {
            var row = this.thisLabelEditorWorkingHighlights[i];

            if (!string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase) ||
                !this.IsSelectedLabelEditorHighlight(row))
            {
                continue;
            }

            var localRect = this.ConvertLabelEditorPixelRectToLocalRect(new Rect(row.X, row.Y, row.Width, row.Height));

            foreach (var hitTarget in LabelEditorGeometry.BuildLabelEditorHandleHitRects(localRect, scale))
            {
                if (!hitTarget.HitRect.Contains(localPoint))
                {
                    continue;
                }

                workingIndex = i;
                dragMode = hitTarget.DragMode;
                return true;
            }
        }

        return false;
    }

    // ###########################################################################################
    // Starts dragging an existing highlight rectangle for move or resize operations.
    // ###########################################################################################
    private void StartLabelEditorDrag(int workingIndex, Point startPixelPoint, LabelEditorDragMode dragMode)
    {
        if (workingIndex < 0 || workingIndex >= this.thisLabelEditorWorkingHighlights.Count)
        {
            return;
        }

        var anchorHighlight = this.thisLabelEditorWorkingHighlights[workingIndex];

        if (!this.IsSelectedLabelEditorHighlight(anchorHighlight))
        {
            this.SetSingleSelectedLabelEditorHighlight(anchorHighlight, refresh: false);
        }
        else
        {
            this.thisSelectedLabelEditorHighlight = anchorHighlight;
        }

        this.thisLabelEditorDragMode = dragMode;
        this.thisLabelEditorDragStartPixelPoint = startPixelPoint;
        this.thisLabelEditorOriginalSelectionBounds = new Rect(
            anchorHighlight.X,
            anchorHighlight.Y,
            anchorHighlight.Width,
            anchorHighlight.Height);

        this.CaptureSelectedLabelEditorDragState();
        this.RefreshLabelEditorOverlay();
    }

    // ###########################################################################################
    // Applies the current drag delta to the selected rectangle for move or resize operations.
    // Holding Shift during mouse drag suppresses snap alignment.
    // Uses the drag-start rectangles as the stable source so pointer movement does not compound.
    // ###########################################################################################
    private void UpdateLabelEditorDrag(Point currentPixelPoint, KeyModifiers modifiers)
    {
        if (!this.HasSelectedLabelEditorHighlightsForCurrentSchematic() ||
            this.thisLabelEditorDragMode == LabelEditorDragMode.None)
        {
            return;
        }

        var selectedRows = this.GetSelectedLabelEditorHighlightsForCurrentSchematic();
        if (selectedRows.Count == 0)
        {
            return;
        }

        var sourceRects = new Dictionary<EditableComponentHighlight, Rect>();

        foreach (var row in selectedRows)
        {
            if (!this.thisLabelEditorOriginalDragRectangles.TryGetValue(row, out var originalRect))
            {
                originalRect = new Rect(row.X, row.Y, row.Width, row.Height);
            }

            sourceRects[row] = originalRect;
        }

        double dx = currentPixelPoint.X - this.thisLabelEditorDragStartPixelPoint.X;
        double dy = currentPixelPoint.Y - this.thisLabelEditorDragStartPixelPoint.Y;
        bool suppressSnap = modifiers.HasFlag(KeyModifiers.Shift);

        var snapGuides = new List<(Point Start, Point End)>();

        if (this.thisLabelEditorDragMode == LabelEditorDragMode.Move)
        {
            double originalLeft = sourceRects.Values.Min(rect => rect.Left);
            double originalTop = sourceRects.Values.Min(rect => rect.Top);
            double originalRight = sourceRects.Values.Max(rect => rect.Right);
            double originalBottom = sourceRects.Values.Max(rect => rect.Bottom);

            var originalSelectionBounds = new Rect(
                originalLeft,
                originalTop,
                originalRight - originalLeft,
                originalBottom - originalTop);

            var movedSelectionBounds = new Rect(
                originalSelectionBounds.X + dx,
                originalSelectionBounds.Y + dy,
                originalSelectionBounds.Width,
                originalSelectionBounds.Height);

            this.ApplyLabelEditorMoveSnap(
                selectedRows,
                sourceRects,
                ref movedSelectionBounds,
                snapGuides,
                suppressSnap);

            double snappedDx = movedSelectionBounds.X - originalSelectionBounds.X;
            double snappedDy = movedSelectionBounds.Y - originalSelectionBounds.Y;

            foreach (var row in selectedRows)
            {
                var originalRect = sourceRects[row];
                row.X = originalRect.X + snappedDx;
                row.Y = originalRect.Y + snappedDy;
                row.Width = originalRect.Width;
                row.Height = originalRect.Height;
            }

            this.RefreshLabelEditorOverlay(snapGuides);
            return;
        }

        foreach (var row in selectedRows)
        {
            var originalRect = sourceRects[row];

            double left = originalRect.Left;
            double top = originalRect.Top;
            double right = originalRect.Right;
            double bottom = originalRect.Bottom;

            switch (this.thisLabelEditorDragMode)
            {
                case LabelEditorDragMode.ResizeTopLeft:
                    left += dx;
                    top += dy;
                    break;

                case LabelEditorDragMode.ResizeTop:
                    top += dy;
                    break;

                case LabelEditorDragMode.ResizeTopRight:
                    right += dx;
                    top += dy;
                    break;

                case LabelEditorDragMode.ResizeRight:
                    right += dx;
                    break;

                case LabelEditorDragMode.ResizeBottomRight:
                    right += dx;
                    bottom += dy;
                    break;

                case LabelEditorDragMode.ResizeBottom:
                    bottom += dy;
                    break;

                case LabelEditorDragMode.ResizeBottomLeft:
                    left += dx;
                    bottom += dy;
                    break;

                case LabelEditorDragMode.ResizeLeft:
                    left += dx;
                    break;
            }

            this.ApplyLabelEditorResizeSnap(row, ref left, ref top, ref right, ref bottom, snapGuides, suppressSnap);

            const double minimumSize = 1.0;

            if (right < left + minimumSize)
            {
                if (this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeLeft ||
                    this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeTopLeft ||
                    this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeBottomLeft)
                {
                    left = right - minimumSize;
                }
                else
                {
                    right = left + minimumSize;
                }
            }

            if (bottom < top + minimumSize)
            {
                if (this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeTop ||
                    this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeTopLeft ||
                    this.thisLabelEditorDragMode == LabelEditorDragMode.ResizeTopRight)
                {
                    top = bottom - minimumSize;
                }
                else
                {
                    bottom = top + minimumSize;
                }
            }

            row.X = left;
            row.Y = top;
            row.Width = Math.Max(1.0, right - left);
            row.Height = Math.Max(1.0, bottom - top);
        }

        this.RefreshLabelEditorOverlay(snapGuides);
    }

    // ###########################################################################################
    // Finishes the current move or resize operation for the selected rectangle.
    // Clears any temporary snap guides and records the pre-drag state for undo when needed.
    // ###########################################################################################
    private void CompleteLabelEditorDrag()
    {
        if (this.thisLabelEditorDragMode != LabelEditorDragMode.None)
        {
            var beforeDragState = this.CreateLabelEditorUndoStateFromOriginalDragState();
            var afterDragState = this.CreateLabelEditorUndoState();

            if (!AreLabelEditorUndoStatesEqual(beforeDragState, afterDragState))
            {
                this.PushLabelEditorUndoState(beforeDragState);
            }
        }

        this.thisLabelEditorDragMode = LabelEditorDragMode.None;
        this.thisLabelEditorOriginalDragRectangles.Clear();

        if (this.SchematicsLabelEditorOverlay.SnapGuides.Count > 0)
        {
            this.SetLabelEditorOverlayTransientState(
                snapGuides: Array.Empty<(Point Start, Point End)>());
        }
    }

    // ###########################################################################################
    // Applies keyboard move, expand, or shrink operations to the selected editor rectangle.
    // Arrow keys move by 1 px, Shift expands in the pressed direction, and Alt shrinks from
    // the opposite side of the pressed direction. Each committed step is undoable.
    // Keyboard operations do not snap, but exact neighbor matches show the dashed guide.
    // ###########################################################################################
    private bool ApplySelectedLabelEditorKeyboardStep(Key key, KeyModifiers modifiers)
    {
        if (!this.thisIsLabelEditorMode ||
            !this.HasSelectedLabelEditorHighlightsForCurrentSchematic() ||
            this.SchematicsNewLabelPromptBorder.IsVisible)
        {
            return false;
        }

        if (modifiers.HasFlag(KeyModifiers.Shift) && modifiers.HasFlag(KeyModifiers.Alt))
        {
            return false;
        }

        var selectedRows = this.GetSelectedLabelEditorHighlightsForCurrentSchematic();
        if (selectedRows.Count == 0)
        {
            return false;
        }

        var undoState = this.CreateLabelEditorUndoState();

        var sourceRects = selectedRows.ToDictionary(
            row => row,
            row => new Rect(row.X, row.Y, row.Width, row.Height));

        bool isShift = modifiers.HasFlag(KeyModifiers.Shift);
        bool isAlt = modifiers.HasFlag(KeyModifiers.Alt);
        const double step = 1.0;
        bool changed = false;

        foreach (var row in selectedRows)
        {
            var originalRect = sourceRects[row];

            double x = originalRect.X;
            double y = originalRect.Y;
            double width = originalRect.Width;
            double height = originalRect.Height;

            if (!isShift && !isAlt)
            {
                switch (key)
                {
                    case Key.Left:
                        x -= step;
                        changed = true;
                        break;

                    case Key.Right:
                        x += step;
                        changed = true;
                        break;

                    case Key.Up:
                        y -= step;
                        changed = true;
                        break;

                    case Key.Down:
                        y += step;
                        changed = true;
                        break;
                }
            }
            else if (isShift)
            {
                switch (key)
                {
                    case Key.Left:
                        x -= step;
                        width += step;
                        changed = true;
                        break;

                    case Key.Right:
                        width += step;
                        changed = true;
                        break;

                    case Key.Up:
                        y -= step;
                        height += step;
                        changed = true;
                        break;

                    case Key.Down:
                        height += step;
                        changed = true;
                        break;
                }
            }
            else if (isAlt)
            {
                switch (key)
                {
                    case Key.Left:
                        if (width > step)
                        {
                            width -= step;
                            changed = true;
                        }
                        break;

                    case Key.Right:
                        if (width > step)
                        {
                            x += step;
                            width -= step;
                            changed = true;
                        }
                        break;

                    case Key.Up:
                        if (height > step)
                        {
                            height -= step;
                            changed = true;
                        }
                        break;

                    case Key.Down:
                        if (height > step)
                        {
                            y += step;
                            height -= step;
                            changed = true;
                        }
                        break;
                }
            }

            row.X = x;
            row.Y = y;
            row.Width = Math.Max(1.0, width);
            row.Height = Math.Max(1.0, height);
        }

        if (!changed)
        {
            return false;
        }

        var snapGuides = new List<(Point Start, Point End)>();

        if (!isShift && !isAlt)
        {
            double movedLeft = selectedRows.Min(row => row.X);
            double movedTop = selectedRows.Min(row => row.Y);
            double movedRight = selectedRows.Max(row => row.X + row.Width);
            double movedBottom = selectedRows.Max(row => row.Y + row.Height);

            var movedSelectionBounds = new Rect(
                movedLeft,
                movedTop,
                movedRight - movedLeft,
                movedBottom - movedTop);

            this.ApplyLabelEditorMoveSnap(
                selectedRows,
                sourceRects,
                ref movedSelectionBounds,
                snapGuides,
                suppressSnap: false,
                snapOnMatch: false);
        }
        else if (LabelEditorGeometry.TryGetKeyboardLabelEditorResizeDragMode(key, modifiers, out var keyboardResizeDragMode))
        {
            foreach (var row in selectedRows)
            {
                double left = row.X;
                double top = row.Y;
                double right = row.X + row.Width;
                double bottom = row.Y + row.Height;

                this.ApplyLabelEditorResizeSnap(
                    row,
                    ref left,
                    ref top,
                    ref right,
                    ref bottom,
                    snapGuides,
                    suppressSnap: false,
                    snapOnMatch: false,
                    dragModeOverride: keyboardResizeDragMode);
            }
        }

        this.PushLabelEditorUndoState(undoState);
        this.RefreshLabelEditorOverlay(snapGuides);
        return true;
    }

    // ###########################################################################################
    // Updates the schematic cursor for label-editor interactions.
    // Shows Hand over resize handles and SizeAll over movable rectangles.
    // ###########################################################################################
    private void UpdateLabelEditorCursor(Point pointerInContainer)
    {
        if (!this.thisIsLabelEditorMode)
        {
            this.SchematicsContainer.Cursor = Cursor.Default;
            return;
        }

        if (this.thisLabelEditorDragMode != LabelEditorDragMode.None)
        {
            this.SchematicsContainer.Cursor = this.thisLabelEditorDragMode == LabelEditorDragMode.Move
                ? new Cursor(StandardCursorType.SizeAll)
                : new Cursor(StandardCursorType.Hand);
            return;
        }

        if (this.thisIsDrawingLabelEditorRectangle)
        {
            this.SchematicsContainer.Cursor = Cursor.Default;
            return;
        }

        if (this.TryGetSelectedLabelEditorHandleAtContainerPoint(pointerInContainer, out _))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.Hand);
            return;
        }

        if (this.TryGetSelectedLabelEditorHighlightAtContainerPoint(pointerInContainer, out _))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);
            return;
        }

        if (this.TryGetLabelEditorHighlightAtContainerPoint(pointerInContainer, out _))
        {
            this.SchematicsContainer.Cursor = new Cursor(StandardCursorType.SizeAll);
            return;
        }

        this.SchematicsContainer.Cursor = Cursor.Default;
    }

    // ###########################################################################################
    // Returns the currently selected editor highlights for the active schematic in working-list order.
    // ###########################################################################################
    private List<EditableComponentHighlight> GetSelectedLabelEditorHighlightsForCurrentSchematic()
    {
        string schematicName = this.GetCurrentSchematicName();

        return this.thisLabelEditorWorkingHighlights
            .Where(row => string.Equals(row.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
            .Where(row => this.thisSelectedLabelEditorHighlights.Contains(row))
            .ToList();
    }

    // ###########################################################################################
    // Returns the first selected editor highlight for the active schematic, or null when none exist.
    // ###########################################################################################
    private EditableComponentHighlight? GetFirstSelectedLabelEditorHighlightForCurrentSchematic()
    {
        return this.GetSelectedLabelEditorHighlightsForCurrentSchematic().FirstOrDefault();
    }

    // ###########################################################################################
    // Returns true when the given highlight is part of the current editor selection.
    // ###########################################################################################
    private bool IsSelectedLabelEditorHighlight(EditableComponentHighlight highlight)
    {
        return this.thisSelectedLabelEditorHighlights.Contains(highlight);
    }

    // ###########################################################################################
    // Clears the current multi-selection and optionally refreshes the editor overlay.
    // ###########################################################################################
    private void ClearSelectedLabelEditorHighlights(bool refresh = true)
    {
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlight = null;

        if (refresh)
        {
            this.RefreshLabelEditorOverlay();
        }
    }

    // ###########################################################################################
    // Replaces the current multi-selection with one highlight and sets it as the primary selection.
    // ###########################################################################################
    private void SetSingleSelectedLabelEditorHighlight(EditableComponentHighlight highlight, bool refresh = true)
    {
        this.thisSelectedLabelEditorHighlights.Clear();
        this.thisSelectedLabelEditorHighlights.Add(highlight);
        this.thisSelectedLabelEditorHighlight = highlight;

        if (refresh)
        {
            this.RefreshLabelEditorOverlay();
        }
    }

    // ###########################################################################################
    // Toggles one highlight inside the current multi-selection and updates the primary selection.
    // ###########################################################################################
    private void ToggleSelectedLabelEditorHighlight(EditableComponentHighlight highlight)
    {
        if (this.thisSelectedLabelEditorHighlights.Contains(highlight))
        {
            this.thisSelectedLabelEditorHighlights.Remove(highlight);

            if (ReferenceEquals(this.thisSelectedLabelEditorHighlight, highlight))
            {
                this.thisSelectedLabelEditorHighlight = this.GetFirstSelectedLabelEditorHighlightForCurrentSchematic();
            }
        }
        else
        {
            this.thisSelectedLabelEditorHighlights.Add(highlight);
            this.thisSelectedLabelEditorHighlight = highlight;
        }

        this.RefreshLabelEditorOverlay();
    }

    // ###########################################################################################
    // Returns true when there is at least one selected editor highlight on the current schematic.
    // ###########################################################################################
    private bool HasSelectedLabelEditorHighlightsForCurrentSchematic()
    {
        return this.GetSelectedLabelEditorHighlightsForCurrentSchematic().Count > 0;
    }

    // ###########################################################################################
    // Computes the combined selection bounds for all selected editor highlights on the current schematic.
    // ###########################################################################################
    private bool TryGetSelectedLabelEditorBounds(out Rect selectionBounds)
    {
        selectionBounds = default;

        var selected = this.GetSelectedLabelEditorHighlightsForCurrentSchematic();
        if (selected.Count == 0)
        {
            return false;
        }

        double left = selected.Min(row => row.X);
        double top = selected.Min(row => row.Y);
        double right = selected.Max(row => row.X + row.Width);
        double bottom = selected.Max(row => row.Y + row.Height);

        selectionBounds = new Rect(left, top, right - left, bottom - top);
        return true;
    }

    // ###########################################################################################
    // Captures the original rectangles of all selected highlights before a move or resize starts.
    // ###########################################################################################
    private void CaptureSelectedLabelEditorDragState()
    {
        this.thisLabelEditorOriginalDragRectangles.Clear();

        foreach (var row in this.GetSelectedLabelEditorHighlightsForCurrentSchematic())
        {
            this.thisLabelEditorOriginalDragRectangles[row] = new Rect(row.X, row.Y, row.Width, row.Height);
        }
    }

    // ###########################################################################################
    // Returns the selected editor highlight under the pointer, if any.
    // ###########################################################################################
    private bool TryGetSelectedLabelEditorHighlightAtContainerPoint(Point pointerInContainer, out int workingIndex)
    {
        workingIndex = -1;

        if (!this.TryGetLabelEditorHighlightAtContainerPoint(pointerInContainer, out var hitIndex))
        {
            return false;
        }

        var hitHighlight = this.thisLabelEditorWorkingHighlights[hitIndex];
        if (!this.IsSelectedLabelEditorHighlight(hitHighlight))
        {
            return false;
        }

        workingIndex = hitIndex;
        return true;
    }

    // ###########################################################################################
    // Compatibility overload used by cursor and hover logic.
    // ###########################################################################################
    private bool TryGetSelectedLabelEditorHandleAtContainerPoint(Point pointerInContainer, out LabelEditorDragMode dragMode)
    {
        return this.TryGetSelectedLabelEditorHandleAtContainerPoint(pointerInContainer, out _, out dragMode);
    }

    // ###########################################################################################
    // Converts one label-editor bitmap pixel point into schematics container coordinates so popups
    // can be positioned near duplicated or newly created rectangles.
    // ###########################################################################################
    private Point ConvertLabelEditorPixelPointToContainerPoint(Point pixelPoint)
    {
        if (this.currentFullResBitmap == null ||
            this.currentFullResBitmap.PixelSize.Width <= 0 ||
            this.currentFullResBitmap.PixelSize.Height <= 0)
        {
            return new Point(0, 0);
        }

        var contentRect = this.GetLabelEditorImageContentRect();

        double localX = contentRect.X + ((pixelPoint.X / this.currentFullResBitmap.PixelSize.Width) * contentRect.Width);
        double localY = contentRect.Y + ((pixelPoint.Y / this.currentFullResBitmap.PixelSize.Height) * contentRect.Height);

        return new Point(
            (localX * this.schematicsMatrix.M11) + (localY * this.schematicsMatrix.M21) + this.schematicsMatrix.M31,
            (localX * this.schematicsMatrix.M12) + (localY * this.schematicsMatrix.M22) + this.schematicsMatrix.M32);
    }

    // ###########################################################################################
    // Duplicates the currently selected label-editor rectangle, places the copy next to the source,
    // and opens the new-label prompt for the duplicated component.
    // The duplicated component prefers the source category and makes it the new default category.
    // ###########################################################################################
    private bool TryDuplicateSelectedLabelEditorHighlight()
    {
        if (!this.thisIsLabelEditorMode ||
            this.SchematicsNewLabelPromptBorder.IsVisible ||
            this.thisIsDrawingLabelEditorRectangle ||
            this.thisLabelEditorDragMode != LabelEditorDragMode.None ||
            this.currentFullResBitmap == null)
        {
            return false;
        }

        var selectedHighlights = this.GetSelectedLabelEditorHighlightsForCurrentSchematic();
        if (selectedHighlights.Count != 1)
        {
            return false;
        }

        var sourceHighlight = selectedHighlights[0];
        string duplicatedCategory = sourceHighlight.Category?.Trim() ?? string.Empty;

        const double duplicateGapPixels = 12.0;

        double bitmapWidth = this.currentFullResBitmap.PixelSize.Width;
        double bitmapHeight = this.currentFullResBitmap.PixelSize.Height;

        double duplicateX = sourceHighlight.X + sourceHighlight.Width + duplicateGapPixels;
        double duplicateY = sourceHighlight.Y;

        if (duplicateX + sourceHighlight.Width > bitmapWidth)
        {
            duplicateX = sourceHighlight.X - sourceHighlight.Width - duplicateGapPixels;
        }

        duplicateX = Math.Clamp(
            duplicateX,
            0.0,
            Math.Max(0.0, bitmapWidth - sourceHighlight.Width));

        duplicateY = Math.Clamp(
            duplicateY,
            0.0,
            Math.Max(0.0, bitmapHeight - sourceHighlight.Height));

        this.PushLabelEditorUndoState(this.CreateLabelEditorUndoState());

        var duplicatedHighlight = new EditableComponentHighlight
        {
            SchematicName = sourceHighlight.SchematicName,
            BoardLabel = string.Empty,
            Category = duplicatedCategory,
            X = duplicateX,
            Y = duplicateY,
            Width = sourceHighlight.Width,
            Height = sourceHighlight.Height
        };

        if (!string.IsNullOrWhiteSpace(duplicatedCategory))
        {
            this.thisLastCreatedLabelEditorCategory = duplicatedCategory;
        }

        this.thisLabelEditorWorkingHighlights.Add(duplicatedHighlight);
        this.SetSingleSelectedLabelEditorHighlight(duplicatedHighlight, refresh: false);
        this.thisPendingNewLabelEditorHighlight = duplicatedHighlight;

        this.RefreshLabelEditorOverlay();

        Point promptAnchorPoint = this.ConvertLabelEditorPixelPointToContainerPoint(
            new Point(
                duplicatedHighlight.X + (duplicatedHighlight.Width / 2.0),
                duplicatedHighlight.Y + (duplicatedHighlight.Height / 2.0)));

        this.ShowNewLabelEditorPrompt(promptAnchorPoint);

        Logger.Info(
            $"Label editor duplicated rectangle from board label [{sourceHighlight.BoardLabel}] on schematic [{sourceHighlight.SchematicName}]");

        return true;
    }
}