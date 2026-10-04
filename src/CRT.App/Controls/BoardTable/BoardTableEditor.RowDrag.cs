using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // BoardTableEditor - DRAGGING A ROW, WITH A PLACEHOLDER (owner request, 2026-09-24: "insert
    // a placeholder when I drag a row up and down, like when moving an image in the worklog").
    //
    // It works the way the worklog's photo and file lists do (WorklogEntryEditorWindow.Attachments):
    // pressing the row's grip at the far left ARMS a drag, moving a few pixels STARTS it, and from
    // then on the dragged row moves LIVE into the place of the row under the pointer - drawn as an
    // empty red dashed slot, so what is seen is the space it will take and the rows around it
    // already sit in the order a drop produces. Releasing leaves it there.
    //
    // *** THIS REPLACES ProDataGrid's OWN ROW DRAG. *** That draws only a thin line between two rows
    // and moves the row once, on the drop, through a drop handler - there is no way to show a
    // travelling slot with it. Every rule of the move (never onto a red deleted row, one undo step
    // per drag, a drag back to the start leaving no trace) is in CRT.Data's BoardTableRowDrag;
    // this part only turns pointer positions into "the row under the pointer" or "past the edge".
    //
    // THE POINTER IS CAPTURED BY THE GRID, not by the row header that was pressed: moving rows
    // hands their containers to other rows, so the header under the press can end up showing a
    // different row - or none - mid-drag.
    //
    // *** THE TARGET IS READ OFF A FROZEN LAYOUT, NEVER THE LIVE ONE (reported flicker,
    // 2026-09-24). *** The first version asked the grid "which row is under the pointer" and then
    // scrolled the moved row into view. Over the half-visible row at the bottom that scroll slid a
    // new row under a pointer that had not moved, which was moved onto, which scrolled again - the
    // row ran away or flickered between two places. Exactly the feedback the worklog's frozen
    // boundaries exist to stop (WorklogEntryEditorWindow.CapturePhotoRowBoundaries), so it is done
    // the same way: the row SLOTS on screen are captured when the drag starts, and the pointer's
    // slot decides the target. Nothing scrolls except a step past the top or bottom edge, and the
    // slots are captured again after that scroll (and after a mouse-wheel scroll).
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // As the worklog's PhotoDragThreshold: a click with a shaky hand is not a drag.
        private const double RowDragThreshold = 4.0;

        private const string DropPlaceholderClass = "BoardTableDropPlaceholder";

        // On the grid while rows can be moved - it is what shows the grips (see the markup).
        private const string RowsDraggableClass = "RowsDraggable";

        // Pressed on a grip, not yet moved far enough to count as a drag.
        private BoardTableRow? thisArmedDragRow;
        private Point thisDragStartPoint;

        // Set once the drag has really started.
        private BoardTableRowDrag? thisRowDrag;

        // Moving a row re-lays out the grid, which raises pointer events synchronously that would
        // move it again - the worklog's reason for the same guard.
        private bool thisIsApplyingRowDragMove;

        // The frozen layout: each row slot on screen (top and bottom, in the grid's coordinates),
        // and the index in the sheet's rows that the first slot showed when it was captured. With
        // every row the same height, slot i shows row (first + i) for as long as nothing scrolls.
        private readonly List<(double Top, double Bottom)> thisRowSlots = new();
        private int thisFirstSlotIndex = -1;

        internal BoardTableRow? DraggedRowForTests => this.thisRowDrag?.Row;

        private void WireRowDragging()
        {
            // Tunnel, and handled ones too: the grid's own pointer handling (selection) must not see
            // a press that starts a drag, and must not stop this part seeing the moves.
            this.TableGrid.AddHandler(PointerPressedEvent, this.OnRowDragPointerPressed, RoutingStrategies.Tunnel, handledEventsToo: true);
            this.TableGrid.AddHandler(PointerMovedEvent, this.OnRowDragPointerMoved, RoutingStrategies.Tunnel, handledEventsToo: true);
            this.TableGrid.AddHandler(PointerReleasedEvent, this.OnRowDragPointerReleased, RoutingStrategies.Tunnel, handledEventsToo: true);

            // A mouse-wheel scroll mid-drag moves the rows under the frozen slots: capture them again
            // once the grid has laid itself out at the new position.
            this.TableGrid.AddHandler(
                PointerWheelChangedEvent,
                (_, _) =>
                {
                    if (this.thisRowDrag is not null)
                    {
                        Dispatcher.UIThread.Post(this.CaptureRowSlots, DispatcherPriority.Background);
                    }
                },
                RoutingStrategies.Tunnel,
                handledEventsToo: true);

            // Capture taken away (the window losing focus, say): the drag ends where it is.
            this.TableGrid.PointerCaptureLost += (_, _) => this.EndRowDrag();

            this.UpdateRowsDraggable();
        }

        // Rows can be moved - and show their grips - except while a picked pill in the colour key
        // or the search box hides some (BoardTableRowFilter, BoardTableSearch): a place among rows
        // that cannot be seen means nothing. Never in a read-only table (IsReadOnly).
        private void UpdateRowsDraggable() =>
            this.TableGrid.Classes.Set(BoardTableEditor.RowsDraggableClass, !this.IsNarrowed && !this.thisIsReadOnly);

        private void OnRowDragPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (this.IsNarrowed ||
                this.thisIsReadOnly ||
                this.thisCurrentSheet is null ||
                !e.GetCurrentPoint(this.TableGrid).Properties.IsLeftButtonPressed ||
                e.Source is not Visual source)
            {
                return;
            }

            DataGridRowHeader? header = source as DataGridRowHeader ?? source.FindAncestorOfType<DataGridRowHeader>();
            if (header?.DataContext is not BoardTableRow { IsDeleted: false } row || !this.thisCurrentSheet.Rows.Contains(row))
            {
                return;
            }

            e.Handled = true;

            this.TableGrid.CommitEdit();
            this.SelectCell(row, this.CurrentCell?.ColumnIndex ?? 0);

            this.thisArmedDragRow = row;
            this.thisDragStartPoint = e.GetPosition(this.TableGrid);
            e.Pointer.Capture(this.TableGrid);
        }

        private void OnRowDragPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.thisArmedDragRow is not { } row || this.thisCurrentSheet is null)
            {
                return;
            }

            // Released somewhere the release never reached (outside the window): without this the
            // next move would carry on a drag the contributor had ended.
            if (!e.GetCurrentPoint(this.TableGrid).Properties.IsLeftButtonPressed)
            {
                this.EndRowDrag();
                return;
            }

            Point current = e.GetPosition(this.TableGrid);

            if (this.thisRowDrag is null)
            {
                if (Math.Abs(current.Y - this.thisDragStartPoint.Y) < BoardTableEditor.RowDragThreshold &&
                    Math.Abs(current.X - this.thisDragStartPoint.X) < BoardTableEditor.RowDragThreshold)
                {
                    return;
                }

                this.thisRowDrag = this.thisCurrentSheet.BeginRowDrag(row);
                if (this.thisRowDrag is null)
                {
                    this.EndRowDrag();
                    return;
                }

                // Frozen BEFORE anything moves - see the header.
                this.CaptureRowSlots();
                this.ApplyDropPlaceholderClasses();
            }

            e.Handled = true;

            if (this.thisIsApplyingRowDragMove)
            {
                return;
            }

            this.thisIsApplyingRowDragMove = true;
            try
            {
                this.MoveDraggedRowTo(current);
            }
            finally
            {
                this.thisIsApplyingRowDragMove = false;
            }
        }

        private void OnRowDragPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this.thisArmedDragRow is null)
            {
                return;
            }

            if (this.thisRowDrag is not null)
            {
                e.Handled = true;
            }

            this.EndRowDrag();
            e.Pointer.Capture(null);
        }

        // ###########################################################################################
        // Into the row slot under the pointer, by the FROZEN layout - or, with the pointer past the
        // top or bottom of the rows on screen, one live row further that way, scrolled into view,
        // so a row can be carried beyond what is visible by holding it at the edge and moving the
        // pointer. Only that step scrolls, and the slots are captured again after it.
        // ###########################################################################################
        private void MoveDraggedRowTo(Point pointInGrid)
        {
            if (this.thisRowDrag is not { } drag || this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            Rect rows = this.RowsArea();

            if (pointInGrid.Y < rows.Top || pointInGrid.Y > rows.Bottom)
            {
                if (drag.Step(up: pointInGrid.Y < rows.Top))
                {
                    this.TableGrid.ScrollIntoView(drag.Row, null);
                    this.TableGrid.UpdateLayout();
                    this.CaptureRowSlots();
                    this.ApplyDropPlaceholderClasses();
                }

                return;
            }

            int slot = this.RowSlotAt(pointInGrid.Y);
            int index = slot < 0 || this.thisFirstSlotIndex < 0 ? -1 : this.thisFirstSlotIndex + slot;

            if (index >= 0 && index < sheet.Rows.Count && drag.MoveOnto(sheet.Rows[index]))
            {
                this.ApplyDropPlaceholderClasses();
            }
        }

        // Which frozen slot a height in the grid falls in, or -1 between or beyond them.
        private int RowSlotAt(double y)
        {
            for (int i = 0; i < this.thisRowSlots.Count; i++)
            {
                if (y >= this.thisRowSlots[i].Top && y < this.thisRowSlots[i].Bottom)
                {
                    return i;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // Freezes the row slots as the grid shows them now - every row container on screen, top to
        // bottom, and which row the first one shows. Containers the grid keeps for recycling but
        // is not showing are left out: only a slot that can be seen can be pointed at.
        // ###########################################################################################
        private void CaptureRowSlots()
        {
            this.thisRowSlots.Clear();
            this.thisFirstSlotIndex = -1;

            if (this.thisCurrentSheet is not { } sheet)
            {
                return;
            }

            Rect area = this.RowsArea();

            var shown = this.TableGrid.GetVisualDescendants()
                .OfType<DataGridRow>()
                .Where(container => container.IsVisible && container.DataContext is BoardTableRow)
                .Select(container => (
                    Row: (BoardTableRow)container.DataContext!,
                    Top: container.TranslatePoint(default, this.TableGrid)?.Y,
                    container.Bounds.Height))
                .Where(slot => slot.Top is { } top && top + slot.Height > area.Top && top < area.Bottom)
                .OrderBy(slot => slot.Top!.Value)
                .ToList();

            if (shown.Count == 0)
            {
                return;
            }

            this.thisFirstSlotIndex = sheet.Rows.IndexOf(shown[0].Row);

            foreach (var slot in shown)
            {
                this.thisRowSlots.Add((slot.Top!.Value, slot.Top.Value + slot.Height));
            }
        }

        // Where the rows are drawn, in the grid's coordinates - below the column headers.
        private Rect RowsArea()
        {
            DataGridRowsPresenter? presenter = this.TableGrid.GetVisualDescendants().OfType<DataGridRowsPresenter>().FirstOrDefault();

            if (presenter?.TranslatePoint(default, this.TableGrid) is not { } topLeft)
            {
                return new Rect(this.TableGrid.Bounds.Size);
            }

            return new Rect(topLeft, presenter.Bounds.Size);
        }

        // ###########################################################################################
        // Ends a drag - or a press that never became one. The row stays where it was dragged to; the
        // model decides whether that was a change (see BoardTableRowDrag.Finish). Safe to call when
        // nothing is being dragged, which is how a sheet switch or a reload cancels one cleanly.
        // ###########################################################################################
        private void EndRowDrag()
        {
            BoardTableRowDrag? drag = this.thisRowDrag;

            this.thisRowDrag = null;
            this.thisArmedDragRow = null;
            this.thisRowSlots.Clear();
            this.thisFirstSlotIndex = -1;

            if (drag is null)
            {
                return;
            }

            drag.Finish();
            this.ApplyDropPlaceholderClasses();

            if (this.thisCurrentSheet?.Rows.Contains(drag.Row) == true)
            {
                this.SelectCell(drag.Row, this.CurrentCell?.ColumnIndex ?? 0);
            }

            this.UpdateToolbar();
            this.FocusGrid();
        }

        // ###########################################################################################
        // The dragged row's container wears the placeholder class; every other one does not. Set on
        // every row shown, after every move: rows are recycled as they move and scroll, so the class
        // has to follow the ROW, not stay on whichever container first had it. OnLoadingRow does the
        // same for containers the grid prepares later.
        // ###########################################################################################
        private void ApplyDropPlaceholderClasses()
        {
            foreach (DataGridRow shown in this.TableGrid.GetVisualDescendants().OfType<DataGridRow>())
            {
                shown.Classes.Set(BoardTableEditor.DropPlaceholderClass, this.IsDropPlaceholder(shown.DataContext));
            }
        }

        private bool IsDropPlaceholder(object? item) =>
            this.thisRowDrag is { } drag && ReferenceEquals(item, drag.Row);
    }
}
