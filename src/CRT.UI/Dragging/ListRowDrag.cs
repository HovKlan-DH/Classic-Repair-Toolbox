using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A row that can be dragged through a list as a dashed placeholder. The row's template shows
    // the placeholder while IsDropPlaceholder is true, PlaceholderHeight tall - both must notify,
    // or the template never hears the change (the worklog's Files list found this out).
    // ###########################################################################################
    public interface IDraggableRow
    {
        bool IsDropPlaceholder { get; set; }

        double PlaceholderHeight { get; set; }
    }

    // ###########################################################################################
    // DRAGGING A ROW UP OR DOWN A LIST, WITH A PLACEHOLDER - the pointer handling, shared by the
    // Drafts tab's "Schematic images" window and the Maintainer tab's drop-down placement
    // (2026-09-27). The maths is RowDragSlots'; this is the Avalonia rim around it. It was written
    // for the schematic images first and lifted out unchanged when the Maintainer tab needed the
    // same thing, so the race below cannot come back in a second copy.
    //
    // Press a row's handle (the host's PointerPressed calls Press), move a few pixels, and the row
    // becomes a dashed slot that travels with the pointer; the other rows stand in the order a
    // drop will produce. Releasing hands the host (row, from, to) - only when the row really moved.
    //
    // *** THE RACE: ROWS THAT "RAPIDLY SWITCH POSITION AND CANNOT SETTLE" (owner report). *** It
    // happened when each move measured the live layout it had just changed. So:
    //   - the slots are FROZEN when the drag starts (RowDragSlots), and every later decision reads
    //     that snapshot - the same pointer position always picks the same slot;
    //   - a move is applied under a re-entrancy guard: reordering the rows re-creates containers
    //     under the pointer, which raises further pointer events synchronously, and each of those
    //     would otherwise move again;
    //   - the drop target is where the placeholder already IS, never a fresh measurement against
    //     the pointer, so the drop cannot land anywhere other than what was shown.
    //
    // AUTO-SCROLL: holding the row near the top or bottom of the scroll viewer's viewport scrolls
    // it (RowDragSlots.EdgeScrollDelta, on a timer). Safe against the frozen slots because they are
    // in the LIST's coordinates, which scroll with it. EXPLICIT POINTER CAPTURE on the list keeps
    // the drag receiving the pointer when it leaves the list or the window, which is exactly where
    // one pushes to scroll.
    //
    // The move and release handlers sit on the HOST, tunnelling: the pressed row turns into the
    // placeholder the moment the drag starts, which takes its own handlers out of the tree. With
    // the list holding the capture, every event routes through the host on its way to it.
    //
    // A drag that ends without a release (the capture taken away - the window losing focus, say)
    // puts the rows BACK in the order they had and reports nothing: what is on screen never
    // disagrees with what the host has saved.
    // ###########################################################################################
    internal sealed class ListRowDrag<TRow>
        where TRow : class, IDraggableRow
    {
        // Far enough that a click with a shaky hand is not a drag - the worklog's value.
        private const double DragThreshold = 4.0;

        // The strip at the top and bottom of the viewport that scrolls it, and the most one tick may
        // scroll. See RowDragSlots.EdgeScrollDelta.
        private const double AutoScrollBand = 40.0;
        private const double AutoScrollMaxStep = 18.0;
        private static readonly TimeSpan AutoScrollInterval = TimeSpan.FromMilliseconds(30);

        private readonly ItemsControl thisList;
        private readonly ScrollViewer? thisViewer;
        private readonly ObservableCollection<TRow> thisRows;
        private readonly Action<TRow, int, int> thisDropped;
        private readonly Func<TRow, bool>? thisCanDrag;

        // The row pressed on - set by the press, whether or not it has become a drag yet.
        private TRow? thisPressedRow;

        private Point thisStartPoint;

        private bool thisIsApplyingMove;

        // Where the drag started, and the whole order then: an unchanged drop reports nothing, and a
        // cancelled drag is put back exactly.
        private int thisStartIndex = -1;
        private List<TRow> thisOrderBeforeDrag = new();

        // The frozen slots - see the header.
        private readonly List<double> thisMidpoints = new();

        // The pointer as last seen, in the viewport, for the auto-scroll ticks - which fire with no
        // pointer event of their own.
        private Point? thisLastPointerInViewport;

        private IPointer? thisCapturedPointer;

        private DispatcherTimer? thisAutoScrollTimer;

        // ###########################################################################################
        // `host` is any ancestor of the list that sees the whole drag (the window, or the view the
        // list sits in). `viewer` scrolls the list - null when nothing does, and then there is no
        // auto-scroll. `dropped` is told (row, from, to) after a drop that moved the row. `canDrag`
        // refuses a press on a row that must stay where it is.
        // ###########################################################################################
        public ListRowDrag(
            InputElement host,
            ItemsControl list,
            ScrollViewer? viewer,
            ObservableCollection<TRow> rows,
            Action<TRow, int, int> dropped,
            Func<TRow, bool>? canDrag = null)
        {
            ArgumentNullException.ThrowIfNull(host);

            this.thisList = list ?? throw new ArgumentNullException(nameof(list));
            this.thisViewer = viewer;
            this.thisRows = rows ?? throw new ArgumentNullException(nameof(rows));
            this.thisDropped = dropped ?? throw new ArgumentNullException(nameof(dropped));
            this.thisCanDrag = canDrag;

            host.AddHandler(InputElement.PointerMovedEvent, this.OnPointerMoved, RoutingStrategies.Tunnel);
            host.AddHandler(InputElement.PointerReleasedEvent, this.OnPointerReleased, RoutingStrategies.Tunnel);

            list.PointerCaptureLost += this.OnPointerCaptureLost;
        }

        // True from the first move past the threshold until the drop, cancel or reset.
        public bool IsDragging { get; private set; }

        // ###########################################################################################
        // A press on a row's handle ARMS the drag; only movement past the threshold starts it, so a
        // plain click can never reorder anything. The host's handle calls this from its own
        // PointerPressed - a button on the row handles its own press, so it never arrives here.
        // ###########################################################################################
        public void Press(object? sender, PointerPressedEventArgs e)
        {
            if (sender is not Control { DataContext: TRow row } control ||
                !e.GetCurrentPoint(control).Properties.IsLeftButtonPressed ||
                this.IsDragging ||
                (this.thisCanDrag is not null && !this.thisCanDrag(row)))
            {
                return;
            }

            this.thisPressedRow = row;
            this.thisStartPoint = e.GetPosition(this.thisList);
        }

        private void OnPointerMoved(object? sender, PointerEventArgs e)
        {
            if (this.thisPressedRow is null)
            {
                return;
            }

            if (!e.GetCurrentPoint(this.thisList).Properties.IsLeftButtonPressed)
            {
                // The release happened somewhere that never reached the release handler. Treat
                // this as that release - a drop where the placeholder is - rather than leaving a
                // drag the user has already ended running on.
                this.Finish();
                return;
            }

            Point current = e.GetPosition(this.thisList);

            if (this.thisViewer is not null)
            {
                this.thisLastPointerInViewport = e.GetPosition(this.thisViewer);
            }

            if (!this.IsDragging)
            {
                if (Math.Abs(current.Y - this.thisStartPoint.Y) < DragThreshold &&
                    Math.Abs(current.X - this.thisStartPoint.X) < DragThreshold)
                {
                    return;
                }

                // Only a placeholder that was really established starts the drag - the row can have
                // gone between the press and the first move.
                if (!this.Begin(e.Pointer))
                {
                    this.Reset();
                    return;
                }
            }

            this.MovePlaceholderTo(current.Y);
        }

        private void OnPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this.thisPressedRow is null)
            {
                return;
            }

            if (!this.IsDragging)
            {
                // A click, not a drag.
                this.Reset();
                return;
            }

            this.Finish();
        }

        // ###########################################################################################
        // The capture taken away mid-drag - by the system, not by a release (which ends the drag
        // BEFORE its own capture goes, so it never arrives here as a drag). Nothing is reported and
        // the rows go back to where they were.
        // ###########################################################################################
        private void OnPointerCaptureLost(object? sender, PointerCaptureLostEventArgs e)
        {
            if (this.IsDragging)
            {
                this.Cancel();
            }
        }

        // ###########################################################################################
        // Turns the pressed row into the placeholder, freezes the slots, captures the pointer and
        // starts the auto-scroll. Returns false when no placeholder could be established.
        // ###########################################################################################
        private bool Begin(IPointer pointer)
        {
            this.thisMidpoints.Clear();

            TRow? row = this.thisPressedRow;
            int index = row is null ? -1 : this.thisRows.IndexOf(row);

            if (row is null || index < 0)
            {
                return false;
            }

            // The gap is exactly the row it replaces. The containers carry no margin of their own
            // (hosts space the rows with the panel), so the container IS the row's height.
            Control? container = this.thisList.ContainerFromIndex(index);
            if (container is not null && container.Bounds.Height > 0)
            {
                row.PlaceholderHeight = container.Bounds.Height;
            }

            // The slots, from the layout as it is BEFORE anything moves.
            var measured = new List<(double? Top, double Height)>(this.thisRows.Count);

            for (int i = 0; i < this.thisRows.Count; i++)
            {
                Control? rowContainer = this.thisList.ContainerFromIndex(i);
                Point? topLeft = rowContainer?.TranslatePoint(new Point(0, 0), this.thisList);

                double height = rowContainer is not null && rowContainer.Bounds.Height > 0
                    ? rowContainer.Bounds.Height
                    : this.thisRows[i].PlaceholderHeight;

                measured.Add((topLeft?.Y, height));
            }

            this.thisMidpoints.AddRange(RowDragSlots.BuildMidpoints(measured));

            this.thisStartIndex = index;
            this.thisOrderBeforeDrag = this.thisRows.ToList();

            row.IsDropPlaceholder = true;
            this.IsDragging = true;

            pointer.Capture(this.thisList);
            this.thisCapturedPointer = pointer;

            if (this.thisViewer is not null)
            {
                if (this.thisAutoScrollTimer is null)
                {
                    this.thisAutoScrollTimer = new DispatcherTimer(DispatcherPriority.Input)
                    {
                        Interval = AutoScrollInterval,
                    };

                    this.thisAutoScrollTimer.Tick += (_, _) => this.AutoScroll();
                }

                this.thisAutoScrollTimer.Start();
            }

            return true;
        }

        // ###########################################################################################
        // Moves the placeholder to the slot the pointer is over - once per slot change, and never
        // re-entrantly (see the header).
        // ###########################################################################################
        private void MovePlaceholderTo(double pointerYInList)
        {
            if (this.thisIsApplyingMove || this.thisPressedRow is null)
            {
                return;
            }

            this.thisIsApplyingMove = true;

            try
            {
                int target = RowDragSlots.ResolveDropIndex(this.thisMidpoints, pointerYInList, this.thisStartIndex);
                int current = this.thisRows.IndexOf(this.thisPressedRow);

                if (target < 0 || current < 0)
                {
                    return;
                }

                target = Math.Clamp(target, 0, this.thisRows.Count - 1);

                // Leaving the collection alone when the slot has not changed matters: a Move on
                // every pointer frame would rebuild containers continuously and flicker the list.
                if (target != current)
                {
                    this.thisRows.Move(current, target);
                }
            }
            finally
            {
                this.thisIsApplyingMove = false;
            }
        }

        // ###########################################################################################
        // One auto-scroll tick. The placeholder follows the pointer against the layout as it is
        // NOW (the previous tick's scroll has been laid out by then), and the list is then scrolled
        // for the next tick - so every decision reads a settled layout.
        // ###########################################################################################
        private void AutoScroll()
        {
            ScrollViewer? viewer = this.thisViewer;

            if (!this.IsDragging || viewer is null || this.thisLastPointerInViewport is not Point pointer)
            {
                return;
            }

            Point? inList = viewer.TranslatePoint(pointer, this.thisList);
            if (inList is Point listPoint)
            {
                this.MovePlaceholderTo(listPoint.Y);
            }

            double delta = RowDragSlots.EdgeScrollDelta(pointer.Y, viewer.Viewport.Height, AutoScrollBand, AutoScrollMaxStep);

            if (delta == 0)
            {
                return;
            }

            double maximum = Math.Max(0, viewer.Extent.Height - viewer.Viewport.Height);
            double next = Math.Clamp(viewer.Offset.Y + delta, 0, maximum);

            if (Math.Abs(next - viewer.Offset.Y) > 0.1)
            {
                viewer.Offset = new Vector(viewer.Offset.X, next);
            }
        }

        // Lets a headless test run one auto-scroll tick - DispatcherTimer does not tick there.
        internal void AutoScrollForTests() => this.AutoScroll();

        // Lets a headless test take the capture away mid-drag the way the system does, through the
        // real PointerCaptureLost path rather than by calling the cancel directly.
        internal void LoseCaptureForTests() => this.thisCapturedPointer?.Capture(null);

        // ###########################################################################################
        // The drop: the placeholder's position IS the new place, and an unchanged position reports
        // nothing (a drag back to where it started is not an edit).
        // ###########################################################################################
        private void Finish()
        {
            TRow? row = this.thisPressedRow;
            bool wasDragging = this.IsDragging;
            int from = this.thisStartIndex;

            this.Reset();

            if (!wasDragging || row is null)
            {
                return;
            }

            int to = this.thisRows.IndexOf(row);

            if (to < 0 || to == from)
            {
                return;
            }

            this.thisDropped(row, from, to);
        }

        private void Cancel()
        {
            List<TRow> original = this.thisOrderBeforeDrag;

            this.Reset();

            for (int i = 0; i < original.Count; i++)
            {
                int current = this.thisRows.IndexOf(original[i]);

                if (current >= 0 && current != i)
                {
                    this.thisRows.Move(current, i);
                }
            }
        }

        // ###########################################################################################
        // Ends any drag and returns every row to its ordinary look - ALL rows, so an interrupted
        // drag can never leave one drawn as a permanent gap. The host calls it too, whenever it
        // replaces the rows. The flags go first and the capture last: releasing the capture raises
        // PointerCaptureLost at once, and it must find no drag.
        // ###########################################################################################
        public void Reset()
        {
            this.IsDragging = false;
            this.thisPressedRow = null;
            this.thisStartIndex = -1;
            this.thisLastPointerInViewport = null;
            this.thisMidpoints.Clear();

            this.thisAutoScrollTimer?.Stop();

            foreach (TRow row in this.thisRows)
            {
                row.IsDropPlaceholder = false;
            }

            IPointer? pointer = this.thisCapturedPointer;
            this.thisCapturedPointer = null;

            if (pointer is not null && ReferenceEquals(pointer.Captured, this.thisList))
            {
                pointer.Capture(null);
            }
        }
    }
}
