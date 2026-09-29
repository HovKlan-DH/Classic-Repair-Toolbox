using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE MATHS OF DRAGGING A ROW UP OR DOWN A LIST WITH A PLACEHOLDER - where the pointer says
    // the row should go. Shared by the worklog editor's Photos and Files lists, and - through
    // ListRowDrag - the Drafts tab's "Schematic images" window and the Maintainer tab's
    // drop-down placement (2026-09-27), so the one thing that went wrong once cannot go wrong
    // again in another copy. In CRT.UI since then, because both applications use it; still
    // Avalonia-free, and still tested by CRT.App.Tests' RowDragSlotsTests.
    //
    // *** WHAT WENT WRONG: THE ROWS OSCILLATED. *** Measuring the LIVE layout feeds each move back
    // into its own input: moving the placeholder re-lays out the list, which moves the rows the
    // next measurement reads, which can pick a different slot, which moves them back - every frame
    // for as long as the pointer is held there. Rows of different heights make it unavoidable,
    // because a swap then shifts the layout by the difference between two heights.
    //
    // So the slots are FROZEN when the drag starts (BuildMidpoints, from the layout as it was
    // before anything moved) and every later decision reads that snapshot (ResolveDropIndex). The
    // same pointer position then always gives the same slot, and there is nothing to oscillate.
    //
    // Everything here is in the LIST's own coordinates, which scroll with it - so scrolling during
    // a drag (the wheel, or EdgeScrollDelta's auto-scroll) moves the pointer through the frozen
    // slots exactly as it should, without re-measuring anything.
    // ###########################################################################################
    internal static class RowDragSlots
    {
        // ###########################################################################################
        // The midpoint of every row, top to bottom, from each row's top and height in the list.
        //
        // One midpoint per row, ALWAYS - index i means row i. A row with no measured top (its
        // container not realised) is placed straight after the previous one rather than skipped:
        // skipping would shorten the list and shift every later midpoint's meaning by one, and
        // ResolveDropIndex relies on the 1:1 correspondence to answer an index into the rows.
        // ###########################################################################################
        public static List<double> BuildMidpoints(IReadOnlyList<(double? Top, double Height)> rows)
        {
            var midpoints = new List<double>(rows?.Count ?? 0);
            double runningY = 0;

            foreach ((double? measuredTop, double height) in rows ?? [])
            {
                double top = measuredTop ?? runningY;

                midpoints.Add(top + (height / 2.0));

                runningY = top + height;
            }

            return midpoints;
        }

        // ###########################################################################################
        // Which index the dragged row belongs at: the number of OTHER rows whose middle the pointer
        // is past. Above every middle (or above the list entirely) is 0, and past them all is the
        // last index, so a drag flung past either end lands AT that end rather than being thrown
        // away. -1 only when there are no slots at all.
        //
        // draggedIndex is where the dragged row stood when the slots were frozen.
        //
        // *** THE DRAGGED ROW'S OWN MIDDLE IS NOT COUNTED (2026-09-27). *** The worklog's version
        // answered "the first row whose middle the pointer is above", which counts it - and the
        // answer is an index in the list WITHOUT the dragged row (ObservableCollection.Move removes
        // it first). Dragging DOWN therefore landed one row too far: crossing its own middle already
        // put it below the next row, so the placeholder ran a row ahead of the pointer the whole
        // way down. Dragging up was right, because the row's own middle is then below the pointer.
        // ###########################################################################################
        public static int ResolveDropIndex(IReadOnlyList<double> midpoints, double pointerY, int draggedIndex)
        {
            if (midpoints is null || midpoints.Count == 0)
            {
                return -1;
            }

            int rowsAbove = 0;

            for (int i = 0; i < midpoints.Count; i++)
            {
                if (i != draggedIndex && pointerY >= midpoints[i])
                {
                    rowsAbove++;
                }
            }

            return Math.Min(rowsAbove, midpoints.Count - 1);
        }

        // ###########################################################################################
        // How far to scroll a list's viewport on one auto-scroll tick while a row is dragged near
        // (or past) its top or bottom edge: negative scrolls up, positive down, 0 leaves it alone.
        //
        // Faster the deeper into the edge band the pointer is, reaching maxStep at the edge and
        // staying there beyond it - so resting just inside the band creeps, and pushing past the
        // edge moves briskly. The band is capped at a third of the viewport, so on a short list the
        // two bands can never overlap and scroll it both ways at once.
        // ###########################################################################################
        public static double EdgeScrollDelta(double pointerY, double viewportHeight, double band, double maxStep)
        {
            if (viewportHeight <= 0 || band <= 0 || maxStep <= 0)
            {
                return 0;
            }

            double effectiveBand = Math.Min(band, viewportHeight / 3.0);

            if (pointerY < effectiveBand)
            {
                double depth = Math.Min(1.0, (effectiveBand - pointerY) / effectiveBand);
                return -Math.Max(1.0, maxStep * depth);
            }

            double bottomEdge = viewportHeight - effectiveBand;

            if (pointerY > bottomEdge)
            {
                double depth = Math.Min(1.0, (pointerY - bottomEdge) / effectiveBand);
                return Math.Max(1.0, maxStep * depth);
            }

            return 0;
        }
    }
}
