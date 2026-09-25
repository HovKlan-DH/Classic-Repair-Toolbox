using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // The maths behind the user-drawn traces on a schematic: converting between the normalized
    // 0..1 coordinates a trace is STORED in and the canvas coordinates it is DRAWN in, deciding
    // whether a click landed on a node or a segment, and snapping a dragged node to its
    // neighbours.
    //
    // Traces are persisted normalized (0..1 of the image content rect) rather than in canvas
    // pixels, so the same trace lands in the right place at any zoom level and on any window
    // size. Everything here is that conversion and the hit-testing that depends on it.
    //
    // Extracted from Tabs/Schematics/PolylineManagement.cs, which keeps the parts that genuinely
    // own Avalonia objects: ManagedPolyline wraps a real Polyline shape and its Ellipse markers,
    // so walking the drawn traces stays there and calls into these helpers per node/segment.
    // ###########################################################################################
    internal static class TraceGeometry
    {
        // ###########################################################################################
        // Canvas point -> normalized 0..1 within the image content rect, clamped to the image.
        //
        // Returns (0,0) for a degenerate rect: the image has not been laid out or measured yet, and
        // dividing by its zero width would otherwise produce NaN and poison every stored node.
        // ###########################################################################################
        public static Point CanvasToNormalized(Point canvasPoint, Rect contentRect)
        {
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return new Point(0, 0);
            }

            double nx = (canvasPoint.X - contentRect.X) / contentRect.Width;
            double ny = (canvasPoint.Y - contentRect.Y) / contentRect.Height;

            return new Point(
                Math.Max(0.0, Math.Min(1.0, nx)),
                Math.Max(0.0, Math.Min(1.0, ny)));
        }

        // ###########################################################################################
        // Normalized 0..1 -> canvas point. The inverse of CanvasToNormalized, minus the clamp:
        // the caller may legitimately be mapping a node that sits outside the visible content.
        //
        // A degenerate rect returns the input unchanged rather than (0,0) - the caller is mid-layout
        // and will re-map once real bounds arrive; collapsing every node onto the origin in the
        // meantime would be visible.
        // ###########################################################################################
        public static Point NormalizedToCanvas(Point normalizedPoint, Rect contentRect)
        {
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return normalizedPoint;
            }

            return new Point(
                contentRect.X + (normalizedPoint.X * contentRect.Width),
                contentRect.Y + (normalizedPoint.Y * contentRect.Height));
        }

        // ###########################################################################################
        // Snaps a dragged node onto its immediate neighbours' axes, so a trace drawn by hand can be
        // made exactly horizontal or vertical without pixel-perfect mouse work.
        //
        // The two axes are decided INDEPENDENTLY: X may snap to one neighbour while Y snaps to the
        // other, which is what lets a corner node line up with the node before it horizontally and
        // the node after it vertically at the same time. Only the immediately adjacent nodes are
        // candidates - snapping to a distant node in the same trace would move the line somewhere
        // the user cannot see the reason for.
        //
        // neighbourCanvasPoints is what the caller supplies: the previous and/or next node already
        // converted to canvas coordinates (one of them, at an end of the trace; both in the middle).
        // ###########################################################################################
        public static Point ApplyNodeSnapping(
            Point currentCanvasPoint,
            IReadOnlyList<Point> neighbourCanvasPoints,
            double tolerance)
        {
            double snapX = currentCanvasPoint.X;
            double snapY = currentCanvasPoint.Y;

            double closestXDistance = tolerance;
            double closestYDistance = tolerance;

            foreach (var neighbour in neighbourCanvasPoints)
            {
                double dx = Math.Abs(currentCanvasPoint.X - neighbour.X);
                if (dx < closestXDistance)
                {
                    snapX = neighbour.X;
                    closestXDistance = dx;
                }

                double dy = Math.Abs(currentCanvasPoint.Y - neighbour.Y);
                if (dy < closestYDistance)
                {
                    snapY = neighbour.Y;
                    closestYDistance = dy;
                }
            }

            return new Point(snapX, snapY);
        }

        public static double DistanceSquared(Point a, Point b)
        {
            return ((a.X - b.X) * (a.X - b.X)) + ((a.Y - b.Y) * (a.Y - b.Y));
        }

        public static double Distance(Point a, Point b) => Math.Sqrt(DistanceSquared(a, b));

        // ###########################################################################################
        // Shortest distance from a point to a LINE SEGMENT (not an infinite line), plus the point on
        // the segment closest to it.
        //
        // The projection is what a click on a trace becomes when the user inserts a node: it lands
        // exactly on the line rather than where the pointer was, so the trace does not kink.
        //
        // t is clamped to 0..1, which is the whole difference between a segment and a line - without
        // it, clicking off the end of a short segment would project onto empty space beyond it.
        // A zero-length segment (both endpoints equal) short-circuits, since normalising by its
        // length would divide by zero.
        // ###########################################################################################
        public static double DistancePointToSegment(Point point, Point segmentStart, Point segmentEnd, out Point projection)
        {
            double lengthSquared = DistanceSquared(segmentStart, segmentEnd);

            if (lengthSquared == 0.0)
            {
                projection = segmentStart;
                return Distance(point, segmentStart);
            }

            double t = Math.Max(0, Math.Min(1, (((point.X - segmentStart.X) * (segmentEnd.X - segmentStart.X)) +
                                                ((point.Y - segmentStart.Y) * (segmentEnd.Y - segmentStart.Y))) / lengthSquared));

            projection = new Point(
                segmentStart.X + (t * (segmentEnd.X - segmentStart.X)),
                segmentStart.Y + (t * (segmentEnd.Y - segmentStart.Y)));

            return Distance(point, projection);
        }

        // ###########################################################################################
        // Whether a trace loaded from disk is in the LEGACY canvas-pixel format rather than the
        // current normalized one.
        //
        // Traces used to be saved as raw canvas coordinates, which broke as soon as the window was
        // resized or the image zoomed. The test is whether any coordinate falls OUTSIDE the range a
        // normalized one can occupy: normalized values are 0..1 by construction, so anything above
        // that - or below zero - cannot be one. The margin above 1.0 is deliberate slack for
        // rounding at the extreme edge of an image.
        //
        // The negative half of the test matters as much as the positive one. Canvas pixels could be
        // negative (a node drawn while the image was panned off the left or top edge), and a legacy
        // trace lying entirely at negative coordinates passes an upper-bound-only check: it would be
        // read as already normalized, and ToNormalized would then pass e.g. -15 through untouched
        // for NormalizedToCanvas to map thousands of pixels off-image. A normalized coordinate is
        // never negative, so nothing is lost by rejecting one.
        //
        // This is a heuristic rather than a version flag because the old format carried no version
        // marker - this is what the data itself can tell us. PolylineManagement applies it with
        // Any() across a trace's nodes, so one out-of-range node is enough to convert the whole
        // trace, which is what makes the heuristic safe: it only has to catch a legacy trace
        // SOMEWHERE, not at every node.
        // ###########################################################################################
        public static bool IsLegacyCanvasCoordinate(double x, double y)
        {
            return x > 2.0 || y > 2.0 || x < 0.0 || y < 0.0;
        }
    }
}
