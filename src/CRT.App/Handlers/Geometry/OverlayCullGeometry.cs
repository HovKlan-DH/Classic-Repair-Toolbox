using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Bounding-box maths for deciding which overlay primitives are on screen, extracted from the
    // KiCad overlay control so it can be tested without a display.
    //
    // Why this matters: the trace overlay's Render() was measured at ~50 ms for 6,692 primitives,
    // and profiling showed it runs on every zoom step - so it draws the whole board even when the
    // viewport only shows a corner of it. Skipping primitives that cannot be seen is the direct
    // saving. The bounds must be generous rather than exact: a primitive wrongly judged off-screen
    // disappears, which is far worse than one drawn needlessly.
    // ###########################################################################################
    public static class OverlayCullGeometry
    {
        // ###########################################################################################
        // Grows a bounding box by half a stroke width, since a stroke straddles the path it follows.
        // Round caps and joins can push a little further, so the whole thickness is added rather than
        // half of it.
        // ###########################################################################################
        public static Rect InflateForStroke(Rect bounds, double thickness)
        {
            double margin = double.IsNaN(thickness) || double.IsInfinity(thickness)
                ? 1.0
                : Math.Max(1.0, Math.Abs(thickness));

            return bounds.Inflate(margin);
        }

        // ###########################################################################################
        // Returns the bounding box of a run of points, or an empty rect when there are none.
        // ###########################################################################################
        public static Rect BoundsOfPoints(IReadOnlyList<Point> points)
        {
            if (points == null || points.Count == 0)
            {
                return default;
            }

            double minX = points[0].X;
            double maxX = points[0].X;
            double minY = points[0].Y;
            double maxY = points[0].Y;

            for (int i = 1; i < points.Count; i++)
            {
                if (points[i].X < minX) minX = points[i].X;
                if (points[i].X > maxX) maxX = points[i].X;
                if (points[i].Y < minY) minY = points[i].Y;
                if (points[i].Y > maxY) maxY = points[i].Y;
            }

            return new Rect(minX, minY, maxX - minX, maxY - minY);
        }

        // ###########################################################################################
        // Returns a box that contains a rect at any rotation about its own centre.
        //
        // The circumscribed circle is used rather than the exact rotated box: it is correct for every
        // angle, costs nothing to compute, and only ever errs towards drawing something that turned
        // out not to be visible.
        // ###########################################################################################
        public static Rect BoundsOfRotatedRect(Rect rect, double rotationDegrees)
        {
            if (KiCadPadGeometry.IsAxisAligned(rotationDegrees))
            {
                return rect;
            }

            double diagonal = Math.Sqrt((rect.Width * rect.Width) + (rect.Height * rect.Height));
            Point centre = rect.Center;

            return new Rect(
                centre.X - (diagonal / 2.0),
                centre.Y - (diagonal / 2.0),
                diagonal,
                diagonal);
        }

        // ###########################################################################################
        // Maps the on-screen viewport back into the overlay's own coordinate space, so primitives can
        // be tested against it without transforming each one.
        //
        // A view matrix that cannot be inverted means the view is degenerate; the whole viewport is
        // returned in that case so everything still draws rather than the overlay going blank.
        // ###########################################################################################
        public static Rect GetVisibleLocalRect(Rect viewport, Matrix viewMatrix)
        {
            if (!RectGeometry.TryInvert(viewMatrix, out var inverse))
            {
                return viewport;
            }

            return viewport.TransformToAABB(inverse);
        }

        // ###########################################################################################
        // Returns true when a primitive should be drawn.
        //
        // An empty bounds is treated as visible. Some primitives legitimately have no extent of their
        // own, and refusing to draw them would be a silent visual regression rather than a saving.
        // ###########################################################################################
        public static bool IsVisible(Rect primitiveBounds, Rect visibleRect)
        {
            if (primitiveBounds.Width <= 0 && primitiveBounds.Height <= 0)
            {
                return true;
            }

            return primitiveBounds.Intersects(visibleRect);
        }
    }
}
