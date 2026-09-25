using Avalonia;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // The eight small marker segments drawn at a selected rectangle's corners and side centres -
    // the visual affordance that says "this can be resized".
    //
    // Two overlays draw them: the component label editor's selected highlight, and a hovered
    // worklog area. They had the same ~55 lines of sizing maths copied into each, which meant a
    // tweak to marker sizing had to be made twice or the two would silently diverge - and the
    // comment claiming they matched was an invariant nothing enforced. Building the rectangles
    // here leaves each overlay with only its own brush and DrawRectangle calls.
    //
    // Everything is sized in SCREEN terms (divided by scale) so the markers stay a constant size
    // however far the board is zoomed, matching the hit rectangles in LabelEditorGeometry.
    // ###########################################################################################
    public static class SelectionMarkerGeometry
    {
        // ###########################################################################################
        // The marker rectangles for the given border rect, in the same local coordinate space.
        //
        // On a small rectangle the side markers shrink and then disappear entirely rather than
        // overlapping the corner markers - two markers running together would read as one long
        // edge handle and imply a resize behaviour the hit rectangles do not offer. The corners
        // themselves also shrink so they can never exceed half the rectangle.
        // ###########################################################################################
        public static List<Rect> BuildSelectionMarkerRects(Rect rect, double scale)
        {
            double safeScale = Math.Max(0.0001, scale);

            double markerThickness = Math.Clamp(2.5 / safeScale, 1.0, 2.5);
            double baseCornerLength = Math.Clamp(6.5 / safeScale, 3.0, 6.5);
            double baseSideLength = Math.Clamp(5.0 / safeScale, 2.5, 5.5);
            double halfThickness = markerThickness / 2.0;

            double maxCornerLengthX = Math.Max(markerThickness, (rect.Width / 2.0) + halfThickness);
            double maxCornerLengthY = Math.Max(markerThickness, (rect.Height / 2.0) + halfThickness);

            double cornerLengthX = Math.Min(baseCornerLength, maxCornerLengthX);
            double cornerLengthY = Math.Min(baseCornerLength, maxCornerLengthY);

            double minimumGap = Math.Clamp(2.0 / safeScale, markerThickness, 3.0);

            double horizontalSideLength = Math.Max(0.0, rect.Width - (cornerLengthX * 2.0) - minimumGap);
            double verticalSideLength = Math.Max(0.0, rect.Height - (cornerLengthY * 2.0) - minimumGap);

            if (horizontalSideLength > 0.0)
            {
                horizontalSideLength = Math.Min(baseSideLength, horizontalSideLength);
            }

            if (verticalSideLength > 0.0)
            {
                verticalSideLength = Math.Min(baseSideLength, verticalSideLength);
            }

            double horizontalSideHalf = horizontalSideLength / 2.0;
            double verticalSideHalf = verticalSideLength / 2.0;

            double left = rect.Left;
            double top = rect.Top;
            double right = rect.Right;
            double bottom = rect.Bottom;
            double centerX = rect.Center.X;
            double centerY = rect.Center.Y;

            var rects = new List<Rect>(12)
            {
                // Each corner is an L: one horizontal arm and one vertical arm.
                new Rect(left - halfThickness, top - halfThickness, cornerLengthX, markerThickness),
                new Rect(left - halfThickness, top - halfThickness, markerThickness, cornerLengthY),

                new Rect(right - cornerLengthX + halfThickness, top - halfThickness, cornerLengthX, markerThickness),
                new Rect(right - halfThickness, top - halfThickness, markerThickness, cornerLengthY),

                new Rect(left - halfThickness, bottom - halfThickness, cornerLengthX, markerThickness),
                new Rect(left - halfThickness, bottom - cornerLengthY + halfThickness, markerThickness, cornerLengthY),

                new Rect(right - cornerLengthX + halfThickness, bottom - halfThickness, cornerLengthX, markerThickness),
                new Rect(right - halfThickness, bottom - cornerLengthY + halfThickness, markerThickness, cornerLengthY),
            };

            if (horizontalSideLength > 0.0)
            {
                rects.Add(new Rect(centerX - horizontalSideHalf, top - halfThickness, horizontalSideLength, markerThickness));
                rects.Add(new Rect(centerX - horizontalSideHalf, bottom - halfThickness, horizontalSideLength, markerThickness));
            }

            if (verticalSideLength > 0.0)
            {
                rects.Add(new Rect(left - halfThickness, centerY - verticalSideHalf, markerThickness, verticalSideLength));
                rects.Add(new Rect(right - halfThickness, centerY - verticalSideHalf, markerThickness, verticalSideLength));
            }

            return rects;
        }
    }
}
