using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Placement maths for the on-schematic badges that must keep a constant size on screen while
    // the board itself zooms - the worklog "#N" pills, and anything else pinned to a point on the
    // image rather than to the viewport.
    //
    // Extracted here because the compensation involved is easy to get wrong in a way that only
    // shows up at zoom levels other than the one it was eyeballed at, which is exactly the kind of
    // bug a test should be holding rather than a person re-checking by hand.
    // ###########################################################################################
    public static class BadgeGeometry
    {
        // ###########################################################################################
        // The offset to add to an anchor point so a centre-scaled badge's TOP-LEFT corner lands on
        // that anchor, at any scale.
        //
        // Canvas.SetLeft/SetTop position a control's pre-transform layout box. A ScaleTransform with
        // RenderTransformOrigin 0.5,0.5 then scales about the box's centre: the centre stays put and
        // the edges move, so the rendered top-left is
        //
        //     layoutTopLeft + unscaledSize/2 - (unscaledSize * scale)/2
        //
        // Setting that equal to the anchor and solving for layoutTopLeft gives the offset returned
        // here. At scale 1 it is zero (nothing to compensate); below 1 it is positive, above 1
        // negative.
        //
        // Subtracting only half the SCALED size - the obvious-looking shortcut - is correct at
        // scale 0.5 alone and drifts everywhere else, which is how the worklog badges came to slide
        // away from their marked areas as the user zoomed in.
        // ###########################################################################################
        // The offset to add to an anchor point so a centre-scaled badge's CENTRE lands on that
        // anchor, at any scale - leaving the badge straddling the point, half of it either side.
        //
        // Simpler than pinning a corner: a centred ScaleTransform leaves the layout box's centre
        // exactly where it was, so the compensation a corner would need cancels out entirely and
        // the offset is just half the UNSCALED size back.
        //
        // There is no scale parameter, deliberately. An argument the body ignores invites a caller
        // to believe it matters and hides a wrong value behind a silent no-op; the independence is
        // better stated by the signature than by a comment. GetCenterScaledRenderedTopLeft below is
        // where the scale genuinely does matter.
        // ###########################################################################################
        public static Point GetCenterScaledCentreOffset(Size unscaledSize)
        {
            if (unscaledSize.Width <= 0 && unscaledSize.Height <= 0)
            {
                return new Point(0, 0);
            }

            return new Point(-(unscaledSize.Width / 2.0), -(unscaledSize.Height / 2.0));
        }

        // ###########################################################################################
        // Where a centre-scaled badge's CENTRE actually renders, given the layout position it was
        // placed at. The counterpart to GetCenterScaledRenderedTopLeft, for the centred case.
        // ###########################################################################################
        // Nudges a badge back inside the viewport when its anchor sits close enough to an edge that
        // part of it would be clipped.
        //
        // The badges straddle their anchor, so an area whose corner is near the left edge of the
        // view puts half the badge off-screen where its "#N" cannot be read and it cannot be
        // clicked. Pushing it in by just the overhang keeps it touching the edge it belongs to,
        // rather than snapping it to some arbitrary inset.
        //
        // Works entirely in the viewport's own coordinate space: the caller converts to and from
        // whatever transform the badge canvas carries. Returns the adjustment to ADD to the badge's
        // rendered top-left, which is (0,0) whenever it already fits - so a caller can apply it
        // unconditionally.
        //
        // A badge larger than the viewport is pinned to the top-left rather than being pushed back
        // and forth between two edges it cannot satisfy at once.
        // ###########################################################################################
        public static Point GetViewportNudge(Rect renderedBadge, Size viewportSize, double margin = 0.0)
        {
            if (viewportSize.Width <= 0 || viewportSize.Height <= 0)
            {
                return new Point(0, 0);
            }

            double dx = 0;
            double dy = 0;

            double left = margin;
            double top = margin;
            double right = viewportSize.Width - margin;
            double bottom = viewportSize.Height - margin;

            if (renderedBadge.Width < right - left)
            {
                if (renderedBadge.Left < left)
                {
                    dx = left - renderedBadge.Left;
                }
                else if (renderedBadge.Right > right)
                {
                    dx = right - renderedBadge.Right;
                }
            }
            else
            {
                // Wider than the space available - pin the left edge so the "#N" stays readable.
                dx = left - renderedBadge.Left;
            }

            if (renderedBadge.Height < bottom - top)
            {
                if (renderedBadge.Top < top)
                {
                    dy = top - renderedBadge.Top;
                }
                else if (renderedBadge.Bottom > bottom)
                {
                    dy = bottom - renderedBadge.Bottom;
                }
            }
            else
            {
                dy = top - renderedBadge.Top;
            }

            return new Point(dx, dy);
        }

        // ###########################################################################################
        // Where a centre-scaled badge's top-left corner actually renders, given the layout position
        // it was placed at. The inverse of the offset above, and the thing a test can assert
        // directly: place a badge with the offset applied and this must return the anchor unchanged.
        // ###########################################################################################
        public static Point GetCenterScaledRenderedTopLeft(Point layoutTopLeft, Size unscaledSize, double scale)
        {
            double x = layoutTopLeft.X + (unscaledSize.Width / 2.0) - (unscaledSize.Width * scale / 2.0);
            double y = layoutTopLeft.Y + (unscaledSize.Height / 2.0) - (unscaledSize.Height * scale / 2.0);

            return new Point(x, y);
        }
    }
}
