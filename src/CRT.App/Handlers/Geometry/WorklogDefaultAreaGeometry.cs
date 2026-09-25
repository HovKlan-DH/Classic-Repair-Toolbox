using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Gives a worklog entry that was never drawn on a schematic a sensible area to be marked at,
    // for the moment its "Show marked area" is first ticked.
    //
    // Most worklog entries get their area from a drag on the schematic. One kind does not: an entry
    // created from the oscilloscope capture's "Attach image to worklog" flow, where the user is at
    // the bench with a probe rather than at the schematic view. Those are stored with a zero-sized
    // area and ShowMarkedArea off, so they park as a corner pill (see ParkedBadgeGeometry) rather
    // than drawing a rectangle.
    //
    // The problem is what happens when the user later TICKS "Show marked area" on one of those. A
    // zero-sized rect is a degenerate rectangle: it draws as nothing, or as a hairline, and it can
    // never be grabbed and dragged into place - so the entry would look broken with no way to fix
    // it from the UI. This gives it a real, visible, grabbable square in the board's bottom-right
    // corner instead, which the user can then drag to where it belongs.
    //
    // Bottom-right specifically: the parked pills live in the TOP-right (ParkedBadgeGeometry), so
    // an area appearing at the opposite corner cannot be mistaken for one of those, and the two
    // never overlap while an entry is mid-way through being moved.
    //
    // Everything is in the schematic's own PIXEL coordinates, which is how WorklogEntryRecord
    // stores an area - not in viewport space, so it is independent of zoom and pan.
    // ###########################################################################################
    public static class WorklogDefaultAreaGeometry
    {
        // The square's side as a fraction of the smaller image dimension. Big enough to see and to
        // grab on a 4000px board scan, small enough not to blanket a small one.
        private const double SideFraction = 0.08;

        // Clamps for that side in pixels, so a very large or very small schematic still yields a
        // usable handle rather than something either unusable or absurd.
        private const double MinimumSide = 24.0;
        private const double MaximumSide = 400.0;

        // How far the square sits in from the image's bottom-right edge, as a fraction of the side.
        // A small inset keeps the whole rectangle - including its stroke - inside the image, so it
        // does not read as clipped.
        private const double InsetFraction = 0.25;

        // ###########################################################################################
        // Whether an entry's stored area is one that was never actually drawn - i.e. it has no usable
        // width or height, so there is nothing to show even if ShowMarkedArea were ticked.
        //
        // Tested rather than compared to exactly zero: an area that has been through a JSON
        // round-trip, or a hand edit, can carry a tiny non-zero value that is still not a rectangle
        // anyone could see or grab. A negative one is likewise unusable.
        // ###########################################################################################
        public static bool IsUnset(Rect area) => area.Width < 1.0 || area.Height < 1.0;

        // ###########################################################################################
        // A default marked area for a schematic of the given pixel size: a square inset from the
        // bottom-right corner.
        //
        // An image with no usable size yields Rect.Empty - the caller then simply leaves the entry
        // parked, which is the honest outcome: there is no board to place anything on.
        // ###########################################################################################
        public static Rect BuildDefaultArea(Size imagePixelSize)
        {
            double width = imagePixelSize.Width;
            double height = imagePixelSize.Height;

            if (width < 1.0 || height < 1.0)
            {
                return default;
            }

            double side = Math.Clamp(Math.Min(width, height) * SideFraction, MinimumSide, MaximumSide);

            // A tiny image can be smaller than the clamped minimum, which would push the square off
            // the board entirely. Fitting it to the image keeps the whole rectangle inside.
            side = Math.Min(side, Math.Min(width, height));

            double inset = side * InsetFraction;

            // Clamped at zero so the square never starts off the left or top edge of a small image,
            // where the inset alone could push it past the origin.
            double x = Math.Max(0.0, width - side - inset);
            double y = Math.Max(0.0, height - side - inset);

            return new Rect(x, y, side, side);
        }

        // ###########################################################################################
        // The area an entry should be given when its marked area is being shown: its own, when it
        // already has a real one, and otherwise a fresh default.
        //
        // This is the single decision point for "does this entry need an area inventing", so the
        // editor and anything else that ticks the box on the user's behalf cannot disagree about it.
        // ###########################################################################################
        public static Rect ResolveAreaForShowing(Rect existingArea, Size imagePixelSize) =>
            IsUnset(existingArea) ? BuildDefaultArea(imagePixelSize) : existingArea;
    }
}
