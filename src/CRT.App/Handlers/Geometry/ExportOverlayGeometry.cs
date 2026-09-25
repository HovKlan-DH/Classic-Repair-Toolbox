using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Where a worklog's marked area and its "#N" badge land on a schematic image inside the
    // exported PDF, as FRACTIONS of the drawn image (0..1 across and down).
    //
    // WHY FRACTIONS, AND WHY THIS IS A SEPARATE CLASS:
    //
    // An entry's area is stored in the schematic's own PIXEL coordinates (a 3552x2477 board scan
    // records an area at x=2060), while the exported page draws that image at whatever width the
    // page margins leave - a number QuestPDF works out during layout and never tells us. The only
    // stable way to place anything over it is proportionally.
    //
    // The first shipped version got this wrong in a way worth recording, because the shape of the
    // mistake is easy to repeat: it computed the fractions correctly and then passed them to
    // QuestPDF's PaddingLeft/PaddingTop multiplied by 100, believing those took a percentage.
    // They do not - every QuestPDF padding is an absolute LENGTH (points by default; there is no
    // percentage unit anywhere in the API). So "58% from the left" became "58 points from the
    // left", and a marked area that covers a tenth of the board was drawn covering most of it.
    // Nothing threw, and the document looked plausible until it was held next to the screen.
    //
    // Keeping the maths here rather than inside the writer is what makes that class of error
    // catchable: these numbers can be asserted directly against the same pixel rects the UI
    // draws from, with no PDF, no page size, and no QuestPDF involved.
    // ###########################################################################################
    public static class ExportOverlayGeometry
    {
        // One marked area as fractions of the drawn image: Left/Top are its top-left corner,
        // Width/Height its size. All four are 0..1.
        public readonly record struct AreaFractions(double Left, double Top, double Width, double Height)
        {
            // What is left over to the right of / below the area. The exported layout builds its
            // bands from these, and computing them here keeps the "1 - left - width" arithmetic
            // (and its rounding) in one place rather than at three call sites.
            public double RemainingRight => Math.Max(0.0, 1.0 - this.Left - this.Width);

            public double RemainingBottom => Math.Max(0.0, 1.0 - this.Top - this.Height);
        }

        // ###########################################################################################
        // Converts one entry's pixel area into fractions of the image.
        //
        // CLIPPED, not rejected: a board image replaced by a differently-sized scan after an area
        // was drawn leaves coordinates that fall partly outside it. What remains visible is still
        // where the work was, so the area is intersected with the picture rather than dropped - a
        // missing rectangle tells the reader nothing.
        //
        // Note it is the INTERSECTION, not the original size moved inward. Clamping only the
        // origin keeps the full width and shifts the rectangle off the thing it marks, which is
        // worse than either clipping or dropping it.
        //
        // A zero or negative image dimension returns null: there is no meaningful fraction of a
        // zero-width image, and the caller draws the picture with no overlay at all.
        // ###########################################################################################
        public static AreaFractions? TryBuildAreaFractions(
            double areaX, double areaY, double areaWidth, double areaHeight,
            int imagePixelWidth, int imagePixelHeight)
        {
            if (imagePixelWidth <= 0 || imagePixelHeight <= 0)
                return null;

            if (areaWidth <= 0 || areaHeight <= 0)
                return null;

            // CLIPPED, not merely clamped: both edges are converted first and then intersected
            // with the image, so an area hanging off the left or top loses the part that is not
            // visible instead of keeping its full size and sliding inward.
            //
            // Clamping the ORIGIN alone was wrong in a way that looked right: an entry at
            // areaX = -50, areaWidth = 100 on a 1000px image clamped its left edge to 0 but kept
            // width at 0.1, drawing a rectangle twice the visible size and offset to the right of
            // the copper it marks. Deriving the far edge before clamping is what keeps the drawn
            // rectangle the intersection of the area and the picture.
            double rawLeft = areaX / imagePixelWidth;
            double rawTop = areaY / imagePixelHeight;
            double rawRight = (areaX + areaWidth) / imagePixelWidth;
            double rawBottom = (areaY + areaHeight) / imagePixelHeight;

            double left = Clamp01(rawLeft);
            double top = Clamp01(rawTop);
            double right = Clamp01(rawRight);
            double bottom = Clamp01(rawBottom);

            double width = right - left;
            double height = bottom - top;

            // An area clipped away to nothing (entirely off the image) is not worth drawing.
            if (width <= 0 || height <= 0)
                return null;

            return new AreaFractions(left, top, width, height);
        }

        // ###########################################################################################
        // The ASPECT RATIO a band of the image occupies, given its size as fractions.
        //
        // This is the one number the exported layout is actually built from. QuestPDF can size a
        // container by aspect ratio against whatever width it is given, so a band that must be
        // `widthFraction` across and `heightFraction` down of an image whose own pixel ratio is
        // known can be expressed as a single width/height ratio and left to lay itself out - with
        // no page dimension appearing anywhere.
        //
        //     band width  = widthFraction  * imageWidth
        //     band height = heightFraction * imageHeight
        //     ratio       = (widthFraction * imageWidth) / (heightFraction * imageHeight)
        //
        // Returns null when either fraction is zero - a band with no extent has no ratio, and the
        // caller simply omits it rather than emitting a degenerate container.
        // ###########################################################################################
        public static double? TryBuildBandAspectRatio(
            double widthFraction, double heightFraction, int imagePixelWidth, int imagePixelHeight)
        {
            if (widthFraction <= 0 || heightFraction <= 0)
                return null;

            if (imagePixelWidth <= 0 || imagePixelHeight <= 0)
                return null;

            double ratio = (widthFraction * imagePixelWidth) / (heightFraction * imagePixelHeight);

            return double.IsFinite(ratio) && ratio > 0 ? ratio : null;
        }

        private static double Clamp01(double value) =>
            double.IsFinite(value) ? Math.Clamp(value, 0.0, 1.0) : 0.0;
    }
}
