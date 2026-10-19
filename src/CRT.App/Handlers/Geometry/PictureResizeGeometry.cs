using System;
using Handlers.DataHandling;

namespace Handlers.Geometry
{
    // The four corners of a picture in the Feedback tab's text, each a handle that resizes it.
    public enum PictureCorner
    {
        TopLeft,
        TopRight,
        BottomLeft,
        BottomRight
    }

    // ###########################################################################################
    // *** RESIZING A PICTURE IN THE FEEDBACK TEXT BY ITS CORNERS (owner request, 2026-10-15). ***
    // The width chosen is the picture's width in the mail (CRT.Data's FeedbackComposition), its
    // height following the picture's own proportions. FeedbackComposer's handles call this.
    // ###########################################################################################
    public static class PictureResizeGeometry
    {
        // ###########################################################################################
        // The widths a picture may be dragged between: down to MinPictureWidth (or the picture's own
        // width, when smaller), up to its own width or the mail's column, whichever is smaller.
        // ###########################################################################################
        public static (int Min, int Max) WidthRange(int naturalWidth)
        {
            int max = FeedbackComposition.DefaultPictureWidth(naturalWidth);
            return (Math.Min(FeedbackComposition.MinPictureWidth, max), max);
        }

        // The height that goes with a width, in the picture's own proportions; never below 1.
        public static int HeightFor(int width, int naturalWidth, int naturalHeight)
        {
            if (naturalWidth <= 0 || naturalHeight <= 0)
                return Math.Max(1, width);

            return Math.Max(1, (int)Math.Round((double)width * naturalHeight / naturalWidth));
        }

        // ###########################################################################################
        // A corner dragged (dx, dy) from a picture `startWidth` wide: the width that follows. A right
        // corner grows as the pointer goes right, a left one as it goes left; a bottom corner as it
        // goes down, a top one as it goes up - the vertical movement counted in width through the
        // picture's proportions. Of the two, the one that moved the picture more wins, as a corner
        // handle in a drawing program behaves. Clamped to WidthRange.
        // ###########################################################################################
        public static int DraggedWidth(int startWidth, int naturalWidth, int naturalHeight, PictureCorner corner, double dx, double dy)
        {
            double sx = corner is PictureCorner.TopRight or PictureCorner.BottomRight ? 1 : -1;
            double sy = corner is PictureCorner.BottomLeft or PictureCorner.BottomRight ? 1 : -1;
            double aspect = naturalWidth > 0 && naturalHeight > 0 ? (double)naturalWidth / naturalHeight : 1;

            double byX = sx * dx;
            double byY = sy * dy * aspect;
            double delta = Math.Abs(byX) >= Math.Abs(byY) ? byX : byY;

            (int min, int max) = PictureResizeGeometry.WidthRange(naturalWidth);
            return Math.Clamp((int)Math.Round(startWidth + delta), min, max);
        }
    }
}
