using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A PICTURE AND ITS MARKED COMPONENTS ARE DRAWN IN THE PANE'S PREVIEW (CRT 4.0 step 2, slice
    // 4b - the plan: "a colour picker and opacity slider per picture, with a live preview on the
    // picture"). The picture fitted whole into the area, centred; a highlight - kept in the picture's
    // own pixels - scaled and moved with it.
    // ###########################################################################################
    public static class PicturePreviewGeometry
    {
        // The rectangle `image` (its pixel size) is drawn in: as large as `area` allows, its sides kept
        // in proportion, centred. Empty for an image or an area of no size.
        public static Rect Fit(Size image, Size area)
        {
            if (image.Width <= 0 || image.Height <= 0 || area.Width <= 0 || area.Height <= 0)
            {
                return default;
            }

            double scale = Math.Min(area.Width / image.Width, area.Height / image.Height);
            double width = image.Width * scale;
            double height = image.Height * scale;

            return new Rect((area.Width - width) / 2.0, (area.Height - height) / 2.0, width, height);
        }

        // A highlight in the picture's pixels, where it is drawn on the picture fitted into `target`.
        public static Rect Place(Rect highlight, Size image, Rect target)
        {
            if (image.Width <= 0 || image.Height <= 0)
            {
                return default;
            }

            double scaleX = target.Width / image.Width;
            double scaleY = target.Height / image.Height;

            return new Rect(
                target.X + (highlight.X * scaleX),
                target.Y + (highlight.Y * scaleY),
                highlight.Width * scaleX,
                highlight.Height * scaleY);
        }
    }
}
