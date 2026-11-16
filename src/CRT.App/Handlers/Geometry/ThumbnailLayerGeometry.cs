using Avalonia;
using System;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE SIZE A THUMBNAIL'S HIGHLIGHT LAYER IS DRAWN AT (2026-11-16, performance walkthrough).
    //
    // The layer was drawn at the thumbnail's own size - up to 800 wide (AppConfig.ThumbnailMaxWidth) - on the UI thread,
    // for every thumbnail whose highlights changed: typing in the component filter selects what it finds, and each
    // pause drew a layer per matching thumbnail, up to 25 on a large board, each a new native bitmap. The tiles show a
    // thumbnail some 100-250 wide in the tab; the layer is stretched over its thumbnail as the thumbnail is (both
    // Stretch="Uniform" in one cell), so it is drawn at most MaxWidth wide, its proportions kept - a third as many
    // pixels for an 800-wide thumbnail. Never wider than the thumbnail, never below 1x1.
    // ###########################################################################################
    public static class ThumbnailLayerGeometry
    {
        public const int MaxWidth = 480;

        public static PixelSize SizeFor(PixelSize thumbnail)
        {
            int width = Math.Max(1, thumbnail.Width);
            int height = Math.Max(1, thumbnail.Height);

            if (width <= ThumbnailLayerGeometry.MaxWidth)
            {
                return new PixelSize(width, height);
            }

            double scale = ThumbnailLayerGeometry.MaxWidth / (double)width;
            return new PixelSize(ThumbnailLayerGeometry.MaxWidth, Math.Max(1, (int)Math.Round(height * scale)));
        }
    }
}
