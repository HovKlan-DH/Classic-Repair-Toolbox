using Avalonia;
using System;

namespace Handlers.Geometry
{
    // #######################################################################################
    // Where a PICTURE of the KiCad overlay is taken, and where it is put back
    // (KiCadOverlayRenderControl's snapshot - 2026-11-06, owner report: "incredible slow ... when
    // dragging one of the markers").
    //
    // While the calibration box is dragged or nudged, the overlay is the same few thousand traces
    // and pads moved by a matrix (KiCadCalibrationTransform). Drawing them again every frame is
    // what was slow - measured on 250469's top (5,351 primitives): 35-65 ms of Skia raster per
    // frame, every frame. So the overlay is drawn ONCE into a bitmap the size of what is on
    // screen, and each frame draws that one bitmap under the matrix instead.
    //
    // Three spaces are involved:
    //   built   - the coordinates the primitives were built in (the overlay control's own space,
    //             BEFORE its OverlayTransform);
    //   screen  - the control's parent, after the OverlayTransform and the control's own
    //             visual transform (its RenderTransform: the zoom and pan);
    //   picture - the snapshot bitmap, in DIPs, whose top-left is `area`'s top-left on screen.
    //
    // A point p built in the overlay shows on screen at p * overlay * visual (Avalonia's row-vector
    // order: the left matrix applies first). The picture is taken at one (overlay, visual) and drawn
    // back, in the control's local space, under ANOTHER overlay - the matrices below are exactly the
    // composition that keeps every pixel where the vector drawing would put it, which the tests hold.
    // #######################################################################################
    internal static class OverlaySnapshotGeometry
    {
        // How far beyond the view the picture reaches on every side, as a share of the view: a box
        // dragged that far still shows the traces it brings into view.
        public const double MarginShare = 0.25;

        // The most pixels one picture may hold (64 MB at 4 bytes a pixel). Past it the margin goes;
        // past that too there is no picture, and the overlay is drawn as vectors.
        public const long MaximumPixels = 16_000_000;

        // ###################################################################################
        // How many primitives make a picture worth taking. Below it the vectors are as cheap to draw
        // as the picture, and sharper: measured (Skia, 1410 x 987), the Open128's "Power" sheet's 222
        // primitives drew in 2.5 ms against 6 ms for its picture, 250469's top side's 5,351 in 35-65 ms
        // against 4.5 - some 0.012 ms a primitive, so even at about 500.
        // ###################################################################################
        public const int PrimitivesWorthAPicture = 500;

        public static bool IsWorthAPicture(int primitives, int worthAPicture = PrimitivesWorthAPicture) =>
            primitives >= worthAPicture;

        // ###################################################################################
        // The part of the screen worth a picture: the view with its margin, cut to where the overlay
        // draws anything, its edges on whole device pixels (so a picture drawn back unmoved lands
        // pixel for pixel). Empty when nothing of the overlay is near the view - a picture of
        // nothing, which is still a picture. Null when no picture can be taken: no view, no scaling,
        // or more pixels than MaximumPixels even without the margin.
        // ###################################################################################
        public static Rect? AreaToCapture(Size viewSize, Rect contentOnScreen, double scaling)
        {
            if (viewSize.Width <= 0 || viewSize.Height <= 0 || scaling <= 0 || double.IsNaN(scaling))
            {
                return null;
            }

            var view = new Rect(viewSize);
            var withMargin = view.Inflate(new Thickness(viewSize.Width * MarginShare, viewSize.Height * MarginShare));

            foreach (Rect around in new[] { withMargin, view })
            {
                Rect area = OverlaySnapshotGeometry.OnWholePixels(around.Intersect(contentOnScreen), scaling);

                if (area.Width <= 0 || area.Height <= 0)
                {
                    return default(Rect);
                }

                if (OverlaySnapshotGeometry.PixelSizeOf(area, scaling) is { } pixels &&
                    (long)pixels.Width * pixels.Height <= MaximumPixels)
                {
                    return area;
                }
            }

            return null;
        }

        // ###################################################################################
        // A rect grown outward to whole device pixels: its edges on multiples of 1 / scaling.
        // ###################################################################################
        public static Rect OnWholePixels(Rect rect, double scaling)
        {
            if (rect.Width <= 0 || rect.Height <= 0 || scaling <= 0)
            {
                return default;
            }

            double left = Math.Floor(rect.Left * scaling) / scaling;
            double top = Math.Floor(rect.Top * scaling) / scaling;
            double right = Math.Ceiling(rect.Right * scaling) / scaling;
            double bottom = Math.Ceiling(rect.Bottom * scaling) / scaling;

            return new Rect(left, top, right - left, bottom - top);
        }

        // The device pixels a picture of `area` takes; null for an empty area.
        public static PixelSize? PixelSizeOf(Rect area, double scaling)
        {
            if (area.Width <= 0 || area.Height <= 0 || scaling <= 0)
            {
                return null;
            }

            return new PixelSize(
                Math.Max(1, (int)Math.Round(area.Width * scaling)),
                Math.Max(1, (int)Math.Round(area.Height * scaling)));
        }

        // ###################################################################################
        // *** THE PART OF THE PICTURE DRAWN BACK, IN ITS PIXELS - NEVER ITS SIZE IN DIPS (owner report,
        // 2026-11-06: "it looks like it is moving or zooming the entire trace block slightly ... as soon
        // as I grab a marker"). *** DrawImage's source rectangle reaches Skia as it is, which counts it
        // in the image's own pixels (Avalonia 12.1.4: WriteableBitmapImpl.Draw -> Canvas.DrawImage). The
        // picture is taken at the screen's scaling, so on a display at 125% its size in DIPs is 80% of its
        // pixels: given that, only the top-left 80% was drawn, stretched over the whole area - a 25% zoom
        // the moment the picture took over from the traces. At 100% the two are the same number, so
        // nothing but a scaled display ever showed it.
        // ###################################################################################
        public static Rect SourceRectOf(PixelSize picture) => new(0, 0, picture.Width, picture.Height);

        // ###################################################################################
        // *** WHETHER A PICTURE STILL HOLDS EVERYTHING OF THE OVERLAY THE VIEW SHOWS NOW (code review, 2026-11-07). ***
        // A picture reaches only MarginShare past the view as it was when taken, and it was kept from one move to the next
        // as long as the zoom stayed - so a box moved past the margin, a little at a time, drew a band with no traces in it,
        // and a picture of nothing (the overlay away from the view as it was taken) drew nothing at all once the overlay came
        // into view. The overlay's content as it lies on screen under `overlayNow`, cut to the view, must lie inside the
        // picture as it is drawn back (TryPictureToLocal). Nothing of the overlay in view needs nothing.
        // ###################################################################################
        public static bool StillCovers(Size viewSize, Rect builtContent, Rect area, Matrix overlayThen, Matrix overlayNow, Matrix visual)
        {
            Rect needed = new Rect(viewSize).Intersect(builtContent.TransformToAABB(overlayNow * visual));

            if (needed.Width <= 0 || needed.Height <= 0)
            {
                return true;
            }

            if (area.Width <= 0 || area.Height <= 0 ||
                !OverlaySnapshotGeometry.TryPictureToLocal(area, overlayThen, visual, overlayNow, out Matrix pictureToLocal) ||
                !(pictureToLocal * visual).TryInvert(out Matrix screenToPicture))
            {
                return false;
            }

            // A hair of rounding is no gap.
            return new Rect(area.Size).Inflate(0.01).Contains(needed.TransformToAABB(screenToPicture));
        }

        // ###################################################################################
        // Built space to picture space: what the primitives are drawn under to take the picture.
        // ###################################################################################
        public static Matrix BuiltToPicture(Rect area, Matrix overlay, Matrix visual) =>
            overlay * visual * Matrix.CreateTranslation(-area.X, -area.Y);

        // ###################################################################################
        // Picture space to the control's local space, for a picture taken under (overlayThen,
        // visualThen) drawn back while the overlay transform is `overlayNow`. The control's visual
        // transform is applied after this by the compositor, as for everything the control draws -
        // so with nothing changed, picture to screen is a plain move by the area's corner.
        //
        // False when either matrix the picture was taken under cannot be undone.
        // ###################################################################################
        public static bool TryPictureToLocal(
            Rect area,
            Matrix overlayThen,
            Matrix visualThen,
            Matrix overlayNow,
            out Matrix pictureToLocal)
        {
            pictureToLocal = Matrix.Identity;

            if (!visualThen.TryInvert(out Matrix visualInverse) ||
                !overlayThen.TryInvert(out Matrix overlayInverse))
            {
                return false;
            }

            pictureToLocal = Matrix.CreateTranslation(area.X, area.Y) * visualInverse * overlayInverse * overlayNow;
            return true;
        }
    }
}
