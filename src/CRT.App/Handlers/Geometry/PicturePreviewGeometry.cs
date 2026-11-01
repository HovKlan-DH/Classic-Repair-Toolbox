using Avalonia;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A PICTURE AND ITS MARKED COMPONENTS ARE DRAWN IN THE PANE'S PREVIEW (CRT 4.0 step 2, slice
    // 4b - the plan: "a colour picker and opacity slider per picture, with a live preview on the
    // picture"). The picture fitted whole into the area, centred; a highlight - kept in the picture's
    // own pixels - scaled and moved with it.
    //
    // ZOOMED (2026-10-29 - owner request: "maybe to zoom in/out in that small thumbnail so I can view
    // relevant parts of it"): `zoom` times the fitted size, never less than whole (1), and `centre` -
    // the point of the picture in the middle of the area, 0..1 across and down - kept so the picture
    // always covers the area along a side it is larger than, and stays centred along one it is not.
    // Zooming keeps the point under the pointer where it is.
    // ###########################################################################################
    public static class PicturePreviewGeometry
    {
        // How far in the preview zooms - well past a chip's legs on a whole-board photo.
        public const double MaxZoom = 16.0;

        // The point of the picture in the middle of the area while it is shown whole.
        public static Point Middle { get; } = new(0.5, 0.5);

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

        // ###########################################################################################
        // The rectangle the WHOLE picture is drawn in at `zoom` with `centre` in the middle - larger
        // than the area once zoomed in, its parts outside the area cut off by the preview. At zoom 1
        // it is Fit's. Empty for an image or an area of no size.
        // ###########################################################################################
        public static Rect View(Size image, Size area, double zoom, Point centre)
        {
            Rect fit = PicturePreviewGeometry.Fit(image, area);

            if (fit.Width <= 0)
            {
                return default;
            }

            double scale = PicturePreviewGeometry.ZoomWithin(zoom);
            Point kept = PicturePreviewGeometry.ClampCentre(image, area, scale, centre);
            double width = fit.Width * scale;
            double height = fit.Height * scale;

            return new Rect((area.Width / 2.0) - (kept.X * width), (area.Height / 2.0) - (kept.Y * height), width, height);
        }

        // ###########################################################################################
        // `centre` kept so the picture covers the area: along a side the zoomed picture is larger
        // than the area, no further than half the area from its edge; along one it is not, the middle.
        // ###########################################################################################
        public static Point ClampCentre(Size image, Size area, double zoom, Point centre)
        {
            Rect fit = PicturePreviewGeometry.Fit(image, area);

            if (fit.Width <= 0)
            {
                return PicturePreviewGeometry.Middle;
            }

            double scale = PicturePreviewGeometry.ZoomWithin(zoom);

            return new Point(
                PicturePreviewGeometry.ClampAxis(centre.X, area.Width, fit.Width * scale),
                PicturePreviewGeometry.ClampAxis(centre.Y, area.Height, fit.Height * scale));
        }

        // ###########################################################################################
        // Zoomed by `factor` at `pointer` (in the area): the new zoom - within 1..MaxZoom - and the
        // centre that keeps the picture's point under the pointer under it, as far as the picture
        // can stay covering the area.
        // ###########################################################################################
        public static (double Zoom, Point Centre) ZoomAt(Size image, Size area, double zoom, Point centre, Point pointer, double factor)
        {
            Rect view = PicturePreviewGeometry.View(image, area, zoom, centre);
            Rect fit = PicturePreviewGeometry.Fit(image, area);

            if (view.Width <= 0 || fit.Width <= 0 || !double.IsFinite(factor) || factor <= 0)
            {
                return (PicturePreviewGeometry.ZoomWithin(zoom), PicturePreviewGeometry.ClampCentre(image, area, zoom, centre));
            }

            double next = PicturePreviewGeometry.ZoomWithin(PicturePreviewGeometry.ZoomWithin(zoom) * factor);

            // The picture's point under the pointer, 0..1 - where it is to stay.
            double underX = (pointer.X - view.X) / view.Width;
            double underY = (pointer.Y - view.Y) / view.Height;

            double width = fit.Width * next;
            double height = fit.Height * next;

            var wanted = new Point(
                underX + (((area.Width / 2.0) - pointer.X) / width),
                underY + (((area.Height / 2.0) - pointer.Y) / height));

            return (next, PicturePreviewGeometry.ClampCentre(image, area, next, wanted));
        }

        // The centre after the picture was dragged by `delta` (in the area's units) - it moves with the pointer.
        public static Point Pan(Size image, Size area, double zoom, Point centre, Vector delta)
        {
            Rect view = PicturePreviewGeometry.View(image, area, zoom, centre);

            if (view.Width <= 0 || view.Height <= 0)
            {
                return PicturePreviewGeometry.Middle;
            }

            Point kept = PicturePreviewGeometry.ClampCentre(image, area, zoom, centre);

            return PicturePreviewGeometry.ClampCentre(
                image,
                area,
                zoom,
                new Point(kept.X - (delta.X / view.Width), kept.Y - (delta.Y / view.Height)));
        }

        // ###########################################################################################
        // How many of the picture's pixels across the preview needs to draw it sharp: the zoomed
        // picture's width on screen, in device pixels (`scaling`, the screen's). 0 for nothing shown.
        // ###########################################################################################
        public static int PixelsWanted(Size image, Size area, double zoom, double scaling)
        {
            Rect fit = PicturePreviewGeometry.Fit(image, area);

            if (fit.Width <= 0)
            {
                return 0;
            }

            double wanted = fit.Width * PicturePreviewGeometry.ZoomWithin(zoom) * (double.IsFinite(scaling) && scaling > 0 ? scaling : 1.0);

            return (int)Math.Ceiling(Math.Min(wanted, image.Width));
        }

        // ###########################################################################################
        // The widths every card's picture is decoded to for its previews SHOWN WHOLE, smallest first -
        // enough for the previews as wide as they are dragged (up to ContributeWidths.PreviewRange.Max)
        // on a screen scaled up to 150 %. In steps, so dragging the previews wider decodes the cards
        // again only when a drag ends past one (code review, 2026-10-30: decoded at 640 whatever the
        // width, a preview dragged to 900 on a 150 % screen was drawn soft at 1350 device pixels).
        // ###########################################################################################
        public static IReadOnlyList<int> WholeWidths { get; } = [640, 1024, 1536];

        // The width to decode the cards' pictures to for previews `previewWidth` wide on a screen
        // scaled by `scaling`: the first step at least that many device pixels - else the widest.
        public static int WholeWidth(double previewWidth, double scaling)
        {
            double wanted = (double.IsFinite(previewWidth) && previewWidth > 0 ? previewWidth : 0) *
                            (double.IsFinite(scaling) && scaling > 0 ? scaling : 1.0);

            return WholeWidths.FirstOrDefault(step => step >= wanted, WholeWidths[^1]);
        }

        // The widths a zoomed preview's picture is decoded to, sharpest last - each a step up, so a
        // zoom asks for a new one only now and then, and at most one this large is held.
        public static IReadOnlyList<int> SharpWidths { get; } = [1280, 2048, 3072];

        // ###########################################################################################
        // The width to decode the picture to for `wanted` pixels across: the first step at least that
        // wide - else the sharpest - and never wider than the picture itself (`imageWidth`, its own
        // pixels; 0 when not known, which caps nothing).
        // ###########################################################################################
        public static int SharpWidth(int wanted, double imageWidth) =>
            PicturePreviewGeometry.NoWiderThanThePicture(SharpWidths.FirstOrDefault(step => step >= wanted, SharpWidths[^1]), imageWidth);

        // ###########################################################################################
        // `width` - a step of WholeWidths or SharpWidths - never wider than the picture itself
        // (`imageWidth`, its own pixels; 0 when not known, which caps nothing): a picture is never
        // decoded larger than it is (code review, 2026-10-30: a 600 px scan was decoded to 1536).
        // ###########################################################################################
        public static int NoWiderThanThePicture(int width, double imageWidth) =>
            imageWidth >= 1 && imageWidth < width ? (int)Math.Floor(imageWidth) : width;

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

        // A point in the picture's pixels, where it is drawn on the picture fitted into `target`.
        public static Point Place(Point point, Size image, Rect target)
        {
            if (image.Width <= 0 || image.Height <= 0)
            {
                return default;
            }

            return new Point(
                target.X + (point.X * target.Width / image.Width),
                target.Y + (point.Y * target.Height / image.Height));
        }

        private static double ZoomWithin(double zoom) =>
            double.IsFinite(zoom) ? Math.Clamp(zoom, 1.0, PicturePreviewGeometry.MaxZoom) : 1.0;

        // One axis of the centre: the middle when the picture is no larger than the area along it,
        // else no nearer its edge than half the area.
        private static double ClampAxis(double centre, double area, double shown)
        {
            if (shown <= area || shown <= 0)
            {
                return 0.5;
            }

            double half = area / 2.0 / shown;
            double value = double.IsFinite(centre) ? centre : 0.5;

            return Math.Clamp(value, half, 1.0 - half);
        }
    }
}
