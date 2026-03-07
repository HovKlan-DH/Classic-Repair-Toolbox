using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;

namespace TabSchematics
{
    // ###########################################################################################
    // Draws component highlight rectangles over a schematic image.
    // Uses the current view matrix for visible-area culling and keeps stroke thickness stable
    // across zoom levels (by compensating for the view scale).
    // ###########################################################################################
    public sealed class SchematicHighlightsOverlay : Control
    {
//        private readonly List<int> _queryResults = [];
        private readonly List<int> _queryResults = new(); // .NET6 compliant
        private HighlightSpatialIndex? _highlightIndex;

        public HighlightSpatialIndex? HighlightIndex
        {
            get => this._highlightIndex;
            set
            {
                this._highlightIndex = value;
                this.InvalidateVisual();
            }
        }

        public PixelSize BitmapPixelSize { get; set; } = new(0, 0);

        public Matrix ViewMatrix { get; set; } = Matrix.Identity;

        public Color HighlightColor { get; set; } = Colors.IndianRed;

        public double HighlightOpacity { get; set; } = 0.20;

        // ###########################################################################################
        // Forces a re-render whenever the control is re-arranged (e.g. after a splitter drag),
        // ensuring highlights are redrawn with up-to-date bounds.
        // ###########################################################################################
        protected override Size ArrangeOverride(Size finalSize)
        {
            var result = base.ArrangeOverride(finalSize);
            this.InvalidateVisual();
            return result;
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            var index = this._highlightIndex;
            if (index == null || index.Count == 0)
                return;

            if (this.Bounds.Width <= 0 || this.Bounds.Height <= 0)
                return;

            if (this.BitmapPixelSize.Width <= 0 || this.BitmapPixelSize.Height <= 0)
                return;

            var contentRect = GetImageContentRect(this.Bounds.Size, this.BitmapPixelSize);

            var viewportRect = new Rect(0, 0, this.Bounds.Width, this.Bounds.Height);
            var visibleLocalRect = viewportRect;

            if (TryInvert(this.ViewMatrix, out var inv))
                visibleLocalRect = viewportRect.TransformToAABB(inv);

            visibleLocalRect = visibleLocalRect.Intersect(contentRect);
            if (visibleLocalRect.Width <= 0 || visibleLocalRect.Height <= 0)
                return;

            var visiblePixelRect = LocalToPixelRect(visibleLocalRect, contentRect, this.BitmapPixelSize);
            if (visiblePixelRect.Width <= 0 || visiblePixelRect.Height <= 0)
                return;

            index.Query(visiblePixelRect, this._queryResults);

            double scale = Math.Max(0.0001, this.ViewMatrix.M11);
            double strokeThickness = Math.Clamp(1.0 / scale, 0.25, 2.0);

            double fillOpacity = Math.Clamp(this.HighlightOpacity, 0.0, 1.0);
            var fillBrush = new SolidColorBrush(this.HighlightColor, fillOpacity);
            var penBrush = new SolidColorBrush(this.HighlightColor, Math.Min(1.0, fillOpacity * 1.4));
            var pen = new Pen(penBrush, strokeThickness);

            for (int i = 0; i < this._queryResults.Count; i++)
            {
                int idx = this._queryResults[i];
                var pixelRect = index.GetRect(idx);
                var localRect = PixelToLocalRect(pixelRect, contentRect, this.BitmapPixelSize);

                if (!localRect.Intersects(visibleLocalRect))
                    continue;

                context.DrawRectangle(fillBrush, pen, localRect);
            }
        }

        // ###########################################################################################
        // Computes the image content rect in the overlay's local coordinate space.
        // SchematicsImage uses HorizontalAlignment="Left" and VerticalAlignment="Top", so the
        // bitmap content always starts at (0, 0) - no centering offset is applied.
        // ###########################################################################################
        private static Rect GetImageContentRect(Size controlSize, PixelSize bitmapPixelSize)
        {
            if (controlSize.Width <= 0 || controlSize.Height <= 0)
                return new Rect(controlSize);

            double containerAspect = controlSize.Width / controlSize.Height;
            double bitmapAspect = (double)bitmapPixelSize.Width / bitmapPixelSize.Height;

            if (bitmapAspect > containerAspect)
            {
                // Width-constrained - content starts at (0, 0), no vertical centering
                return new Rect(0, 0, controlSize.Width, controlSize.Width / bitmapAspect);
            }
            else
            {
                // Height-constrained - content starts at (0, 0), no horizontal centering
                return new Rect(0, 0, controlSize.Height * bitmapAspect, controlSize.Height);
            }
        }

        private static Rect LocalToPixelRect(Rect localRect, Rect contentRect, PixelSize pixelSize)
        {
            double sx = pixelSize.Width / contentRect.Width;
            double sy = pixelSize.Height / contentRect.Height;

            double x = (localRect.X - contentRect.X) * sx;
            double y = (localRect.Y - contentRect.Y) * sy;
            double w = localRect.Width * sx;
            double h = localRect.Height * sy;

            return new Rect(x, y, w, h).Intersect(new Rect(0, 0, pixelSize.Width, pixelSize.Height));
        }

        private static Rect PixelToLocalRect(Rect pixelRect, Rect contentRect, PixelSize pixelSize)
        {
            double sx = contentRect.Width / pixelSize.Width;
            double sy = contentRect.Height / pixelSize.Height;

            double x = contentRect.X + (pixelRect.X * sx);
            double y = contentRect.Y + (pixelRect.Y * sy);
            double w = pixelRect.Width * sx;
            double h = pixelRect.Height * sy;

            return new Rect(x, y, w, h);
        }

        private static bool TryInvert(Matrix m, out Matrix inv)
        {
            // Invert 2D affine matrix:
            // [ a  b  0 ]
            // [ c  d  0 ]
            // [ e  f  1 ]
            double a = m.M11, b = m.M12, c = m.M21, d = m.M22, e = m.M31, f = m.M32;
            double det = (a * d) - (b * c);

            if (Math.Abs(det) < 1e-12)
            {
                inv = Matrix.Identity;
                return false;
            }

            double idet = 1.0 / det;

            double na = d * idet;
            double nb = -b * idet;
            double nc = -c * idet;
            double nd = a * idet;

            double ne = -((e * na) + (f * nc));
            double nf = -((e * nb) + (f * nd));

            inv = new Matrix(na, nb, nc, nd, ne, nf);
            return true;
        }
    }
}