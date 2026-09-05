using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using Handlers.Geometry;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // Draws component highlight rectangles over a schematic image.
    // Uses the current view matrix for visible-area culling and keeps stroke thickness stable
    // across zoom levels (by compensating for the view scale).
    // ###########################################################################################
    public sealed class SchematicHighlightsOverlay : Control
    {
        private readonly List<int> _queryResults = [];
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

            var contentRect = RectGeometry.GetImageContentRect(this.Bounds.Size, this.BitmapPixelSize);

            var viewportRect = new Rect(0, 0, this.Bounds.Width, this.Bounds.Height);
            var visibleLocalRect = viewportRect;

            if (RectGeometry.TryInvert(this.ViewMatrix, out var inv))
                visibleLocalRect = viewportRect.TransformToAABB(inv);

            visibleLocalRect = visibleLocalRect.Intersect(contentRect);
            if (visibleLocalRect.Width <= 0 || visibleLocalRect.Height <= 0)
                return;

            var visiblePixelRect = RectGeometry.LocalToPixelRect(visibleLocalRect, contentRect, this.BitmapPixelSize);
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
                var localRect = RectGeometry.PixelToLocalRect(pixelRect, contentRect, this.BitmapPixelSize);

                if (!localRect.Intersects(visibleLocalRect))
                    continue;

                context.DrawRectangle(fillBrush, pen, localRect);
            }
        }
    }
}