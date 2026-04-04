using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // Draws editable component highlight rectangles over the schematic using the same base color
    // as the normal highlight overlay, with selected-state corner and side line markers.
    // ###########################################################################################
    public sealed class ComponentLabelEditorOverlay : Control
    {
        private IReadOnlyList<Rect> thisRectangles = Array.Empty<Rect>();
        private int thisSelectedIndex = -1;
        private int thisHoveredIndex = -1;
        private PixelSize thisBitmapPixelSize = new(0, 0);
        private Matrix thisViewMatrix = Matrix.Identity;
        private Color thisHighlightColor = Colors.IndianRed;
        private double thisHighlightOpacity = 0.20;
        private Rect? thisDraftRectangle;

        public IReadOnlyList<Rect> Rectangles
        {
            get => this.thisRectangles;
            set
            {
                this.thisRectangles = value ?? Array.Empty<Rect>();
                this.InvalidateVisual();
            }
        }

        public int SelectedIndex
        {
            get => this.thisSelectedIndex;
            set
            {
                this.thisSelectedIndex = value;
                this.InvalidateVisual();
            }
        }

        public int HoveredIndex
        {
            get => this.thisHoveredIndex;
            set
            {
                this.thisHoveredIndex = value;
                this.InvalidateVisual();
            }
        }

        public PixelSize BitmapPixelSize
        {
            get => this.thisBitmapPixelSize;
            set
            {
                this.thisBitmapPixelSize = value;
                this.InvalidateVisual();
            }
        }

        public Matrix ViewMatrix
        {
            get => this.thisViewMatrix;
            set
            {
                this.thisViewMatrix = value;
                this.InvalidateVisual();
            }
        }

        public Color HighlightColor
        {
            get => this.thisHighlightColor;
            set
            {
                this.thisHighlightColor = value;
                this.InvalidateVisual();
            }
        }

        public double HighlightOpacity
        {
            get => this.thisHighlightOpacity;
            set
            {
                this.thisHighlightOpacity = value;
                this.InvalidateVisual();
            }
        }

        public Rect? DraftRectangle
        {
            get => this.thisDraftRectangle;
            set
            {
                this.thisDraftRectangle = value;
                this.InvalidateVisual();
            }
        }

        // ###########################################################################################
        // Forces a redraw whenever arrange changes so overlay stays aligned after layout changes.
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

            if (this.Bounds.Width <= 0 || this.Bounds.Height <= 0)
            {
                return;
            }

            if (this.thisBitmapPixelSize.Width <= 0 || this.thisBitmapPixelSize.Height <= 0)
            {
                return;
            }

            var contentRect = GetImageContentRect(this.Bounds.Size, this.thisBitmapPixelSize);
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return;
            }

            double scale = Math.Max(0.0001, this.thisViewMatrix.M11);
            double borderThickness = Math.Clamp(1.0 / scale, 0.5, 1.0);
            double fillOpacity = Math.Clamp(this.thisHighlightOpacity, 0.0, 1.0);

            var fillBrush = new SolidColorBrush(this.thisHighlightColor, fillOpacity);
            var normalPen = new Pen(new SolidColorBrush(this.thisHighlightColor, Math.Min(1.0, fillOpacity * 1.25)), borderThickness);
            var selectedPen = new Pen(new SolidColorBrush(this.thisHighlightColor, 1.0), borderThickness);
            var draftFillBrush = new SolidColorBrush(this.thisHighlightColor, Math.Min(0.12, fillOpacity));
            var draftPen = new Pen(new SolidColorBrush(this.thisHighlightColor, 1.0), borderThickness);

            for (int i = 0; i < this.thisRectangles.Count; i++)
            {
                var pixelRect = this.thisRectangles[i];
                var localRect = PixelToLocalRect(pixelRect, contentRect, this.thisBitmapPixelSize);
                var borderRect = InsetRectForStroke(localRect, borderThickness);

                bool isSelected = i == this.thisSelectedIndex;
                bool showMarkers = isSelected && i == this.thisHoveredIndex; 

                context.DrawRectangle(fillBrush, null, localRect);
                context.DrawRectangle(null, isSelected ? selectedPen : normalPen, borderRect);

                if (showMarkers)
                {
                    this.DrawSelectionMarkers(context, borderRect, scale);
                }
            }

            if (this.thisDraftRectangle.HasValue)
            {
                var localDraftRect = PixelToLocalRect(this.thisDraftRectangle.Value, contentRect, this.thisBitmapPixelSize);
                var draftBorderRect = InsetRectForStroke(localDraftRect, borderThickness);

                context.DrawRectangle(draftFillBrush, null, localDraftRect);
                context.DrawRectangle(null, draftPen, draftBorderRect);
            }
        }

        // ###########################################################################################
        // Insets a rectangle by half the stroke thickness so the drawn border remains visually
        // inside the original bounds instead of growing outward.
        // ###########################################################################################
        private static Rect InsetRectForStroke(Rect rect, double strokeThickness)
        {
            double inset = strokeThickness / 2.0;
            double width = Math.Max(0.0, rect.Width - strokeThickness);
            double height = Math.Max(0.0, rect.Height - strokeThickness);

            return new Rect(rect.X + inset, rect.Y + inset, width, height);
        }

        // ###########################################################################################
        // Draws compact square marker segments at the 4 corners and 4 side centers of the selected
        // rectangle. Rectangles are used instead of lines so the marker ends stay crisp and square.
        // ###########################################################################################
        private void DrawSelectionMarkers(DrawingContext context, Rect rect, double scale)
        {
            //            double markerThickness = Math.Clamp(3.0 / scale, 1.25, 3.0);
            //            double cornerLength = Math.Clamp(8.0 / scale, 3.5, 8.0);
            //            double sideLength = Math.Clamp(6.0 / scale, 3.0, 7.0);
            double markerThickness = Math.Clamp(2.5 / scale, 1.0, 2.5);
            double cornerLength = Math.Clamp(6.5 / scale, 3.0, 6.5);
            double sideLength = Math.Clamp(5.0 / scale, 2.5, 5.5);
            double halfThickness = markerThickness / 2.0;
            double sideHalf = sideLength / 2.0;

            var markerBrush = new SolidColorBrush(this.thisHighlightColor, 1.0);

            double left = rect.Left;
            double top = rect.Top;
            double right = rect.Right;
            double bottom = rect.Bottom;
            double centerX = rect.Center.X;
            double centerY = rect.Center.Y;

            // Top-left
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(left - halfThickness, top - halfThickness, cornerLength, markerThickness));
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(left - halfThickness, top - halfThickness, markerThickness, cornerLength));

            // Top-right
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(right - cornerLength + halfThickness, top - halfThickness, cornerLength, markerThickness));
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(right - halfThickness, top - halfThickness, markerThickness, cornerLength));

            // Bottom-left
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(left - halfThickness, bottom - halfThickness, cornerLength, markerThickness));
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(left - halfThickness, bottom - cornerLength + halfThickness, markerThickness, cornerLength));

            // Bottom-right
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(right - cornerLength + halfThickness, bottom - halfThickness, cornerLength, markerThickness));
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(right - halfThickness, bottom - cornerLength + halfThickness, markerThickness, cornerLength));

            // Top center
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(centerX - sideHalf, top - halfThickness, sideLength, markerThickness));

            // Bottom center
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(centerX - sideHalf, bottom - halfThickness, sideLength, markerThickness));

            // Left center
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(left - halfThickness, centerY - sideHalf, markerThickness, sideLength));

            // Right center
            context.DrawRectangle(
                markerBrush,
                null,
                new Rect(right - halfThickness, centerY - sideHalf, markerThickness, sideLength));
        }

        // ###########################################################################################
        // Computes the image content rect in the overlay's local coordinate space.
        // Must match the normal schematic highlight overlay exactly, where the image content starts
        // at top-left with no centering offset applied.
        // ###########################################################################################
        private static Rect GetImageContentRect(Size controlSize, PixelSize bitmapPixelSize)
        {
            if (controlSize.Width <= 0 || controlSize.Height <= 0)
            {
                return new Rect(controlSize);
            }

            double containerAspect = controlSize.Width / controlSize.Height;
            double bitmapAspect = (double)bitmapPixelSize.Width / bitmapPixelSize.Height;

            if (bitmapAspect > containerAspect)
            {
                return new Rect(0, 0, controlSize.Width, controlSize.Width / bitmapAspect);
            }
            else
            {
                return new Rect(0, 0, controlSize.Height * bitmapAspect, controlSize.Height);
            }
        }

        // ###########################################################################################
        // Converts a pixel-space rectangle into the overlay's local coordinate system.
        // ###########################################################################################
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
    }
}