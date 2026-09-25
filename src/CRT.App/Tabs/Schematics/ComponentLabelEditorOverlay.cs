using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using Handlers.Geometry;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // Draws editable component highlight rectangles over the schematic using the same base color
    // as the normal highlight overlay, with selected-state corner and side line markers.
    // ###########################################################################################
    public sealed class ComponentLabelEditorOverlay : Control
    {
        private IReadOnlyList<Rect> thisRectangles = Array.Empty<Rect>();
        private IReadOnlyList<int> thisSelectedIndices = Array.Empty<int>();
        private int thisSelectedIndex = -1;
        private int thisHoveredIndex = -1;
        private PixelSize thisBitmapPixelSize = new(0, 0);
        private Matrix thisViewMatrix = Matrix.Identity;
        private Color thisHighlightColor = Colors.IndianRed;
        private double thisHighlightOpacity = 0.20;
        private Rect? thisDraftRectangle;
        private bool thisUseDashedBorder;
        private Rect? thisSelectionBounds;
        private IReadOnlyList<(Point Start, Point End)> thisSnapGuides = Array.Empty<(Point Start, Point End)>();

        public IReadOnlyList<Rect> Rectangles
        {
            get => this.thisRectangles;
            set
            {
                this.thisRectangles = value ?? Array.Empty<Rect>();
                this.InvalidateVisual();
            }
        }

        public IReadOnlyList<int> SelectedIndices
        {
            get => this.thisSelectedIndices;
            set
            {
                this.thisSelectedIndices = value ?? Array.Empty<int>();
                this.InvalidateVisual();
            }
        }

        public IReadOnlyList<(Point Start, Point End)> SnapGuides
        {
            get => this.thisSnapGuides;
            set
            {
                this.thisSnapGuides = value ?? Array.Empty<(Point Start, Point End)>();
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

        // ###########################################################################################
        // When true, rectangle and draft borders are dashed instead of solid - used by the worklog
        // entry-area overlay to match its mockup. The label editor never sets this, so its solid
        // borders are unaffected.
        // ###########################################################################################
        public bool UseDashedBorder
        {
            get => this.thisUseDashedBorder;
            set
            {
                this.thisUseDashedBorder = value;
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

        public Rect? SelectionBounds
        {
            get => this.thisSelectionBounds;
            set
            {
                this.thisSelectionBounds = value;
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

            var contentRect = RectGeometry.GetImageContentRect(this.Bounds.Size, this.thisBitmapPixelSize);
            if (contentRect.Width <= 0 || contentRect.Height <= 0)
            {
                return;
            }

            double scale = Math.Max(0.0001, this.thisViewMatrix.M11);
            double borderThickness = Math.Clamp(1.0 / scale, 0.5, 1.0);
            double fillOpacity = Math.Clamp(this.thisHighlightOpacity, 0.0, 1.0);

            DashStyle? borderDashStyle = this.thisUseDashedBorder
                ? new DashStyle(new[] { Math.Clamp(6.0 / scale, 2.0, 6.0), Math.Clamp(4.0 / scale, 2.0, 4.0) }, 0)
                : null;

            var fillBrush = new SolidColorBrush(this.thisHighlightColor, fillOpacity);
            var normalPen = new Pen(new SolidColorBrush(this.thisHighlightColor, Math.Min(1.0, fillOpacity * 1.25)), borderThickness, borderDashStyle);
            var selectedPen = new Pen(new SolidColorBrush(this.thisHighlightColor, 1.0), borderThickness, borderDashStyle);
            var draftFillBrush = new SolidColorBrush(this.thisHighlightColor, Math.Min(0.12, fillOpacity));
            var draftPen = new Pen(new SolidColorBrush(this.thisHighlightColor, 1.0), borderThickness, borderDashStyle);
            var snapPen = new Pen(
                new SolidColorBrush(this.thisHighlightColor, 1.0),
                2.0 / scale,
                new DashStyle(
                    new[]
                    {
                        Math.Clamp(6.0 / scale, 2.0, 6.0),
                        Math.Clamp(4.0 / scale, 2.0, 4.0)
                    },
                    0),
                PenLineCap.Round,
                PenLineJoin.Round);

            var selectedIndices = new HashSet<int>();

            if (this.thisSelectedIndices.Count > 0)
            {
                foreach (int index in this.thisSelectedIndices)
                {
                    selectedIndices.Add(index);
                }
            }
            else if (this.thisSelectedIndex >= 0)
            {
                selectedIndices.Add(this.thisSelectedIndex);
            }

            for (int i = 0; i < this.thisRectangles.Count; i++)
            {
                var pixelRect = this.thisRectangles[i];
                var localRect = RectGeometry.PixelToLocalRect(pixelRect, contentRect, this.thisBitmapPixelSize);
                var borderRect = RectGeometry.InsetRectForStroke(localRect, borderThickness);

                bool isSelected = selectedIndices.Contains(i);
                bool showMarkers = isSelected && i == this.thisHoveredIndex;

                context.DrawRectangle(fillBrush, null, localRect);
                context.DrawRectangle(null, isSelected ? selectedPen : normalPen, borderRect);

                if (showMarkers)
                {
                    this.DrawSelectionMarkers(context, borderRect, scale);
                }
            }

            foreach (var guide in this.thisSnapGuides)
            {
                var start = RectGeometry.PixelToLocalPoint(guide.Start, contentRect, this.thisBitmapPixelSize);
                var end = RectGeometry.PixelToLocalPoint(guide.End, contentRect, this.thisBitmapPixelSize);

                if (Math.Abs(start.X - end.X) > 0.01 || Math.Abs(start.Y - end.Y) > 0.01)
                {
                    context.DrawLine(snapPen, start, end);
                }
            }

            if (this.thisDraftRectangle.HasValue)
            {
                var localDraftRect = RectGeometry.PixelToLocalRect(this.thisDraftRectangle.Value, contentRect, this.thisBitmapPixelSize);
                var draftBorderRect = RectGeometry.InsetRectForStroke(localDraftRect, borderThickness);

                context.DrawRectangle(draftFillBrush, null, localDraftRect);
                context.DrawRectangle(null, draftPen, draftBorderRect);
            }
        }

        // ###########################################################################################
        // Draws compact square marker segments at the 4 corners and 4 side centers of the selected
        // rectangle. On very small rectangles, side markers are reduced or suppressed so they do
        // not overlap the corner markers and imply the wrong resize behavior.
        // ###########################################################################################
        // Draws the resize affordance on the selected highlight.
        //
        // The marker layout comes from SelectionMarkerGeometry, shared with the worklog area
        // overlay - the two are meant to look identical, and sharing the maths is what makes that
        // true rather than a comment claiming it.
        // ###########################################################################################
        private void DrawSelectionMarkers(DrawingContext context, Rect rect, double scale)
        {
            var markerBrush = new SolidColorBrush(this.thisHighlightColor, 1.0);

            foreach (var markerRect in SelectionMarkerGeometry.BuildSelectionMarkerRects(rect, scale))
            {
                context.DrawRectangle(markerBrush, null, markerRect);
            }
        }

        // ###########################################################################################
        // Applies all overlay state in one batch so the control only invalidates once per editor
        // refresh instead of once per property assignment.
        // ###########################################################################################
        public void ApplyState(
            IReadOnlyList<Rect>? rectangles,
            int selectedIndex,
            IReadOnlyList<int>? selectedIndices,
            Rect? selectionBounds,
            int hoveredIndex,
            Rect? draftRectangle,
            IReadOnlyList<(Point Start, Point End)>? snapGuides,
            PixelSize bitmapPixelSize,
            Matrix viewMatrix,
            Color highlightColor,
            double highlightOpacity,
            bool isVisible)
        {
            this.thisRectangles = rectangles ?? Array.Empty<Rect>();
            this.thisSelectedIndex = selectedIndex;
            this.thisSelectedIndices = selectedIndices ?? Array.Empty<int>();
            this.thisSelectionBounds = selectionBounds;
            this.thisHoveredIndex = hoveredIndex;
            this.thisDraftRectangle = draftRectangle;
            this.thisSnapGuides = snapGuides ?? Array.Empty<(Point Start, Point End)>();
            this.thisBitmapPixelSize = bitmapPixelSize;
            this.thisViewMatrix = viewMatrix;
            this.thisHighlightColor = highlightColor;
            this.thisHighlightOpacity = highlightOpacity;
            this.IsVisible = isVisible;

            this.InvalidateVisual();
        }


        



    }
}