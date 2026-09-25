using Avalonia;
using Avalonia.Controls;
using Handlers.Geometry;
using System;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // The thin rim that lays out thumbnail tiles from the pure auto-fit maths in
    // Handlers/Geometry/ThumbnailGalleryGeometry.cs - the same "resolve to plain values, call the
    // geometry class" pattern TabSchematics.LabelEditor.Snap.cs uses for LabelEditorSnapGeometry.
    //
    // Every child is measured/arranged at the SAME computed cell size (see
    // ThumbnailGalleryGeometry's own header for why), which a stock WrapPanel cannot do - it sizes
    // each child by its own desired size or a fixed ItemWidth/ItemHeight, not by "all N children
    // share one size that fills the available area".
    // ###########################################################################################
    internal sealed class ThumbnailGalleryPanel : Panel
    {
        public static readonly StyledProperty<double> SpacingProperty =
            AvaloniaProperty.Register<ThumbnailGalleryPanel, double>(nameof(Spacing), 8.0);

        public static readonly StyledProperty<double> MarginPaddingProperty =
            AvaloniaProperty.Register<ThumbnailGalleryPanel, double>(nameof(MarginPadding), 8.0);

        public double Spacing
        {
            get => this.GetValue(SpacingProperty);
            set => this.SetValue(SpacingProperty, value);
        }

        public double MarginPadding
        {
            get => this.GetValue(MarginPaddingProperty);
            set => this.SetValue(MarginPaddingProperty, value);
        }

        // ###########################################################################################
        // Returns the layout the panel would currently use for its children - exposed so the
        // gallery's drag-reorder logic can find the nearest cell to the pointer without
        // recomputing the same maths a second time with different rounding.
        // ###########################################################################################
        public (ThumbnailGalleryGeometry.GridShape Shape, Size CellSize, System.Collections.Generic.IReadOnlyList<Point> Positions) CurrentLayout { get; private set; }

        protected override Size MeasureOverride(Size availableSize)
        {
            var layout = ThumbnailGalleryGeometry.BuildLayout(this.Children.Count, availableSize, this.Spacing, this.MarginPadding);
            this.CurrentLayout = layout;

            for (int index = 0; index < this.Children.Count; index++)
            {
                this.Children[index].Measure(layout.CellSize);
            }

            // An unconstrained axis has no "available" size to hand back, so the grid's own measured
            // extent is returned for it instead. ThumbnailGalleryGeometry.ComputeCellSize resolves
            // an infinite axis to a fixed fallback cell rather than an infinite one, so this
            // arithmetic stays finite - returning an Infinity desired size is a layout error in
            // Avalonia, and every child would have been measured against an infinite cell too.
            bool isWidthUnconstrained = double.IsInfinity(availableSize.Width) || double.IsNaN(availableSize.Width);
            bool isHeightUnconstrained = double.IsInfinity(availableSize.Height) || double.IsNaN(availableSize.Height);

            if (isWidthUnconstrained || isHeightUnconstrained)
            {
                double totalWidth = (this.MarginPadding * 2) + (layout.Shape.Columns * layout.CellSize.Width) + (Math.Max(0, layout.Shape.Columns - 1) * this.Spacing);
                double totalHeight = (this.MarginPadding * 2) + (layout.Shape.Rows * layout.CellSize.Height) + (Math.Max(0, layout.Shape.Rows - 1) * this.Spacing);

                return new Size(
                    isWidthUnconstrained ? totalWidth : availableSize.Width,
                    isHeightUnconstrained ? totalHeight : availableSize.Height);
            }

            return availableSize;
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            var layout = ThumbnailGalleryGeometry.BuildLayout(this.Children.Count, finalSize, this.Spacing, this.MarginPadding);
            this.CurrentLayout = layout;

            for (int index = 0; index < this.Children.Count && index < layout.Positions.Count; index++)
            {
                this.Children[index].Arrange(new Rect(layout.Positions[index], layout.CellSize));
            }

            return finalSize;
        }
    }
}
