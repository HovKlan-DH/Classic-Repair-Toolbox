using Avalonia;
using Avalonia.Controls;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The worklog bar's "Device" and "Workbook" drop-downs with their labels, side by side: each label as wide as
    // its words, each drop-down as wide as the line it shows - and short of room, the room shared as
    // BarPickersGeometry says (owner request, 2026-11-06: "have the two drop-downs resize, so it will show as much
    // as the text in the drop-downs as possible"). A ComboBox child is a drop-down; any other child keeps its own
    // width. Clipped to its place: on CRT's narrowest window the labels alone may not fit beside the buttons.
    // ###########################################################################################
    internal sealed class BarPickersPanel : Panel
    {
        // The drop-downs' widths the last measure worked out, in their order - what arrange lays them out at.
        private double[] thisWidths = [];

        public BarPickersPanel()
        {
            this.ClipToBounds = true;
        }

        protected override Size MeasureOverride(Size availableSize)
        {
            List<Control> shown = this.Children.Where(child => child.IsVisible).ToList();
            var wanted = new List<double>();
            double fixedWidth = 0;
            double height = 0;

            foreach (Control child in shown)
            {
                child.Measure(Size.Infinity);
                height = Math.Max(height, child.DesiredSize.Height);

                if (child is ComboBox)
                {
                    wanted.Add(child.DesiredSize.Width);
                }
                else
                {
                    fixedWidth += child.DesiredSize.Width;
                }
            }

            double room = double.IsPositiveInfinity(availableSize.Width) ? wanted.Sum() : availableSize.Width - fixedWidth;
            this.thisWidths = BarPickersGeometry.Widths(room, wanted);

            int box = 0;
            foreach (Control child in shown.Where(child => child is ComboBox))
            {
                // Measured at the width it gets, so its words are cut short there.
                child.Measure(new Size(this.thisWidths[box++], availableSize.Height));
            }

            return new Size(fixedWidth + this.thisWidths.Sum(), height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            double x = 0;
            int box = 0;

            foreach (Control child in this.Children.Where(child => child.IsVisible))
            {
                double width = child is ComboBox && box < this.thisWidths.Length ? this.thisWidths[box++] : child.DesiredSize.Width;
                child.Arrange(new Rect(x, 0, width, finalSize.Height));
                x += width;
            }

            return finalSize;
        }
    }
}
