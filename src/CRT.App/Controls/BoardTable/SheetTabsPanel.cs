using Avalonia;
using Avalonia.Controls;
using Handlers.Geometry;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The panel the table's sheet tabs stand in (owner request, 2026-11-06 - Avalonia's TabControl wraps its
    // items as words in a paragraph): a line per group that asks for one, the line headings in one column, a
    // long line wrapped under its own tabs. Every position is SheetTabsLayout's; this only measures the tabs
    // and puts them where it says. A heading is a TabItem of the class "SheetGroup", and one starting a line
    // has "StartsLine" too (BoardTableEditor.BuildSheetTabs).
    // ###########################################################################################
    public sealed class SheetTabsPanel : Panel
    {
        private List<Control> Shown() => this.Children.Where(child => child.IsVisible).ToList();

        private static SheetTabsItem ItemOf(Control child) => new(
            child.DesiredSize.Width,
            child.DesiredSize.Height,
            child.Classes.Contains("StartsLine"),
            child.Classes.Contains("SheetGroup"));

        protected override Size MeasureOverride(Size availableSize)
        {
            foreach (Control child in this.Children)
            {
                child.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
            }

            SheetTabsLayout.Arrange(this.Shown().Select(SheetTabsPanel.ItemOf).ToList(), availableSize.Width, out Size size);

            return size;
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            List<Control> shown = this.Shown();
            IReadOnlyList<Rect> rects = SheetTabsLayout.Arrange(shown.Select(SheetTabsPanel.ItemOf).ToList(), finalSize.Width, out _);

            for (int i = 0; i < shown.Count; i++)
            {
                shown[i].Arrange(rects[i]);
            }

            return finalSize;
        }
    }
}
