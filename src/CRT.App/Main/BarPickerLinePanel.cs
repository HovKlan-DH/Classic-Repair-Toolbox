using Avalonia;
using Avalonia.Controls;
using Handlers.Geometry;
using System;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A line of the worklog bar's two drop-downs (Main.Devices.cs): its words, cut short with "...", then its
    // badges side by side at the line's end - a device's status and its count of workbooks (owner request,
    // 2026-11-06: the count "as the last, right of the status") - while there is room for them
    // (BarPickerLineGeometry); with too little, the words alone, the badges laid out past the line's edge and
    // clipped away.
    //
    // A line can open with a LEADING badge too - the device a workbook is on, while the "Workbook" drop-down lists
    // every device's (owner request, 2026-11-07: "show the device name first in a badge"): its own width, up to
    // BarPickerLineGeometry.LeadingShare of the line, cut short beyond that. With one, it is the first child and the
    // words the second; without, the words are the first. Every child after the words is a badge at the end.
    //
    // Toggling a badge's IsVisible from here would ask for another layout from inside this one; laid out
    // outside the clip it is simply not seen.
    // ###########################################################################################
    internal sealed class BarPickerLinePanel : Panel
    {
        private readonly bool thisHasLeading;

        public BarPickerLinePanel(bool hasLeading = false)
        {
            this.thisHasLeading = hasLeading;
            this.ClipToBounds = true;
        }

        private int WordsIndex => this.thisHasLeading ? 1 : 0;

        private Control? Leading => this.thisHasLeading && this.Children.Count > 0 ? this.Children[0] : null;

        private Control? Words => this.Children.Count > this.WordsIndex ? this.Children[this.WordsIndex] : null;

        private System.Collections.Generic.IEnumerable<Control> Badges => this.Children.Skip(this.WordsIndex + 1);

        // Whether the badges were laid out in view at the last arrange - for a test.
        internal bool BadgeShown { get; private set; }

        // The leading badge's width at the last arrange - for a test.
        internal double LeadingWidthShown { get; private set; }

        // The words' own width, unbounded - what the badges are weighed against at arrange, where nothing may be
        // measured again (a child's desired size changing there asks for another layout, and another).
        private double thisWordsNaturalWidth;

        // The leading badge's own width, unbounded - what its cap is taken from.
        private double thisLeadingNaturalWidth;

        private double BadgesWidth => this.Badges.Sum(badge => badge.DesiredSize.Width);

        private bool HasBadges => this.Children.Count > this.WordsIndex + 1;

        protected override Size MeasureOverride(Size availableSize)
        {
            double badgesHeight = 0;

            foreach (Control badge in this.Badges)
            {
                badge.Measure(Size.Infinity);
                badgesHeight = Math.Max(badgesHeight, badge.DesiredSize.Height);
            }

            double badgesWidth = this.BadgesWidth;

            // The leading badge first: its share of the line comes off what the words and the end badges have.
            double leadingWidth = 0;
            double leadingHeight = 0;

            if (this.Leading is { } leading)
            {
                leading.Measure(Size.Infinity);
                this.thisLeadingNaturalWidth = leading.DesiredSize.Width;
                leadingWidth = BarPickerLineGeometry.LeadingWidth(availableSize.Width, this.thisLeadingNaturalWidth);

                if (leadingWidth < this.thisLeadingNaturalWidth)
                {
                    leading.Measure(new Size(leadingWidth, availableSize.Height));
                }

                leadingHeight = leading.DesiredSize.Height;
            }

            double rest = double.IsPositiveInfinity(availableSize.Width) ? availableSize.Width : Math.Max(0, availableSize.Width - leadingWidth);

            Control? words = this.Words;
            words?.Measure(Size.Infinity);
            this.thisWordsNaturalWidth = words?.DesiredSize.Width ?? 0;

            bool showsBadges = this.HasBadges && BarPickerLineGeometry.ShowsBadge(rest, this.thisWordsNaturalWidth, badgesWidth);

            if (words != null && !double.IsPositiveInfinity(rest))
            {
                words.Measure(new Size(BarPickerLineGeometry.WordsWidth(rest, badgesWidth, showsBadges), availableSize.Height));
            }

            double width = leadingWidth + (words?.DesiredSize.Width ?? 0) + (showsBadges ? badgesWidth : 0);
            double height = Math.Max(Math.Max(words?.DesiredSize.Height ?? 0, showsBadges ? badgesHeight : 0), leadingHeight);

            return new Size(width, height);
        }

        protected override Size ArrangeOverride(Size finalSize)
        {
            double leadingWidth = this.Leading is null ? 0 : BarPickerLineGeometry.LeadingWidth(finalSize.Width, this.thisLeadingNaturalWidth);
            this.Leading?.Arrange(new Rect(0, 0, leadingWidth, finalSize.Height));
            this.LeadingWidthShown = leadingWidth;

            double rest = Math.Max(0, finalSize.Width - leadingWidth);
            double badgesWidth = this.BadgesWidth;

            bool showsBadges = this.HasBadges && BarPickerLineGeometry.ShowsBadge(rest, this.thisWordsNaturalWidth, badgesWidth);
            double wordsWidth = BarPickerLineGeometry.WordsWidth(rest, badgesWidth, showsBadges);

            // A TextBlock arranged narrower than it was measured lays its text out again at that width, so its
            // "..." falls where it is cut.
            this.Words?.Arrange(new Rect(leadingWidth, 0, wordsWidth, finalSize.Height));

            double x = showsBadges ? finalSize.Width - badgesWidth : finalSize.Width;

            foreach (Control badge in this.Badges)
            {
                badge.Arrange(new Rect(x, 0, badge.DesiredSize.Width, finalSize.Height));
                x += badge.DesiredSize.Width;
            }

            this.BadgeShown = showsBadges;

            return finalSize;
        }
    }
}
