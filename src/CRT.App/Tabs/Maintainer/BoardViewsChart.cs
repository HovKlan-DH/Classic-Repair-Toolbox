using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE GRAPH OF VIEWS PER DAY on a board's Statistics view (owner request, 2026-10-09: "please
    // create a graph for showing usage of board per day. Hovering a day with mouse should show day
    // and number of usage. I am not sure what kind of graphing library can be used here?").
    //
    // Drawn here rather than with a charting library (owner's choice, the same day): one series of
    // bars needs no dependency that pins its own Avalonia and SkiaSharp, and it takes CRT's own
    // colours in both themes (Chart_* in App.axaml). Every number and rectangle is
    // ViewsChartGeometry's, tested; this only draws them and follows the pointer.
    //
    // Pointing at a day - anywhere in its column, so a day with no views too - colours its bar and
    // writes the day and its views above the plot. The readout is drawn, not a tooltip: CRT opens
    // every tooltip at its control's edge (ToolTipPlacement), and a tooltip there would sit far from
    // a bar in the middle of a wide graph.
    // ###########################################################################################
    internal sealed class BoardViewsChart : Control
    {
        private const double LeftMargin = 40;
        private const double ReadoutHeight = 22;
        private const double AxisHeight = 20;
        private const double FontSize = 11;

        private IReadOnlyList<BoardViewDay> thisDays = [];
        private int thisHovered = -1;

        public BoardViewsChart()
        {
            this.Height = 200;
            this.ClipToBounds = true;
            this.ActualThemeVariantChanged += (_, _) => this.InvalidateVisual();
        }

        // The days drawn, oldest first - ViewsChartGeometry.Days.
        public IReadOnlyList<BoardViewDay> Days
        {
            get => this.thisDays;
            set
            {
                this.thisDays = value ?? [];
                this.thisHovered = -1;
                this.InvalidateVisual();
            }
        }

        // The day pointed at, or null - and what the readout says about it.
        internal BoardViewDay? HoveredDay => this.thisHovered >= 0 && this.thisHovered < this.thisDays.Count ? this.thisDays[this.thisHovered] : null;

        internal string? ReadoutText => this.HoveredDay is BoardViewDay day ? ViewsChartGeometry.DayText(day) : null;

        // Where the bars stand - for the pointer, and for a test placing it.
        internal Rect Plot =>
            new(
                BoardViewsChart.LeftMargin,
                BoardViewsChart.ReadoutHeight,
                Math.Max(0, this.Bounds.Width - BoardViewsChart.LeftMargin - 4),
                Math.Max(0, this.Bounds.Height - BoardViewsChart.ReadoutHeight - BoardViewsChart.AxisHeight));

        protected override void OnPointerMoved(PointerEventArgs e)
        {
            base.OnPointerMoved(e);
            this.PointAt(e.GetPosition(this).X);
        }

        protected override void OnPointerExited(PointerEventArgs e)
        {
            base.OnPointerExited(e);
            this.PointAt(double.NaN);
        }

        // The day whose column `x` is in - none for NaN or outside the plot.
        internal void PointAt(double x)
        {
            int day = double.IsNaN(x) ? -1 : ViewsChartGeometry.DayAt(x, this.thisDays.Count, this.Plot);

            if (day != this.thisHovered)
            {
                this.thisHovered = day;
                this.InvalidateVisual();
            }
        }

        public override void Render(DrawingContext context)
        {
            base.Render(context);

            Rect plot = this.Plot;

            if (this.thisDays.Count == 0 || plot.Width <= 0 || plot.Height <= 0)
                return;

            IBrush bar = ThemeResources.ResolveBrush("Chart_Bar", Brushes.SteelBlue);
            IBrush hover = ThemeResources.ResolveBrush("Chart_BarHover", Brushes.Navy);
            IBrush grid = ThemeResources.ResolveBrush("Chart_Grid", Brushes.Gainsboro);
            IBrush baseline = ThemeResources.ResolveBrush("Chart_Baseline", Brushes.Silver);
            IBrush label = ThemeResources.ResolveBrush("Chart_Label", Brushes.Gray);
            IBrush ink = ThemeResources.ResolveBrush("Fg", Brushes.Black);

            (int top, IReadOnlyList<int> ticks) = ViewsChartGeometry.Scale(this.thisDays.Max(day => day.Views));

            // The scale: hairlines across, their numbers on the left, the baseline a step darker.
            foreach (int tick in ticks)
            {
                double y = plot.Bottom - (tick / (double)top * plot.Height);
                var pen = new Pen(tick == 0 ? baseline : grid, 1);

                context.DrawLine(pen, new Point(plot.X, Math.Round(y) + 0.5), new Point(plot.Right, Math.Round(y) + 0.5));

                FormattedText number = this.Text(tick.ToString("N0", CultureInfo.InvariantCulture), label);
                context.DrawText(number, new Point(plot.X - 6 - number.Width, y - (number.Height / 2)));
            }

            // The bars - a rounded top on a bar wide and tall enough to show one, square on the
            // baseline.
            IReadOnlyList<Rect> bars = ViewsChartGeometry.Bars(this.thisDays.Select(day => day.Views).ToList(), plot, top);

            for (int index = 0; index < bars.Count; index++)
            {
                Rect rect = bars[index];

                if (rect.Height <= 0)
                    continue;

                IBrush fill = index == this.thisHovered ? hover : bar;
                double radius = Math.Min(4, Math.Min(rect.Width / 2, rect.Height));

                if (radius >= 1)
                    context.DrawGeometry(fill, null, BoardViewsChart.TopRounded(rect, radius));
                else
                    context.DrawRectangle(fill, null, rect);
            }

            // The day pointed at with no views has no bar: a mark on the baseline shows which it is.
            if (this.HoveredDay is { Views: 0 } && this.thisHovered < bars.Count)
            {
                Rect empty = bars[this.thisHovered];
                context.DrawRectangle(hover, null, new Rect(empty.X, plot.Bottom - 2, Math.Max(1, empty.Width), 2));
            }

            // The dates along the bottom, never touching.
            double dateWidth = this.Text("30 Sep", label).Width;
            double slot = ViewsChartGeometry.SlotWidth(this.thisDays.Count, plot);

            foreach (int index in ViewsChartGeometry.LabelledDays(this.thisDays.Count, plot.Width, dateWidth))
            {
                FormattedText date = this.Text(ViewsChartGeometry.AxisDate(this.thisDays[index].Day), label);
                double centre = plot.X + (index * slot) + (slot / 2);
                double x = Math.Clamp(centre - (date.Width / 2), plot.X, Math.Max(plot.X, plot.Right - date.Width));

                context.DrawText(date, new Point(x, plot.Bottom + 4));
            }

            // The readout: the day pointed at and its views, above its column, kept inside the graph.
            if (this.ReadoutText is string readout && this.thisHovered < bars.Count)
            {
                FormattedText text = this.Text(readout, ink);
                double centre = bars[this.thisHovered].Center.X;
                double x = Math.Clamp(centre - (text.Width / 2), 0, Math.Max(0, this.Bounds.Width - text.Width));

                context.DrawText(text, new Point(x, (BoardViewsChart.ReadoutHeight - text.Height) / 2));
            }
        }

        private FormattedText Text(string text, IBrush brush) =>
            new(
                text,
                CultureInfo.InvariantCulture,
                FlowDirection.LeftToRight,
                new Typeface(TextElement.GetFontFamily(this)),
                BoardViewsChart.FontSize,
                brush);

        // A bar with its two top corners rounded and its foot square on the baseline.
        private static StreamGeometry TopRounded(Rect rect, double radius)
        {
            var geometry = new StreamGeometry();

            using StreamGeometryContext path = geometry.Open();

            path.BeginFigure(new Point(rect.Left, rect.Bottom), isFilled: true);
            path.LineTo(new Point(rect.Left, rect.Top + radius));
            path.ArcTo(new Point(rect.Left + radius, rect.Top), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
            path.LineTo(new Point(rect.Right - radius, rect.Top));
            path.ArcTo(new Point(rect.Right, rect.Top + radius), new Size(radius, radius), 0, false, SweepDirection.Clockwise);
            path.LineTo(new Point(rect.Right, rect.Bottom));
            path.EndFigure(isClosed: true);

            return geometry;
        }
    }
}
