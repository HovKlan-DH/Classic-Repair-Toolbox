using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // THE STATISTICS VIEW'S GRAPH OF VIEWS PER DAY (owner request, 2026-10-09: "please create a graph
    // for showing usage of board per day. Hovering a day with mouse should show day and number of
    // usage") - every number and rectangle of it, so BoardViewsChart only draws.
    //
    // One bar per UTC day, the server's days (BoardViewStatistics.Daily lists only days WITH a view;
    // the rest are drawn at 0, never left out, or the time axis would lie). Bars are at most 24px wide
    // with a 2px gap where the room allows; the scale's top is a round number with at most four
    // steps under it; the date labels along the bottom never touch. A day is found by its whole
    // column, not only its bar, so a day with no views can be pointed at too.
    // ###########################################################################################
    internal static class ViewsChartGeometry
    {
        public const double MaxBarWidth = 24;

        public const double BarGap = 2;

        // The ranges the view offers, in days - the first is where it starts.
        public static readonly IReadOnlyList<int> Ranges = [30, 90, 365];

        // ###########################################################################################
        // The last `count` UTC days up to and including `today`, oldest first, each with its views -
        // 0 for a day the server did not list.
        // ###########################################################################################
        public static IReadOnlyList<BoardViewDay> Days(IEnumerable<BoardViewDay>? daily, DateOnly today, int count)
        {
            var views = new Dictionary<DateOnly, int>();

            foreach (BoardViewDay day in daily ?? [])
                views[day.Day] = views.GetValueOrDefault(day.Day) + Math.Max(0, day.Views);

            return Enumerable.Range(0, Math.Max(0, count))
                .Select(index => today.AddDays(index - (count - 1)))
                .Select(day => new BoardViewDay(day, views.GetValueOrDefault(day)))
                .ToList();
        }

        // ###########################################################################################
        // The scale: its top and the round numbers on it, from 0 - a step of 1, 2 or 5 times a power
        // of ten, the smallest that reaches `max` in four steps. With no views at all, 0 to 1.
        // ###########################################################################################
        public static (int Top, IReadOnlyList<int> Ticks) Scale(int max)
        {
            if (max <= 0)
                return (1, [0, 1]);

            long step = ViewsChartGeometry.RoundSteps().First(candidate => candidate * 4 >= max);
            int top = (int)(((max + step - 1) / step) * step);

            return (top, Enumerable.Range(0, (int)(top / step) + 1).Select(index => (int)(index * step)).ToList());
        }

        // 1, 2, 5, 10, 20, 50, 100, ...
        private static IEnumerable<long> RoundSteps()
        {
            for (long power = 1; power <= long.MaxValue / 50; power *= 10)
            {
                yield return power;
                yield return 2 * power;
                yield return 5 * power;
            }
        }

        // The width of one day's column.
        public static double SlotWidth(int count, Rect plot) => count <= 0 ? 0 : plot.Width / count;

        // ###########################################################################################
        // Each day's bar, standing on the plot's bottom edge - centred in its column, at most
        // MaxBarWidth wide, with BarGap of air beside it while the column is wide enough to spare it.
        // A day with no views has a bar of no height.
        // ###########################################################################################
        public static IReadOnlyList<Rect> Bars(IReadOnlyList<int> values, Rect plot, int top)
        {
            ArgumentNullException.ThrowIfNull(values);

            double slot = ViewsChartGeometry.SlotWidth(values.Count, plot);
            double gap = slot >= 3 * ViewsChartGeometry.BarGap ? ViewsChartGeometry.BarGap : slot >= 2 ? 1 : 0;
            double width = Math.Max(Math.Min(slot, 1), Math.Min(ViewsChartGeometry.MaxBarWidth, slot - gap));

            return values
                .Select((value, index) =>
                {
                    double height = top <= 0 ? 0 : Math.Clamp(value / (double)top, 0, 1) * plot.Height;
                    double x = plot.X + (index * slot) + ((slot - width) / 2);

                    return new Rect(x, plot.Bottom - height, width, height);
                })
                .ToList();
        }

        // The day whose column `x` is in, or -1 outside the plot.
        public static int DayAt(double x, int count, Rect plot)
        {
            if (count <= 0 || x < plot.X || x >= plot.Right)
                return -1;

            return Math.Clamp((int)((x - plot.X) / ViewsChartGeometry.SlotWidth(count, plot)), 0, count - 1);
        }

        // ###########################################################################################
        // The days whose dates are written under the axis: the first and the last, and as many evenly
        // between as fit with `labelWidth` each and some air - never two labels touching.
        // ###########################################################################################
        public static IReadOnlyList<int> LabelledDays(int count, double plotWidth, double labelWidth)
        {
            if (count <= 0)
                return [];

            if (count == 1)
                return [0];

            int fit = (int)Math.Floor(plotWidth / Math.Max(1, labelWidth + 16));
            int labels = Math.Clamp(fit, 2, Math.Min(count, 7));

            return Enumerable.Range(0, labels)
                .Select(index => (int)Math.Round(index * (count - 1) / (double)(labels - 1)))
                .Distinct()
                .ToList();
        }

        // ###########################################################################################
        // The words a day is shown with, as the rest of CRT writes dates: "2026-October-9 (Friday) -
        // 7 views". UTC days, so no time zone moves them.
        // ###########################################################################################
        public static string DayText(BoardViewDay day) =>
            $"{day.Day.ToString("yyyy-MMMM-d", CultureInfo.InvariantCulture)} ({day.Day.DayOfWeek}) - " +
            (day.Views == 1 ? "1 view" : $"{day.Views.ToString("N0", CultureInfo.InvariantCulture)} views");

        // A date under the axis: "9 Oct".
        public static string AxisDate(DateOnly day) => day.ToString("d MMM", CultureInfo.InvariantCulture);
    }
}
