using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ViewsChartGeometry - the Statistics view's graph of views per day (owner request, 2026-10-09:
// "please create a graph for showing usage of board per day. Hovering a day with mouse should show
// day and number of usage"). Every number and rectangle the graph draws.
// ###########################################################################################
public sealed class ViewsChartGeometryTests
{
    private static readonly DateOnly Today = new(2026, 10, 9);

    // ###########################################################################################
    // *** EVERY DAY OF THE RANGE IS A BAR, A DAY WITH NO VIEWS AT 0. *** The server lists only days
    // with a view; leaving the others out would squeeze the time axis and lie about it.
    // ###########################################################################################
    [Fact]
    public void Every_day_of_the_range_is_there_oldest_first_and_a_day_the_server_did_not_list_is_0()
    {
        IReadOnlyList<BoardViewDay> days = ViewsChartGeometry.Days(
            [new(Today.AddDays(-2), 3), new(Today, 5), new(Today.AddDays(-40), 9)],
            Today,
            count: 4);

        Assert.Equal(
            [new(Today.AddDays(-3), 0), new(Today.AddDays(-2), 3), new(Today.AddDays(-1), 0), new(Today, 5)],
            days);
    }

    [Fact]
    public void No_daily_views_at_all_is_a_range_of_zeros()
    {
        Assert.All(ViewsChartGeometry.Days(null, Today, 30), day => Assert.Equal(0, day.Views));
        Assert.Equal(30, ViewsChartGeometry.Days([], Today, 30).Count);
    }

    // The scale's top is a round number, with at most four steps of 1, 2 or 5 times a power of ten.
    [Theory]
    [InlineData(0, 1, new[] { 0, 1 })]
    [InlineData(1, 1, new[] { 0, 1 })]
    [InlineData(3, 3, new[] { 0, 1, 2, 3 })]
    [InlineData(7, 8, new[] { 0, 2, 4, 6, 8 })]
    [InlineData(20, 20, new[] { 0, 5, 10, 15, 20 })]
    [InlineData(120, 150, new[] { 0, 50, 100, 150 })]
    [InlineData(1234, 1500, new[] { 0, 500, 1000, 1500 })]
    public void The_scale_rises_in_round_steps_to_just_above_the_busiest_day(int max, int top, int[] ticks)
    {
        (int actualTop, IReadOnlyList<int> actualTicks) = ViewsChartGeometry.Scale(max);

        Assert.Equal(top, actualTop);
        Assert.Equal(ticks, actualTicks);
    }

    // ###########################################################################################
    // Bars stand on the bottom edge, as tall as their share of the top, centred in their column -
    // at most 24px wide with 2px of air beside them while the column can spare it.
    // ###########################################################################################
    [Fact]
    public void Wide_columns_get_bars_of_at_most_24px_standing_on_the_baseline_centred()
    {
        var plot = new Rect(40, 10, 300, 100);

        IReadOnlyList<Rect> bars = ViewsChartGeometry.Bars([0, 5, 10], plot, top: 10);

        Assert.Equal(3, bars.Count);
        Assert.All(bars, bar => Assert.Equal(ViewsChartGeometry.MaxBarWidth, bar.Width));
        Assert.All(bars, bar => Assert.Equal(plot.Bottom, bar.Bottom, precision: 6));
        Assert.Equal([0d, 50d, 100d], bars.Select(bar => bar.Height));
        Assert.Equal(40 + 50 - 12, bars[0].X, precision: 6);
    }

    [Fact]
    public void Narrow_columns_keep_a_gap_while_they_can_and_never_a_bar_wider_than_its_column()
    {
        IReadOnlyList<Rect> thirty = ViewsChartGeometry.Bars(Enumerable.Repeat(1, 30).ToList(), new Rect(0, 0, 300, 100), 1);
        IReadOnlyList<Rect> year = ViewsChartGeometry.Bars(Enumerable.Repeat(1, 365).ToList(), new Rect(0, 0, 600, 100), 1);

        Assert.Equal(10 - ViewsChartGeometry.BarGap, thirty[0].Width, precision: 6);
        Assert.All(year, bar => Assert.InRange(bar.Width, 0.5, 600 / 365d));

        // Neighbours never overlap.
        Assert.All(year.Zip(year.Skip(1)), pair => Assert.True(pair.First.Right <= pair.Second.X + 1e-9));
    }

    // A day is found by its whole column - a day with no views, and no bar, too.
    [Theory]
    [InlineData(39.9, -1)]
    [InlineData(40, 0)]
    [InlineData(139.9, 0)]
    [InlineData(140, 1)]
    [InlineData(339.9, 2)]
    [InlineData(340, -1)]
    public void The_day_under_the_pointer_is_found_by_its_whole_column(double x, int day)
    {
        Assert.Equal(day, ViewsChartGeometry.DayAt(x, 3, new Rect(40, 10, 300, 100)));
    }

    // The dates along the bottom: the first and the last day always, and as many evenly between as
    // fit without touching.
    [Fact]
    public void The_dates_along_the_axis_include_both_ends_and_never_crowd()
    {
        Assert.Equal([0, 29], ViewsChartGeometry.LabelledDays(30, 120, 40));
        Assert.Equal([0, 6, 12, 17, 23, 29], ViewsChartGeometry.LabelledDays(30, 340, 40));
        Assert.Equal(7, ViewsChartGeometry.LabelledDays(365, 2000, 40).Count);
        Assert.Equal([0], ViewsChartGeometry.LabelledDays(1, 300, 40));
        Assert.Empty(ViewsChartGeometry.LabelledDays(0, 300, 40));
    }

    // Hovering a day says the day and its views, as CRT writes dates elsewhere.
    [Fact]
    public void A_day_is_said_with_its_weekday_and_its_views()
    {
        Assert.Equal("2026-October-9 (Friday) - 7 views", ViewsChartGeometry.DayText(new BoardViewDay(Today, 7)));
        Assert.Equal("2026-October-8 (Thursday) - 1 view", ViewsChartGeometry.DayText(new BoardViewDay(Today.AddDays(-1), 1)));
        Assert.Equal("2026-October-7 (Wednesday) - 1,250 views", ViewsChartGeometry.DayText(new BoardViewDay(Today.AddDays(-2), 1250)));
        Assert.Equal("9 Oct", ViewsChartGeometry.AxisDate(Today));
    }
}
