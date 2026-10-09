using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace ClassicRepairToolbox.Tests.Ui.Maintainer;

// ###########################################################################################
// The graph of views per day on a board's Statistics view (owner request, 2026-10-09: "please
// create a graph for showing usage of board per day. Hovering a day with mouse should show day and
// number of usage") - BoardViewsChart in BoardDetailView, shown in a real window so its plot has a
// size. The numbers and rectangles are ViewsChartGeometry's, pinned in ViewsChartGeometryTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class BoardViewsChartTests
{
    private static readonly DateOnly Today = new(2026, 10, 9);

    private static BoardDetailAnswer Detail(IReadOnlyList<BoardViewDay>? daily) =>
        new(
            new BoardOverviewEntry("Commodore/C64/250407", "Commodore", "C64", "250407", true, true, false, true, null, null, null, 1),
            [], [], [],
            Views: new BoardViewStatistics(12, 40, 300, 2, [new("DK", "Denmark", 200)], daily));

    private static (Window Window, BoardDetailView View) Shown(IReadOnlyList<BoardViewDay>? daily)
    {
        var view = new BoardDetailView { TodayForTests = BoardViewsChartTests.Today };
        var window = new Window { Content = view, Width = 1100, Height = 900 };
        window.Show();

        view.ShowDetailForTests(BoardViewsChartTests.Detail(daily));
        view.ShowSectionAsync(BoardSection.Statistics).GetAwaiter().GetResult();
        Dispatcher.UIThread.RunJobs();

        return (window, view);
    }

    private static BoardViewsChart? Chart(BoardDetailView view) =>
        view.GetVisualDescendants().OfType<BoardViewsChart>().SingleOrDefault();

    [Fact]
    public void The_graph_starts_on_the_last_30_days_and_the_switch_shows_90_and_365()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardDetailView view) = BoardViewsChartTests.Shown([new(BoardViewsChartTests.Today.AddDays(-2), 7)]);

            BoardViewsChart chart = BoardViewsChartTests.Chart(view)!;
            Assert.Equal(30, chart.Days.Count);
            Assert.Equal(BoardViewsChartTests.Today, chart.Days[^1].Day);
            Assert.Equal(7, chart.Days[^3].Views);
            Assert.Contains("Selected", view.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "ViewsRange30Button").Classes);

            Button ninety = view.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "ViewsRange90Button");
            ninety.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Assert.Equal(90, chart.Days.Count);
            Assert.Contains("Selected", ninety.Classes);

            view.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "ViewsRange365Button")
                .RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
            Assert.Equal(365, chart.Days.Count);
            Assert.DoesNotContain("Selected", ninety.Classes);

            window.Close();
        });
    }

    // Pointing at a day - its bar, or the empty column of a day nobody looked at - says the day and
    // its views; leaving the graph says nothing.
    [Fact]
    public void Pointing_at_a_day_says_its_date_and_views_and_leaving_says_nothing()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardDetailView view) = BoardViewsChartTests.Shown([new(BoardViewsChartTests.Today.AddDays(-2), 7)]);
            BoardViewsChart chart = BoardViewsChartTests.Chart(view)!;
            Rect plot = chart.Plot;
            double slot = plot.Width / 30;

            Assert.True(plot.Width > 0 && plot.Height > 0, $"plot [{plot}]");
            Assert.Null(chart.ReadoutText);

            chart.PointAt(plot.X + (27 * slot) + (slot / 2));
            Assert.Equal("2026-October-7 (Wednesday) - 7 views", chart.ReadoutText);

            chart.PointAt(plot.X + (29 * slot) + 1);
            Assert.Equal("2026-October-9 (Friday) - 0 views", chart.ReadoutText);

            chart.PointAt(double.NaN);
            Assert.Null(chart.ReadoutText);

            window.Close();
        });
    }

    // A server older than the graph sends no days: the counts are shown, and no graph.
    [Fact]
    public void A_server_that_sends_no_days_gets_no_graph()
    {
        UiTest.Run(() =>
        {
            (Window window, BoardDetailView view) = BoardViewsChartTests.Shown(null);

            Assert.Null(BoardViewsChartTests.Chart(view));
            Assert.DoesNotContain(BoardsDisplay.ViewsPerDayHeading, view.SectionTextsForTests("ViewsSection"));

            window.Close();
        });
    }
}
