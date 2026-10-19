using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// FeedbackClickGeometry - which row of the Feedback tab's field a click on its empty space is on
// (code review, 2026-10-16: a click beside a picture in the middle sent the caret to the end).
// The rows below: a part from 0 to 20, a picture from 30 to 130, a part from 140 to 160.
// ###########################################################################################
public sealed class FeedbackClickGeometryTests
{
    private static readonly IReadOnlyList<(double Top, double Bottom)> Rows = [(0, 20), (30, 130), (140, 160)];

    [Theory]
    [InlineData(10, 0)]
    [InlineData(20, 0)]
    [InlineData(80, 1)]
    [InlineData(130, 1)]
    [InlineData(150, 2)]
    public void A_click_within_a_row_is_on_it(double y, int row)
    {
        Assert.Equal(row, FeedbackClickGeometry.RowAt(y, Rows));
    }

    // In the gap between two rows - a picture's margin - the nearer one; the upper on a tie.
    [Theory]
    [InlineData(22, 0)]
    [InlineData(25, 0)]
    [InlineData(28, 1)]
    [InlineData(133, 1)]
    [InlineData(138, 2)]
    public void A_click_in_a_gap_is_on_the_nearer_row(double y, int row)
    {
        Assert.Equal(row, FeedbackClickGeometry.RowAt(y, Rows));
    }

    [Fact]
    public void Above_every_row_is_the_first_and_below_them_the_last()
    {
        Assert.Equal(0, FeedbackClickGeometry.RowAt(-5, Rows));
        Assert.Equal(2, FeedbackClickGeometry.RowAt(900, Rows));
        Assert.Equal(-1, FeedbackClickGeometry.RowAt(10, []));
    }
}
