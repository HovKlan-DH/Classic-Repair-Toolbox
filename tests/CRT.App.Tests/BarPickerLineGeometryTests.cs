using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// A line of the worklog bar's drop-downs - its words, then its badge while there is room for both
// (BarPickerLineGeometry; 2026-11-06: the closed device box read "Working" and not "C64" on a narrow window).
// ###########################################################################################
public sealed class BarPickerLineGeometryTests
{
    // The words keep at least LeastOfWords beside the badge - or all of them, when shorter.
    [Theory]
    [InlineData(200, 120, 60, true)]
    [InlineData(100, 120, 60, true)]
    [InlineData(99, 120, 60, false)]
    [InlineData(85, 25, 60, true)]
    [InlineData(84, 25, 60, false)]
    [InlineData(60, 0, 60, true)]
    [InlineData(30, 120, 60, false)]
    public void The_badge_shows_while_the_words_keep_their_least(double width, double words, double badge, bool shows)
    {
        Assert.Equal(40, BarPickerLineGeometry.LeastOfWords);
        Assert.Equal(shows, BarPickerLineGeometry.ShowsBadge(width, words, badge));
    }

    // The open list measures its widest line unbounded: every badge there shows.
    [Fact]
    public void Unbounded_every_badge_shows()
    {
        Assert.True(BarPickerLineGeometry.ShowsBadge(double.PositiveInfinity, 500, 80));
    }

    [Theory]
    [InlineData(200, 60, true, 140)]
    [InlineData(200, 60, false, 200)]
    [InlineData(40, 60, true, 0)]
    public void The_words_get_the_line_less_the_badge_while_it_shows(double width, double badge, bool shows, double expected)
    {
        Assert.Equal(expected, BarPickerLineGeometry.WordsWidth(width, badge, shows));
    }

    // ###########################################################################################
    // The device badge a workbook's line opens with (owner request, 2026-11-07: "show the device name first in a
    // badge") takes its own width up to LeadingShare of the line, so a long device name never leaves the workbook's
    // own words no room; unbounded, its whole width.
    // ###########################################################################################
    [Theory]
    [InlineData(300, 60, 60)]
    [InlineData(300, 200, 120)]
    [InlineData(0, 60, 0)]
    [InlineData(double.PositiveInfinity, 200, 200)]
    public void The_badge_before_the_words_takes_at_most_its_share_of_the_line(double width, double leading, double expected)
    {
        Assert.Equal(0.4, BarPickerLineGeometry.LeadingShare);
        Assert.Equal(expected, BarPickerLineGeometry.LeadingWidth(width, leading), 6);
    }
}
