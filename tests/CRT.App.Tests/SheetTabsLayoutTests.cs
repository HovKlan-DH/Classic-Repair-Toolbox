using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SheetTabsLayout - where the Maintainer tab's sheet tabs stand (owner request, 2026-11-06: "if there are
// too many, it wraps into next line, which makes it hard to overview ... just have the 'Components' part on a
// new-line"): a line per group that asks for one, the line headings in one column so the tabs line up, a
// long line wrapped under its own tabs, and a heading inside a line never left at a line's end alone.
// ###########################################################################################
public sealed class SheetTabsLayoutTests
{
    private static SheetTabsItem Heading(double width, bool startsLine = false) => new(width, 30, startsLine, true);

    private static SheetTabsItem Tab(double width) => new(width, 30, false, false);

    // Board's four tabs and KiCad's one on the first line, Components' seven on the second.
    private static List<SheetTabsItem> Strip() =>
    [
        Heading(60, startsLine: true), Tab(130), Tab(50), Tab(50), Tab(70),
        Heading(70), Tab(130),
        Heading(110, startsLine: true), Tab(70), Tab(140), Tab(70), Tab(80), Tab(170), Tab(50), Tab(50),
    ];

    [Fact]
    public void Components_has_a_line_of_its_own_and_both_line_headings_share_one_column()
    {
        IReadOnlyList<Rect> rects = SheetTabsLayout.Arrange(Strip(), 2000, out Size size);

        // The first line: Board, its tabs, KiCad and its tab.
        Assert.All(rects.Take(7), rect => Assert.Equal(0, rect.Y));
        Assert.Equal(0, rects[0].X);

        // The second line: Components at the left edge, its tabs after the column - as wide as its heading,
        // the wider - and Board's first tab starts at the very same place.
        Assert.All(rects.Skip(7), rect => Assert.Equal(30, rect.Y));
        Assert.Equal(0, rects[7].X);
        Assert.Equal(110, rects[8].X);
        Assert.Equal(110, rects[1].X);

        Assert.Equal(60, size.Height);
    }

    [Fact]
    public void Only_the_components_group_starts_a_line_of_its_own()
    {
        Assert.False(BoardTableTabs.StartsLine(BoardTableTabs.BoardGroup));
        Assert.False(BoardTableTabs.StartsLine(BoardTableTabs.KiCadGroup));
        Assert.True(BoardTableTabs.StartsLine(BoardTableTabs.ComponentsGroup));
    }

    // Too narrow for a line: the rest wraps under the line's own tabs, never back under the headings.
    [Fact]
    public void A_line_too_long_for_the_width_wraps_under_its_own_tabs()
    {
        IReadOnlyList<Rect> rects = SheetTabsLayout.Arrange(Strip(), 600, out Size size);

        IEnumerable<Rect> components = rects.Skip(8);

        Assert.All(components, rect => Assert.True(rect.X >= 110, $"A tab at {rect.X} stands under the headings."));
        Assert.All(rects, rect => Assert.True(rect.Right <= 600, $"A tab ends at {rect.Right}, past the width."));
        Assert.True(size.Height > 60, "The components' line did not wrap.");
    }

    // KiCad's heading needs its tab beside it: where only the heading would fit, both go to the next line.
    [Fact]
    public void A_heading_inside_a_line_goes_along_with_its_first_tab()
    {
        // Board's tabs end at 410 (after the 110 column); KiCad's heading (70) would fit within 520, its tab
        // (130) not - so the heading goes to the next line with it, under the line's tabs.
        IReadOnlyList<Rect> rects = SheetTabsLayout.Arrange(Strip(), 520, out _);

        Assert.Equal(410, rects[4].Right);
        Assert.Equal(new Point(110, 30), rects[5].Position);
        Assert.Equal(new Point(180, 30), rects[6].Position);
    }

    [Fact]
    public void Nothing_to_lay_out_takes_no_room_and_no_width_given_is_one_line_each()
    {
        Assert.Empty(SheetTabsLayout.Arrange([], 500, out Size none));
        Assert.Equal(default, none);

        SheetTabsLayout.Arrange(Strip(), double.PositiveInfinity, out Size unbounded);
        Assert.Equal(60, unbounded.Height);
    }
}
