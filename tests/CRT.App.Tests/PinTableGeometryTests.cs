using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PinTableGeometry - the component popup's pin list, column by column (owner request, 2026-11-09: "align on empty spacing
// between columns - it should be the same for all, and I do like the spacing between 'Pin' and 'Signal'. Should be
// determined by longest text in each column + that spacing"). The widths below are the texts' in CRT's font, measured
// 2026-11-09: "Pin" 16, "oscilloscope" 62, "Net-(FB23-Pad1)" 89, "12.0V (+12VDC power rail)" 137.
// ###########################################################################################
public sealed class PinTableGeometryTests
{
    // Pin, Signal, Net name, Direction, Expected oscilloscope reading.
    private static readonly PinTableColumn[] Columns =
    [
        new(16, 16, Wraps: false),
        new(60, 40, Wraps: true),
        new(89, 89, Wraps: true),
        new(50, 50, Wraps: false),
        new(137, 62, Wraps: true),
    ];

    // The gap is what stood between "Pin" and "Signal" when "Pin" had a 36 wide column: the spacing the owner kept.
    [Fact]
    public void The_gap_is_the_space_that_stood_after_Pin_in_its_column_of_36()
    {
        Assert.Equal(36 - 16, PinTableGeometry.Gap);
    }

    // Every column's widest text, and one gap between each two - none after the last.
    [Fact]
    public void The_list_is_its_columns_widest_texts_and_a_gap_between_each_two()
    {
        Assert.Equal(16 + 60 + 89 + 50 + 137 + (4 * PinTableGeometry.Gap), PinTableGeometry.WidthOf(Columns));
        Assert.Equal(0, PinTableGeometry.WidthOf([]));
    }

    // A column not shown - no nets, no readings - takes neither room nor a gap.
    [Fact]
    public void A_column_not_shown_takes_no_room_and_no_gap()
    {
        PinTableColumn[] noNets = [Columns[0], Columns[1], default, Columns[3], default];

        Assert.Equal(16 + 60 + 50 + (2 * PinTableGeometry.Gap), PinTableGeometry.WidthOf(noNets));
        Assert.Equal([16, 60, 0, 50, 0], PinTableGeometry.TextWidths(noNets, 1000));
    }

    // With the room for it - or the room not known yet - every text is whole.
    [Theory]
    [InlineData(1000)]
    [InlineData(432)]
    [InlineData(0)]
    [InlineData(double.PositiveInfinity)]
    public void With_room_every_column_is_as_wide_as_its_widest_text(double room)
    {
        Assert.Equal([16, 60, 89, 50, 137], PinTableGeometry.TextWidths(Columns, room));
    }

    // ###########################################################################################
    // Short of room, the widest wrapping column narrows first: 32 short, the reading alone gives them up (137 -> 105), as
    // the signal (60) is already narrower than that and the net (89) is one word. The pin and the direction never wrap.
    // ###########################################################################################
    [Fact]
    public void Short_of_room_the_widest_wrapping_column_narrows_first()
    {
        IReadOnlyList<double> widths = PinTableGeometry.TextWidths(Columns, 400);

        Assert.Equal([16, 60, 89, 50, 105], widths);
        Assert.True(widths.Sum() + (4 * PinTableGeometry.Gap) <= 400);
    }

    // Shorter still, the wrapping columns come down together to one width, none below its widest word.
    [Fact]
    public void Shorter_still_the_wrapping_columns_come_down_together_none_below_its_widest_word()
    {
        PinTableColumn[] twoWide =
        [
            new(16, 16, Wraps: false),
            new(140, 30, Wraps: true),
            new(137, 62, Wraps: true),
        ];

        // 16 + 2 gaps leave 200 for the two: 100 each.
        Assert.Equal([16, 100, 100], PinTableGeometry.TextWidths(twoWide, 256));

        // 16 + 2 gaps leave 110: the reading stops at "oscilloscope" (62), the signal takes the rest.
        Assert.Equal([16, 48, 62], PinTableGeometry.TextWidths(twoWide, 166));
    }

    // ###########################################################################################
    // Too narrow for every word whole - a popup of its least width - the wrapping columns come down further together,
    // breaking their words, so every text stays in sight: 120 for the three, whose widest words take 191, is 40 each.
    // ###########################################################################################
    [Fact]
    public void Too_narrow_for_every_word_the_wrapping_columns_break_their_words_together()
    {
        IReadOnlyList<double> widths = PinTableGeometry.TextWidths(Columns, 146 + 120);

        Assert.Equal([16, 40, 40, 50, 40], widths);
        Assert.Equal(146 + 120, widths.Sum() + (4 * PinTableGeometry.Gap));
    }

    // Narrower than even that, each wrapping column stands at the narrowest, and the list is wider than its room.
    [Fact]
    public void Narrower_still_each_wrapping_column_stands_at_the_narrowest()
    {
        Assert.Equal(
            [16, PinTableGeometry.Narrowest, PinTableGeometry.Narrowest, 50, PinTableGeometry.Narrowest],
            PinTableGeometry.TextWidths(Columns, 150));

        // A word narrower than that keeps its own width.
        PinTableColumn[] shortWord = [new(16, 16, Wraps: false), new(60, 10, Wraps: true)];
        Assert.Equal([16, 10], PinTableGeometry.TextWidths(shortWord, 30));
    }

    // ###########################################################################################
    // A column's widest word is of the pieces its texts wrap between: at spaces, and after a "-" or a "/", as the text's
    // own line breaking breaks them - seen whole, "Net-(U16-Pad10)" held the net column wide and the reading's heading
    // broke "oscilloscope" in a popup of its usual width.
    // ###########################################################################################
    [Theory]
    [InlineData("Net-(U16-Pad10)", new[] { "Net-", "(U16-", "Pad10)" })]
    [InlineData("SYNC/LUM", new[] { "SYNC/", "LUM" })]
    [InlineData("A0 / A8", new[] { "A0", "/", "A8" })]
    [InlineData("12.0V (+12VDC power rail)", new[] { "12.0V", "(+12VDC", "power", "rail)" })]
    [InlineData("  Pulsing  ", new[] { "Pulsing" })]
    [InlineData("oscilloscope", new[] { "oscilloscope" })]
    [InlineData("", new string[0])]
    [InlineData(null, new string[0])]
    public void A_text_wraps_between_its_words_and_after_a_hyphen_or_a_slash(string? text, string[] pieces)
    {
        Assert.Equal(pieces, PinTableGeometry.PiecesOf(text));
    }

    // A width is whole pixels, rounded down, so the columns never take more than the room.
    [Fact]
    public void A_narrowed_width_is_whole_pixels_never_more_than_the_room()
    {
        PinTableColumn[] two = [new(100, 10, Wraps: true), new(100, 10, Wraps: true)];

        IReadOnlyList<double> widths = PinTableGeometry.TextWidths(two, 20 + 151);

        Assert.Equal([75, 75], widths);
    }
}
