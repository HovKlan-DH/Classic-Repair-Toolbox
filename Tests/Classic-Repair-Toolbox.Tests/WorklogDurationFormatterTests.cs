using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// The decimal-hours readback under the "Work done" dialog's Time spent field: 1.25 said back as
// "1 hour and 15 minutes", live as the user types.
//
// The whole point of the line is catching a mistyped decimal (1.4 is an hour and TWENTY-FOUR
// minutes, not an hour and forty), so the arithmetic and the rounding are what matter here, not
// the wording.
//
// No collection attribute: pure maths, no statics, no filesystem, no controls.
public class WorklogDurationFormatterTests
{
    // The example from the request, and the case the whole feature exists for.
    [Fact]
    public void A_quarter_hour_reads_back_as_fifteen_minutes()
    {
        Assert.Equal("1 hour and 15 minutes", WorklogDurationFormatter.Format(1.25));
    }

    // The trap this line is meant to expose: a user typing 1.4 meaning "one hour forty".
    [Fact]
    public void One_point_four_hours_is_twenty_four_minutes_not_forty()
    {
        Assert.Equal("1 hour and 24 minutes", WorklogDurationFormatter.Format(1.4));
    }

    // A whole number of hours says nothing about minutes - "2 hours and 0 minutes" is noise.
    [Theory]
    [InlineData(1.0, "1 hour")]
    [InlineData(2.0, "2 hours")]
    [InlineData(12.0, "12 hours")]
    public void A_whole_number_of_hours_omits_the_minutes(double hours, string expected)
    {
        Assert.Equal(expected, WorklogDurationFormatter.Format(hours));
    }

    // Under an hour, the hours half is omitted entirely rather than shown as a leading zero.
    [Theory]
    [InlineData(0.25, "15 minutes")]
    [InlineData(0.5, "30 minutes")]
    [InlineData(0.75, "45 minutes")]
    public void Less_than_an_hour_omits_the_hours(double hours, string expected)
    {
        Assert.Equal(expected, WorklogDurationFormatter.Format(hours));
    }

    // Singular, on both halves - the plural "1 hours" is the kind of thing that reads as a bug in a
    // dialog a customer's invoice is built from.
    [Fact]
    public void One_of_each_unit_is_singular()
    {
        Assert.Equal("1 hour and 1 minute", WorklogDurationFormatter.Format(1.0 + (1.0 / 60.0)));
    }

    // Nothing typed yet, so nothing to say - the hint line stays blank rather than announcing zero.
    [Theory]
    [InlineData(0.0)]
    [InlineData(-1.5)]
    public void Zero_and_negative_produce_nothing(double hours)
    {
        Assert.Empty(WorklogDurationFormatter.BuildParts(hours));
        Assert.Equal(string.Empty, WorklogDurationFormatter.Format(hours));
    }

    // A value below half a minute rounds away to nothing rather than producing a "0 minutes" tail.
    [Fact]
    public void A_sliver_under_half_a_minute_produces_nothing()
    {
        Assert.Equal(string.Empty, WorklogDurationFormatter.Format(0.001));
    }

    // The total is rounded ONCE and then split. Rounding the halves separately lets 1.9999 h come
    // out as "1 hour and 60 minutes", which is the bug this asserts against.
    [Fact]
    public void A_value_just_under_a_whole_hour_rolls_up_rather_than_showing_sixty_minutes()
    {
        Assert.Equal("2 hours", WorklogDurationFormatter.Format(1.9999));
    }

    // Rounding is to the NEAREST minute, away from zero at the midpoint.
    [Theory]
    [InlineData(0.008, "")]           // 0.48 min -> nothing
    [InlineData(0.009, "1 minute")]   // 0.54 min -> 1
    [InlineData(1.341, "1 hour and 20 minutes")]  // 80.46 min
    [InlineData(1.342, "1 hour and 21 minutes")]  // 80.52 min
    public void Values_round_to_the_nearest_whole_minute(double hours, string expected)
    {
        Assert.Equal(expected, WorklogDurationFormatter.Format(hours));
    }

    // The parts exist so the NUMBERS can be drawn bold and the words plain - the same reason
    // WorkbookSummary hands back Stat parts. A finished string cannot express that, so the split
    // itself is pinned: each part is a bare number plus the words after it, and concatenating them
    // rebuilds Format's output exactly.
    [Fact]
    public void The_parts_split_the_numbers_from_the_words_and_rebuild_the_whole_string()
    {
        var parts = WorklogDurationFormatter.BuildParts(1.25);

        Assert.Equal(2, parts.Count);
        Assert.Equal("1", parts[0].Number);
        Assert.Equal(" hour and ", parts[0].Words);
        Assert.Equal("15", parts[1].Number);
        Assert.Equal(" minutes", parts[1].Words);

        Assert.Equal(
            WorklogDurationFormatter.Format(1.25),
            string.Concat(parts.Select(part => part.Number + part.Words)));
    }

    // The joining " and " belongs to the hours part, so a caller adding the parts in order never
    // has to know whether a second one is coming. With no minutes there is nothing to join to.
    [Fact]
    public void A_lone_hours_part_carries_no_joining_word()
    {
        var parts = WorklogDurationFormatter.BuildParts(3.0);

        Assert.Single(parts);
        Assert.Equal(" hours", parts[0].Words);
    }

    // ---------------------------------------------------------------------------------------------
    // BuildStats - the same duration as WorkbookSummary.Stat parts, for the surfaces that render a
    // LINE of stats (the Workbooks summary strip, the exported PDF). These walk a Stat list and
    // bold each Number, so a duration has to arrive already split.
    // ---------------------------------------------------------------------------------------------

    // The second half is marked JoinedToPrevious so the renderer's " . " separator does not land
    // between "1 hour and" and "15 minutes" and split ONE figure into what reads as two.
    [Fact]
    public void A_two_part_duration_marks_only_its_second_stat_as_joined()
    {
        var stats = WorklogDurationFormatter.BuildStats(1.25);

        Assert.Equal(2, stats.Count);

        Assert.Equal("1", stats[0].Number);
        Assert.Equal(" hour and ", stats[0].Suffix);
        Assert.False(stats[0].JoinedToPrevious);

        Assert.Equal("15", stats[1].Number);
        Assert.Equal(" minutes", stats[1].Suffix);
        Assert.True(stats[1].JoinedToPrevious);
    }

    // A one-part duration is joined to nothing - it is a stat of its own, and its separator from
    // whatever precedes it in the line is the caller's to draw as usual.
    [Fact]
    public void A_one_part_duration_is_a_single_unjoined_stat()
    {
        var stats = WorklogDurationFormatter.BuildStats(0.75);

        Assert.Single(stats);
        Assert.Equal("45", stats[0].Number);
        Assert.Equal(" minutes", stats[0].Suffix);
        Assert.False(stats[0].JoinedToPrevious);
    }

    // Zero contributes NO stat at all rather than a "0 minutes" one: on a headline of real figures
    // it would be the only item reporting the absence of one. The surfaces that show it are all
    // built to omit an item, which is why this returns an empty list rather than a placeholder.
    [Theory]
    [InlineData(0.0)]
    [InlineData(-2.0)]
    [InlineData(0.001)]
    public void A_duration_with_no_whole_minutes_contributes_no_stats(double hours)
    {
        Assert.Empty(WorklogDurationFormatter.BuildStats(hours));
    }

    // The stats say the same thing the parts do - one formatter, so a strip and a PDF built from
    // different halves of this class cannot disagree about how long a job took.
    [Fact]
    public void The_stats_and_the_plain_string_say_the_same_thing()
    {
        var stats = WorklogDurationFormatter.BuildStats(2.5);

        Assert.Equal(
            WorklogDurationFormatter.Format(2.5),
            string.Concat(stats.Select(stat => stat.Number + stat.Suffix)));
    }

    // NaN and infinity cannot come from the NumericUpDown, but a hand-edited entries.json can carry
    // either, and Math.Round throws on them via the (long) cast rather than returning nonsense.
    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    public void Non_finite_values_produce_nothing_rather_than_throwing(double hours)
    {
        Assert.Equal(string.Empty, WorklogDurationFormatter.Format(hours));
    }

    // ###########################################################################################
    // AN IMPLAUSIBLE VALUE PRODUCES NOTHING, rather than nonsense. BuildParts converts hours to
    // minutes through a (long) cast, and an out-of-range cast is UNCHECKED in C# - it does not
    // throw, it yields an implementation-defined value. Without an upper bound a pasted 1e20 came
    // back either as an empty readback (for a number the user had just typed) or as a saturated
    // long rendering as roughly "153722867280912930 hours and 55 minutes", which then flowed on
    // into the summary strip and the exported PDF.
    //
    // These fail against the unbounded version.
    // ###########################################################################################
    [Theory]
    [InlineData(1e20)]
    [InlineData(1e30)]
    [InlineData(double.MaxValue)]
    public void An_implausibly_large_duration_produces_nothing_rather_than_a_nonsense_figure(double hours)
    {
        Assert.Equal(string.Empty, WorklogDurationFormatter.Format(hours));
    }

    // The bound is generous, so nothing a real repair could log is refused by it - a value right at
    // the limit still formats normally.
    [Fact]
    public void A_duration_at_the_limit_still_formats()
    {
        Assert.Equal("876000 hours", WorklogDurationFormatter.Format(WorklogDurationFormatter.MaximumHours));
    }
}
