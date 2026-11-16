using Handlers.Oscilloscope;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The "Oscilloscope" tab's output box keeps its last lines (2026-11-16, performance walkthrough): every 40 ms flush set
// the box's whole text again, and a TextBox lays all of it out at every change, so a long session made each flush
// slower. Every line is in the log as well.
// ###########################################################################################
public sealed class ScopeOutputTextTests
{
    [Fact]
    public void Lines_are_appended_each_ending_in_the_line_end_while_under_the_limit()
    {
        Assert.Equal("a\nb\nc\n", ScopeOutputText.Appended("a\n", ["b", "c"], "\n", kept: 10));
    }

    [Fact]
    public void Past_the_limit_only_the_last_lines_are_kept()
    {
        Assert.Equal("c\nd\ne\n", ScopeOutputText.Appended("a\nb\nc\n", ["d", "e"], "\n", kept: 3));
    }

    // A burst longer than the limit by itself.
    [Fact]
    public void A_burst_of_more_lines_than_are_kept_keeps_its_own_last_lines()
    {
        Assert.Equal("4\r\n5\r\n", ScopeOutputText.Appended("old\r\n", ["1", "2", "3", "4", "5"], "\r\n", kept: 2));
    }

    [Fact]
    public void An_empty_box_and_nothing_appended_stays_empty()
    {
        Assert.Equal(string.Empty, ScopeOutputText.Appended(null, [], "\n", kept: 3));
        Assert.Equal("a\n", ScopeOutputText.Appended("a\n", [], "\n", kept: 3));
    }

    [Fact]
    public void Exactly_the_limit_keeps_every_line()
    {
        Assert.Equal("a\nb\nc\n", ScopeOutputText.Appended("a\nb\n", ["c"], "\n", kept: 3));
    }

    [Fact]
    public void The_box_keeps_a_thousand_lines()
    {
        Assert.Equal(1000, ScopeOutputText.LinesKept);
    }
}
