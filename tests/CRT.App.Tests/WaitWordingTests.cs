using System;
using CRT;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// WaitWording - what CRT and the Maintainer tab both say when a wait ran into the two-minute limit
// (owner decision, 2026-09-28: "it must be solid in validating if it did finish"). One class in
// CRT.UI, so the limit is described the same way in both applications.
//
// The rule: a sentence says what is KNOWN - it finished, it has not (yet), or it could not be
// checked - and never "it failed", because the server usually carries on after the client stops
// listening, and "failed, try again" would send a second publish after one that worked.
// ###########################################################################################
public sealed class WaitWordingTests
{
    [Fact]
    public void The_limit_is_named_from_the_limit_itself()
    {
        Assert.Equal("2 minutes", WaitWording.Limit);
        Assert.Contains("within 2 minutes", WaitWording.NoAnswer, StringComparison.Ordinal);
        Assert.Contains("within 2 minutes", WaitWording.Checking, StringComparison.Ordinal);
    }

    [Fact]
    public void Finished_says_it_did_finish_and_what()
    {
        Assert.Equal(
            "The server did not answer within 2 minutes, but it did finish: your changes are saved.",
            WaitWording.AfterTimeout(true, "your changes are saved", "unused"));
    }

    // Not finished (yet): the server may still be working - look again before repeating it.
    [Fact]
    public void Not_finished_says_what_is_still_the_case_and_to_look_again_before_repeating()
    {
        string text = WaitWording.AfterTimeout(false, "unused", "the submission is still waiting");

        Assert.StartsWith("The server did not answer within 2 minutes, and the submission is still waiting.", text, StringComparison.Ordinal);
        Assert.Contains("look again before trying again", text, StringComparison.Ordinal);
        Assert.DoesNotContain("fail", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unknown_says_it_could_not_be_checked()
    {
        string text = WaitWording.AfterTimeout(null, "unused", "unused");

        Assert.Contains("whether it finished could not be checked", text, StringComparison.Ordinal);
        Assert.DoesNotContain("unused", text, StringComparison.Ordinal);
    }

    // A clause that already ends a sentence is not given a second full stop.
    [Fact]
    public void A_clause_ending_in_a_full_stop_gets_exactly_one()
    {
        Assert.EndsWith("saved.", WaitWording.AfterTimeout(true, "saved.", ""), StringComparison.Ordinal);
        Assert.DoesNotContain("..", WaitWording.AfterTimeout(true, "saved.", ""), StringComparison.Ordinal);
    }

    [Fact]
    public void Local_work_still_running_says_it_carries_on_and_will_say_when_done()
    {
        string text = WaitWording.StillRunning("The PDF export");

        Assert.StartsWith("The PDF export is taking longer than 2 minutes.", text, StringComparison.Ordinal);
        Assert.Contains("carries on", text, StringComparison.Ordinal);
    }

    // ###########################################################################################
    // How long the overlay has waited, in words (owner request, 2026-09-28) - the owner's four
    // examples first, then the edges: singular and plural on each unit, a whole minute without
    // "and 0 seconds", and nothing at all in the first second (the card is up from 0.3 s).
    // ###########################################################################################
    [Theory]
    [InlineData(1, "1 second elapsed")]
    [InlineData(10, "10 seconds elapsed")]
    [InlineData(61, "1 minute and 1 second elapsed")]
    [InlineData(659, "10 minutes and 59 seconds elapsed")]
    [InlineData(2, "2 seconds elapsed")]
    [InlineData(60, "1 minute elapsed")]
    [InlineData(120, "2 minutes elapsed")]
    [InlineData(62, "1 minute and 2 seconds elapsed")]
    [InlineData(121, "2 minutes and 1 second elapsed")]
    [InlineData(4503, "75 minutes and 3 seconds elapsed")]
    public void Elapsed_time_reads_in_words(int seconds, string expected)
    {
        Assert.Equal(expected, WaitWording.Elapsed(TimeSpan.FromSeconds(seconds)));
    }

    // Truncated like a clock: 1.9 s is still one second.
    [Fact]
    public void Part_of_a_second_does_not_count_until_it_is_whole()
    {
        Assert.Equal(string.Empty, WaitWording.Elapsed(TimeSpan.FromMilliseconds(300)));
        Assert.Equal("1 second elapsed", WaitWording.Elapsed(TimeSpan.FromMilliseconds(1900)));
        Assert.Equal(string.Empty, WaitWording.Elapsed(TimeSpan.FromSeconds(-5)));
    }
}
