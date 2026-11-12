using Handlers.Tours;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TourWording - what the introduction tours' cards say around the steps. The first card says plainly that it IS a tour
// (owner request, 2026-11-11: "it should clearly show this is an "Introduction tour""), and a tour of what changed names
// every version it covers.
// ###########################################################################################
public sealed class TourWordingTests
{
    private static readonly TourOfTab Tour = new(TourTab.Schematics, "Schematics", "Here you see your board.", []);

    [Fact]
    public void The_first_card_says_which_kind_of_tour_it_is()
    {
        Assert.Equal("Introduction tour", TourWording.Kicker(TourKind.Introduction));
        Assert.Equal("Tour of what's new", TourWording.Kicker(TourKind.WhatIsNew));
    }

    // A tour names a tab in bold, never in quotes (owner request, 2026-11-11) - the headline is bold as a whole.
    [Fact]
    public void The_headline_names_the_tab_without_quotes()
    {
        Assert.Equal("The Schematics tab", TourWording.WelcomeHeadline(Tour));
        Assert.Equal("You can take the tours again from the <b>Configuration</b> tab.", TourWording.TakeAgainFoot);
    }

    // No count of the steps (owner request, 2026-11-11: "skip the counter").
    [Fact]
    public void An_introduction_says_what_the_tab_is_for_and_that_the_tour_shows_the_most_important_things()
    {
        Assert.Equal(
            "Here you see your board.\n\nThis short tour shows the most important things here. It takes about a minute.",
            TourWording.WelcomeBody(Tour, TourKind.Introduction, 5, []));

        Assert.Equal(
            "Here you see your board.\n\nThis short tour shows the most important thing here.",
            TourWording.WelcomeBody(Tour, TourKind.Introduction, 1, []));
    }

    // Versions skipped are all named: from 4.0.0 straight to 4.2.0 is what 4.1.0 and 4.2.0 brought.
    [Fact]
    public void A_tour_of_what_changed_names_every_version_it_covers()
    {
        Assert.Equal(
            "The <b>Schematics</b> tab has changed since you last used it. This short tour shows the most important changes, new in CRT 4.1.0 and 4.2.0.",
            TourWording.WelcomeBody(Tour, TourKind.WhatIsNew, 3, ["4.1.0", "4.2.0"]));

        Assert.Equal(
            "The <b>Schematics</b> tab has changed since you last used it. This short tour shows the most important change, new in CRT 4.3.0.",
            TourWording.WelcomeBody(Tour, TourKind.WhatIsNew, 1, ["4.3.0"]));
    }

    // ###########################################################################################
    // What an earlier tour left out, as it was not on screen then (code review, 2026-11-12): said as such - alone, or after
    // what changed.
    // ###########################################################################################
    [Fact]
    public void A_tour_of_what_an_earlier_one_left_out_says_so()
    {
        Assert.Equal("Introduction tour, continued", TourWording.Kicker(TourKind.LeftOut));

        Assert.Equal(
            "Here you see your board.\n\nThis short tour shows what an earlier tour of this tab left out, as it was not on screen then.",
            TourWording.WelcomeBody(Tour, TourKind.LeftOut, 0, [], leftOut: 1));

        Assert.Equal(
            "The <b>Schematics</b> tab has changed since you last used it. This short tour shows the most important change, new in CRT 4.1.0. " + TourWording.AlsoLeftOut,
            TourWording.WelcomeBody(Tour, TourKind.WhatIsNew, 1, ["4.1.0"], leftOut: 2));
    }

    [Theory]
    [InlineData(new string[0], "")]
    [InlineData(new[] { "4.1.0" }, "4.1.0")]
    [InlineData(new[] { "4.1.0", "4.2.0" }, "4.1.0 and 4.2.0")]
    [InlineData(new[] { "4.1.0", "4.2.0", "4.3.0" }, "4.1.0, 4.2.0 and 4.3.0")]
    public void Versions_are_listed_as_a_sentence_lists_them(string[] versions, string listed)
    {
        Assert.Equal(listed, TourWording.Listed(versions));
    }

    [Fact]
    public void A_step_says_where_the_tour_stands_and_what_it_is_new_in()
    {
        Assert.Equal("Step 2 of 5", TourWording.StepCounter(2, 5));
        Assert.Equal("New in 4.1.0", TourWording.NewIn("4.1.0"));
    }
}
