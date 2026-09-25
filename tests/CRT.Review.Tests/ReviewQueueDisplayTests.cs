using CRT.Review.Handlers;

namespace CRT.Review.Tests;

// Covers ReviewQueueDisplay - how a queue row reads.
//
// The wording is tested rather than eyeballed because the queue is the screen a reviewer decides
// from: which submission to open next, and whether anything has been waiting too long. Two of the
// rules below exist specifically to stop the list telling a comfortable lie - a missing timestamp
// rendered as 1970, and a wait rounded up into a coarser, kinder unit.
public sealed class ReviewQueueDisplayTests
{
    private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

    private static ReviewQueueRow Row(
        long id = 42,
        string systemId = "Commodore/C64/250407",
        string summary = "Corrected R12.",
        DateTimeOffset? createdUtc = null) =>
        new(id, systemId, "pending", summary, "someone@example.com", createdUtc);

    // -----------------------------------------------------------------------------------
    // The title
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_title_leads_with_the_id_and_names_the_system()
    {
        // A reviewer works BY SYSTEM - several submissions against one board are reviewed
        // together - so the system has to be readable at a glance. The id is there because it is
        // what every other surface calls this submission.
        Assert.Equal("#42 Commodore/C64/250407", ReviewQueueDisplay.Title(ReviewQueueDisplayTests.Row()));
    }

    [Fact]
    public void A_missing_system_says_so_rather_than_leaving_a_gap()
    {
        // A row reading "#42 " with nothing after it looks like a rendering fault.
        Assert.Equal("#42 (unknown system)", ReviewQueueDisplay.Title(ReviewQueueDisplayTests.Row(systemId: "")));
    }

    // -----------------------------------------------------------------------------------
    // The subtitle
    // -----------------------------------------------------------------------------------

    [Fact]
    public void The_subtitle_shows_what_the_contributor_said()
    {
        string subtitle = ReviewQueueDisplay.Subtitle(
            ReviewQueueDisplayTests.Row(createdUtc: ReviewQueueDisplayTests.Now.AddHours(-3)),
            ReviewQueueDisplayTests.Now);

        Assert.Contains("Corrected R12.", subtitle);
        Assert.Contains("waiting 3 hours", subtitle);
    }

    [Fact]
    public void A_submission_with_no_description_SAYS_SO()
    {
        // The summary field is optional, so blank is a real possibility - and an empty second
        // line reads as a rendering fault rather than as missing information.
        string subtitle = ReviewQueueDisplay.Subtitle(
            ReviewQueueDisplayTests.Row(summary: "   "),
            ReviewQueueDisplayTests.Now);

        Assert.Contains("(no description given)", subtitle);
    }

    // -----------------------------------------------------------------------------------
    // How long it has been waiting - the number that shames a backlog
    // -----------------------------------------------------------------------------------

    [Theory]
    [InlineData(0, "just now")]
    [InlineData(30, "just now")]
    [InlineData(60, "waiting 1 minute")]
    [InlineData(120, "waiting 2 minutes")]
    [InlineData(3600, "waiting 1 hour")]
    [InlineData(7200, "waiting 2 hours")]
    [InlineData(86400, "waiting 1 day")]
    [InlineData(172800, "waiting 2 days")]
    public void A_wait_is_described_in_the_unit_a_person_thinks_in(int secondsWaited, string expected)
    {
        Assert.Equal(
            expected,
            ReviewQueueDisplay.Waiting(
                ReviewQueueDisplayTests.Now.AddSeconds(-secondsWaited),
                ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_wait_TRUNCATES_rather_than_rounding_up_into_a_kinder_unit()
    {
        // *** THIS IS THE NUMBER THAT SHAMES A BACKLOG. *** Rounding in the comforting direction
        // is exactly the wrong behaviour: 13 days must never read as "2 weeks" and 29 days must
        // never read as "a month", because the coarser unit always makes a wait sound shorter
        // than it is, and this figure exists to be uncomfortable.
        Assert.Equal(
            "waiting 13 days",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddDays(-13.9), ReviewQueueDisplayTests.Now));

        Assert.Equal(
            "waiting 29 days",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddDays(-29.9), ReviewQueueDisplayTests.Now));

        // And the boundary itself: 59 minutes is not yet an hour.
        Assert.Equal(
            "waiting 59 minutes",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddMinutes(-59.9), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_MISSING_timestamp_shows_nothing_rather_than_fifty_six_years()
    {
        // The parser deliberately keeps CreatedUtc nullable rather than defaulting it. Rendering
        // a default as "waiting 20000 days" would be a spectacular lie on a screen whose whole
        // job is telling the truth about a backlog.
        Assert.Equal(string.Empty, ReviewQueueDisplay.Waiting(null, ReviewQueueDisplayTests.Now));

        string subtitle = ReviewQueueDisplay.Subtitle(ReviewQueueDisplayTests.Row(), ReviewQueueDisplayTests.Now);

        Assert.Equal("Corrected R12.", subtitle);
        Assert.DoesNotContain("waiting", subtitle);
    }

    [Fact]
    public void A_clock_skew_reads_as_just_now_rather_than_a_negative_wait()
    {
        // The client's clock and the server's need not agree to the second. "waiting -3 minutes"
        // looks broken; "just now" is true enough and is what the reviewer would conclude anyway.
        Assert.Equal(
            "just now",
            ReviewQueueDisplay.Waiting(ReviewQueueDisplayTests.Now.AddMinutes(3), ReviewQueueDisplayTests.Now));
    }

    [Fact]
    public void A_null_row_is_refused_rather_than_drawn_blank()
    {
        Assert.Throws<ArgumentNullException>(() => ReviewQueueDisplay.Title(null!));
        Assert.Throws<ArgumentNullException>(() => ReviewQueueDisplay.Subtitle(null!, ReviewQueueDisplayTests.Now));
    }
}
