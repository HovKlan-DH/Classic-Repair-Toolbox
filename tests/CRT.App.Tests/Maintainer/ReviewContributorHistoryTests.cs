using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers ReviewContributorHistory - the Contributor view in words (owner request, 2026-09-30: "all
// information about the contributor, to get an honest opinion if this person can be trusted").
// ###########################################################################################
public sealed class ReviewContributorHistoryTests
{
    private static readonly DateTimeOffset Sent = new(2026, 9, 25, 10, 0, 0, TimeSpan.Zero);

    private static ContributorSubmissionEntry Entry(long id, string state, string? summary = "Fixed U8", string? comment = null) =>
        new(id, "Commodore/C64/250407", summary, state, Sent, Sent.AddDays(1), comment);

    private static ReviewContributorFacts Facts(
        bool? signedIn = true,
        DateTimeOffset? since = null,
        IReadOnlyList<ContributorSubmissionEntry>? submissions = null,
        int published = 0,
        int rejected = 0,
        int? publishedToStable = null) =>
        new("dh@example.com", "Dennis", published, 0, 0, rejected, signedIn, since, submissions, publishedToStable);

    [Fact]
    public void The_heading_names_them_as_the_contributor_line_does()
    {
        Assert.Equal("Dennis (dh@example.com)", ReviewContributorHistory.Heading(Facts()));
    }

    // ###########################################################################################
    // *** SENT WITHOUT AN ACCOUNT IS SAID PLAINLY. *** Such an address was typed in when sending,
    // and nothing checked it - so the record under it belongs to whoever uses that address. The one
    // fact no count gives, and exactly what an honest opinion of a contributor must weigh.
    // ###########################################################################################
    [Fact]
    public void The_account_line_says_whether_the_address_was_checked()
    {
        Assert.Equal(
            "Sent with an account made 2026-September-2 - the address is verified.",
            ReviewContributorHistory.AccountLine(Facts(signedIn: true, since: new DateTimeOffset(2026, 9, 2, 8, 0, 0, TimeSpan.Zero))));

        Assert.Equal("Sent with an account - the address is verified.", ReviewContributorHistory.AccountLine(Facts(signedIn: true)));

        Assert.Equal(
            "Sent without an account - the address was typed in when sending, and nothing checked it.",
            ReviewContributorHistory.AccountLine(Facts(signedIn: false)));

        // An older server does not say: no line, never a guess.
        Assert.Null(ReviewContributorHistory.AccountLine(Facts(signedIn: null)));
    }

    // The contributor line's counts, starting a line of their own - so the two cannot count apart.
    [Fact]
    public void The_counts_start_a_line_of_their_own()
    {
        ReviewContributorFacts facts = Facts(published: 4, rejected: 1, publishedToStable: 4);

        Assert.Equal(
            "[5] submissions in total whereof [4] published to stable and [1] rejected",
            string.Concat(ReviewContributorHistory.Counts(facts).Select(run => run.Text)));

        Assert.Equal("No other submissions", string.Concat(ReviewContributorHistory.Counts(Facts()).Select(run => run.Text)));
    }

    // The heading over the list says when the server stopped at its limit, and says so when an
    // older server sent no list at all rather than claiming there is nothing. Never "their"
    // (owner request, 2026-10-01: it read "Their other submissions, newest first").
    [Fact]
    public void The_lists_heading_says_what_the_list_is()
    {
        Assert.Equal(
            "Other submissions, newest first",
            ReviewContributorHistory.SubmissionsHeading(Facts(submissions: [Entry(2, "merged")], published: 1)));

        Assert.Equal(
            "The newest 1 of 3 other submissions",
            ReviewContributorHistory.SubmissionsHeading(Facts(submissions: [Entry(2, "merged")], published: 3)));

        // Nothing at all: no heading - the counts line above already reads "No other submissions",
        // and the two said it twice, one under the other (code review, 2026-10-01).
        Assert.Equal("No other submissions", string.Concat(ReviewContributorHistory.Counts(Facts(submissions: [])).Select(run => run.Text)));
        Assert.Null(ReviewContributorHistory.SubmissionsHeading(Facts(submissions: [])));

        // Counted but not listed still says so, rather than nothing or a contradiction.
        Assert.Equal(
            "None of the other submissions could be listed.",
            ReviewContributorHistory.SubmissionsHeading(Facts(submissions: [], published: 2)));

        Assert.Equal(
            "The server does not list this contributor's earlier submissions yet.",
            ReviewContributorHistory.SubmissionsHeading(Facts(submissions: null, published: 2)));
    }

    // ###########################################################################################
    // *** NEVER "THEIR" OR "THEM" FOR THE CONTRIBUTOR (owner request, 2026-10-01: "never refer to
    // a contributor as "their""). *** Every line of the Contributor view, in every case it has.
    // ###########################################################################################
    [Fact]
    public void No_line_refers_to_the_contributor_as_their_or_them()
    {
        var lines = new List<string?>();

        foreach (ReviewContributorFacts facts in new[]
        {
            Facts(signedIn: true, since: Sent),
            Facts(signedIn: true),
            Facts(signedIn: false, submissions: [Entry(2, "merged")], published: 1),
            Facts(signedIn: null, submissions: null, published: 2),
            Facts(submissions: [Entry(2, "rejected", comment: "No.")], published: 3, rejected: 1, publishedToStable: 1),
            Facts(submissions: [])
        })
        {
            lines.Add(ReviewContributorHistory.Heading(facts));
            lines.Add(ReviewContributorHistory.AccountLine(facts));
            lines.Add(string.Concat(ReviewContributorHistory.Counts(facts).Select(run => run.Text)));
            lines.Add(ReviewContributorHistory.SubmissionsHeading(facts));
        }

        foreach (string line in lines.OfType<string>())
            Assert.DoesNotMatch(@"(?i)\b(their|theirs|them|they)\b", line);
    }

    // A row in CRT's own words: the Boards screen's and the contributor's Drafts tab's.
    [Fact]
    public void A_row_says_what_it_was_where_how_it_went_and_what_they_were_told()
    {
        ContributorSubmissionEntry entry = Entry(12, "merged", comment: "  Thanks - good catch.  ");

        Assert.Equal("Fixed U8", ReviewContributorHistory.Title(entry));
        Assert.Equal(
            "#12 - Commodore / C64 / 250407 - Published to the BETA source - sent 2026-September-25",
            ReviewContributorHistory.Footer(entry));
        Assert.Equal("Told the contributor: Thanks - good catch.", ReviewContributorHistory.Comment(entry));

        Assert.Equal("(no description given)", ReviewContributorHistory.Title(Entry(13, "published", summary: " ")));
        Assert.Null(ReviewContributorHistory.Comment(Entry(13, "published")));
        Assert.Contains("Published to the stable source", ReviewContributorHistory.Footer(Entry(13, "published")), StringComparison.Ordinal);
    }

    // A turn-down is drawn in the failure colour; a request for changes is not a turn-down.
    [Fact]
    public void Only_a_rejection_is_marked_as_turned_down()
    {
        Assert.True(ReviewContributorHistory.IsTurnedDown(Entry(1, "rejected")));
        Assert.False(ReviewContributorHistory.IsTurnedDown(Entry(2, "changes_requested")));
        Assert.False(ReviewContributorHistory.IsTurnedDown(Entry(3, "merged")));
    }
}
