using CRT.Maintainer.Handlers;
using Handlers.DataHandling;

namespace CRT.Maintainer.Tests;

// ###########################################################################################
// Covers ReviewContributorLine - who sent a submission and how their other submissions went, in
// one line above its table (owner request, 2026-09-26).
// ###########################################################################################
public sealed class ReviewContributorLineTests
{
    private static string Text(ReviewContributorFacts facts) => ReviewContributorLine.For(facts)!.Text;

    [Fact]
    public void A_contributor_with_a_record_is_named_with_each_kind_of_outcome_counted()
    {
        Assert.Equal(
            "From dh@hinet.dk - [6] other submissions: [3] published, [1] waiting, [1] changes requested, [1] rejected",
            Text(new ReviewContributorFacts("dh@hinet.dk", null, Published: 3, Waiting: 1, ChangesRequested: 1, Rejected: 1)));
    }

    // A kind with none is left out, so the line stays short; one is singular.
    [Fact]
    public void Kinds_with_nothing_are_left_out()
    {
        Assert.Equal(
            "From dh@hinet.dk - [1] other submission: [1] published",
            Text(new ReviewContributorFacts("dh@hinet.dk", null, 1, 0, 0, 0)));
    }

    // The case a maintainer most needs to notice.
    [Fact]
    public void A_first_contribution_says_there_are_no_other_submissions()
    {
        Assert.Equal("From dh@hinet.dk - no other submissions", Text(new ReviewContributorFacts("dh@hinet.dk", null, 0, 0, 0, 0)));
    }

    // A signed-in contributor has a name as well as an address.
    [Fact]
    public void A_signed_in_contributor_is_named_and_addressed()
    {
        Assert.Equal("From Dennis (dh@hinet.dk) - no other submissions", Text(new ReviewContributorFacts("dh@hinet.dk", "Dennis", 0, 0, 0, 0)));
        Assert.Equal("From an unknown contributor - no other submissions", Text(new ReviewContributorFacts(null, null, 0, 0, 0, 0)));
    }

    // Each count's number is bold in brackets, like the lines under it - and only the numbers.
    [Fact]
    public void Only_the_numbers_are_marked_bold()
    {
        ReviewNoteLine line = ReviewContributorLine.For(new ReviewContributorFacts("dh@hinet.dk", null, 4, 2, 0, 0))!;

        Assert.Equal(["6", "4", "2"], line.Runs.Where(run => run.IsCount).Select(run => run.Text));
    }

    // An older server sends nothing: no line, rather than a wrong one.
    [Fact]
    public void Nothing_said_is_no_line()
    {
        Assert.Null(ReviewContributorLine.For(null));
    }
}
