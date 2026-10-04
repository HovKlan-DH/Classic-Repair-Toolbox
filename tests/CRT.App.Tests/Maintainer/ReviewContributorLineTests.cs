using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// Covers ReviewContributorLine - who sent a submission and how the contributor's other submissions went, in
// words (owner request, 2026-09-26). The Contributor view's name and counts; a one-line
// "From ... - ..." above the table until 2026-09-30, when the owner asked for it to go.
// ###########################################################################################
public sealed class ReviewContributorLineTests
{
    private static string Counts(ReviewContributorFacts facts) =>
        string.Concat(ReviewContributorLine.Counts(facts).Select(run => run.Text));

    // ###########################################################################################
    // "[5] submissions in total whereof [1] published to stable and [4] rejected" (owner request,
    // 2026-10-01 - it read "[5] other submissions: [1] published, [4] rejected"). The last kind is
    // joined by "and", the others by commas.
    // ###########################################################################################
    [Fact]
    public void A_contributors_record_counts_each_kind_of_outcome()
    {
        Assert.Equal(
            "[6] submissions in total whereof [2] published to stable, [1] published to BETA, [1] waiting, [1] changes requested and [1] rejected",
            Counts(new ReviewContributorFacts("dh@hinet.dk", null, Published: 3, Waiting: 1, ChangesRequested: 1, Rejected: 1, PublishedToStable: 2)));

        Assert.Equal(
            "[5] submissions in total whereof [1] published to stable and [4] rejected",
            Counts(new ReviewContributorFacts("dh@hinet.dk", null, Published: 1, Waiting: 0, ChangesRequested: 0, Rejected: 4, PublishedToStable: 1)));
    }

    // ###########################################################################################
    // *** AN OLDER SERVER DOES NOT SAY WHICH REACHED STABLE, SO NOTHING CLAIMS IT. *** Published
    // counts the BETA data and the stable source together; calling a BETA-only submission
    // "published to stable" would be the one wrong thing to say.
    // ###########################################################################################
    [Fact]
    public void Without_the_servers_stable_count_published_is_not_split()
    {
        Assert.Equal(
            "[4] submissions in total whereof [3] published and [1] rejected",
            Counts(new ReviewContributorFacts("dh@hinet.dk", null, Published: 3, Waiting: 0, ChangesRequested: 0, Rejected: 1)));
    }

    // A kind with none is left out, so the line stays short; one is singular.
    [Fact]
    public void Kinds_with_nothing_are_left_out()
    {
        Assert.Equal(
            "[1] submission in total whereof [1] published to BETA",
            Counts(new ReviewContributorFacts("dh@hinet.dk", null, 1, 0, 0, 0, PublishedToStable: 0)));
    }

    // The case a maintainer most needs to notice.
    [Fact]
    public void A_first_contribution_says_there_are_no_other_submissions()
    {
        Assert.Equal("no other submissions", Counts(new ReviewContributorFacts("dh@hinet.dk", null, 0, 0, 0, 0)));
    }

    // A signed-in contributor has a name as well as an address.
    [Fact]
    public void A_signed_in_contributor_is_named_and_addressed()
    {
        Assert.Equal("dh@hinet.dk", ReviewContributorLine.Who(new ReviewContributorFacts("dh@hinet.dk", null, 0, 0, 0, 0)));
        Assert.Equal("Dennis (dh@hinet.dk)", ReviewContributorLine.Who(new ReviewContributorFacts("dh@hinet.dk", "Dennis", 0, 0, 0, 0)));
        Assert.Equal("an unknown contributor", ReviewContributorLine.Who(new ReviewContributorFacts(null, null, 0, 0, 0, 0)));
    }

    // Each count's number is bold in brackets - and only the numbers.
    [Fact]
    public void Only_the_numbers_are_marked_bold()
    {
        IReadOnlyList<ReviewNoteRun> runs = ReviewContributorLine.Counts(new ReviewContributorFacts("dh@hinet.dk", null, 4, 2, 0, 0, PublishedToStable: 3));

        Assert.Equal(["6", "3", "1", "2"], runs.Where(run => run.IsCount).Select(run => run.Text));
    }
}
