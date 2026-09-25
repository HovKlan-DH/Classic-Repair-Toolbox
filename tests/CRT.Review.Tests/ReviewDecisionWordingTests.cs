using CRT.Review.Handlers;

namespace CRT.Review.Tests;

// Covers ReviewDecisionWording - what a decision says before it is sent and after it lands.
//
// Wording is tested rather than eyeballed for the same reason the rest of this screen is: a
// reviewer acts on these sentences, and approving cannot be undone.
public sealed class ReviewDecisionWordingTests
{
    [Fact]
    public void The_client_comment_minimum_is_the_value_the_SERVER_enforces()
    {
        // *** THE CLIENT MUST NEVER BE THE STRICTER OF THE TWO. *** This check exists only to save
        // a round trip; ReviewDecisionRules on the server owns the rule and refuses regardless. If
        // the client demanded MORE, it would refuse something the server would have accepted and
        // the reviewer would have no way past it.
        //
        // The value is asserted literally rather than against the server's constant, because the
        // review app does not reference CRT.Server and should not start doing so for a test. The
        // matching assertion lives on the server side, in ReviewDecisionRulesTests, where both
        // constants are visible - that is the one that fails if either moves.
        Assert.Equal(10, ReviewDecisionWording.MinimumCommentLength);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("no")]
    [InlineData("          x")]
    public void A_comment_that_tells_the_contributor_NOTHING_is_refused(string? comment)
    {
        // Contributing needs no account, so this message is the entire feedback channel. A
        // rejection with nothing in it is indistinguishable from being ignored.
        Assert.False(ReviewDecisionWording.IsUsableComment(comment, out string problem));
        Assert.NotEmpty(problem);
    }

    [Fact]
    public void An_ordinary_comment_is_accepted()
    {
        Assert.True(ReviewDecisionWording.IsUsableComment(
            "The highlight for U8 looks like it is on the wrong pin.", out _));
    }

    [Fact]
    public void An_APPROVAL_names_the_published_revision()
    {
        // That value is what every contributor's next submission diffs against, so seeing it
        // confirmed is how a reviewer knows the publish reached the tree rather than merely being
        // accepted.
        string described = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Approve, new ReviewDecisionResult("merged", "2026-September-22"));

        Assert.Contains("2026-September-22", described);
    }

    [Fact]
    public void An_approval_with_NO_revision_still_reads_sensibly()
    {
        // Rather than "Published at revision ." - a server that answered without one is unusual
        // but must not produce a broken sentence.
        string described = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Approve, new ReviewDecisionResult("merged", ""));

        Assert.Equal("Published.", described);
    }

    [Fact]
    public void REJECT_and_REQUEST_CHANGES_say_DIFFERENT_things()
    {
        // They are different outcomes for the contributor - one ends the submission, the other
        // hands it back to be fixed - so a reviewer must be able to tell which they just did.
        string rejected = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Reject, new ReviewDecisionResult("rejected", ""));

        string returned = ReviewDecisionWording.Describe(
            ReviewDecisionKind.RequestChanges, new ReviewDecisionResult("changes_requested", ""));

        Assert.NotEqual(rejected, returned);
        Assert.Contains("Rejected", rejected);
        Assert.Contains("changes", returned, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void A_null_result_is_refused_rather_than_described_blankly()
    {
        Assert.Throws<ArgumentNullException>(
            () => ReviewDecisionWording.Describe(ReviewDecisionKind.Approve, null!));
    }
}
