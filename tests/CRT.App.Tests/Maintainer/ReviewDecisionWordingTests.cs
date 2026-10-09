using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// Covers ReviewDecisionWording - what a decision says before it is sent and after it lands.
//
// Wording is tested rather than eyeballed for the same reason the rest of this screen is: a
// maintainer acts on these sentences, and approving cannot be undone.
public sealed class ReviewDecisionWordingTests
{
    [Fact]
    public void The_client_comment_minimum_is_the_value_the_SERVER_enforces()
    {
        // *** THE CLIENT MUST NEVER BE THE STRICTER OF THE TWO. *** This check exists only to save
        // a round trip; ReviewDecisionRules on the server owns the rule and refuses regardless. If
        // the client demanded MORE, it would refuse something the server would have accepted and
        // the maintainer would have no way past it.
        //
        // The value is asserted literally rather than against the server's constant, because the
        // Maintainer tab's code (CRT.App) does not reference CRT.Server and should not start doing so for a test. The
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
        // confirmed is how a maintainer knows the publish reached the tree rather than merely being
        // accepted.
        string described = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Approve, new ReviewDecisionResult("merged", "2026-September-22"));

        Assert.Contains("2026-September-22", described);
    }

    [Fact]
    public void An_approval_with_NO_revision_still_reads_sensibly()
    {
        // Rather than "Published to BETA at revision ." - a server that answered without one is
        // unusual but must not produce a broken sentence.
        string described = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Approve, new ReviewDecisionResult("merged", ""));

        Assert.StartsWith("Published to BETA.", described);
    }

    [Fact]
    public void An_approval_says_it_went_to_BETA_and_what_to_do_next()
    {
        // Since the two-stage publish (2026-09-25) an approval is not the end: "Published" alone
        // read as done while every user's data still lacked it.
        string described = ReviewDecisionWording.Describe(
            ReviewDecisionKind.Approve, new ReviewDecisionResult("merged", "2026-September-25"));

        Assert.Contains("BETA", described);
        Assert.Contains("publish it from \"Queue: Awaiting push from BETA to stable\"", described);
    }

    [Fact]
    public void REJECT_and_REQUEST_CHANGES_say_DIFFERENT_things()
    {
        // They are different outcomes for the contributor - one ends the submission, the other
        // hands it back to be fixed - so a maintainer must be able to tell which they just did.
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

    // ###########################################################################################
    // *** THE "PLEASE WAIT" WHILE A DECISION IS SENT (owner report, 2026-09-28). *** A small
    // "Working..." under the table was all an approval showed while it published, and the screen
    // read as hung. The sentence on the overlay names the board and what is happening.
    // ###########################################################################################
    [Fact]
    public void An_approval_that_publishes_says_it_is_publishing_to_BETA_and_names_the_board()
    {
        // null is "one approval publishes" - the ordinary item, as ApprovalWording.ApproveButton reads it.
        foreach (ApprovalStatus? approval in new[] { null, Status(canApprove: true, publishes: true) })
        {
            string text = ReviewDecisionWording.Waiting(ReviewDecisionKind.Approve, approval, "Commodore/C128/310378");

            Assert.StartsWith("Publishing Commodore/C128/310378 to BETA.", text, StringComparison.Ordinal);
            Assert.EndsWith("please wait until it is done.", text, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void The_first_of_two_approvals_does_not_claim_to_publish()
    {
        // A shared-file change needs two approvals; the first publishes nothing, and "Publishing to
        // BETA" over it would promise what is not happening - the same rule as the button's text.
        string text = ReviewDecisionWording.Waiting(
            ReviewDecisionKind.Approve, Status(canApprove: true, publishes: false), "Commodore/C64/250407");

        Assert.DoesNotContain("Publishing", text, StringComparison.Ordinal);
        Assert.Contains("Recording your approval", text, StringComparison.Ordinal);
        Assert.Contains("Commodore/C64/250407", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(ReviewDecisionKind.Reject, "Rejecting")]
    [InlineData(ReviewDecisionKind.RequestChanges, "request for changes")]
    public void Sending_a_decision_back_says_which_one_and_names_the_board(ReviewDecisionKind kind, string expected)
    {
        string text = ReviewDecisionWording.Waiting(kind, null, "Commodore/C64/250407");

        Assert.Contains(expected, text, StringComparison.Ordinal);
        Assert.Contains("Commodore/C64/250407", text, StringComparison.Ordinal);
        Assert.DoesNotContain("BETA", text, StringComparison.Ordinal);
        Assert.EndsWith("please wait until it is done.", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void With_no_board_known_the_sentence_still_reads(string? boardId)
    {
        // The detail is read after the submission is chosen; a decision sent before it arrives has
        // no board id to name.
        foreach (ReviewDecisionKind kind in Enum.GetValues<ReviewDecisionKind>())
        {
            string text = ReviewDecisionWording.Waiting(kind, null, boardId);

            Assert.Contains("this board", text, StringComparison.Ordinal);
            Assert.DoesNotContain("  ", text, StringComparison.Ordinal);
        }
    }

    private static ApprovalStatus Status(bool canApprove, bool publishes) =>
        new(
            Required: [ApproverRole.Maintainer, ApproverRole.Administrator],
            Given: [],
            WaitingFor: [ApproverRole.Maintainer, ApproverRole.Administrator],
            YourRole: ApproverRole.Maintainer,
            CanApprove: canApprove,
            ApprovalPublishes: publishes);
}
