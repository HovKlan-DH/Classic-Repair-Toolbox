using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// MaintainerWaitWording - what the "please wait" overlay says, and what the maintainer is told when
// a wait ran into the two-minute limit (owner decision, 2026-09-28: "it must be solid in validating
// if it did finish").
//
// The rule every after-timeout sentence keeps: it says what a fresh look at the server FOUND - it
// finished, it has not (yet), or it could not be checked - and never "it failed", because the server
// usually carries on after the window stops waiting.
// ###########################################################################################
public sealed class MaintainerWaitWordingTests
{
    [Theory]
    [InlineData(ReviewDecisionKind.Approve, "pending", "merged", "but it did finish: the submission is published to BETA.")]
    [InlineData(ReviewDecisionKind.Approve, "pending", "approved", "but it did finish: your approval is recorded.")]
    [InlineData(ReviewDecisionKind.Approve, "approved", "merged", "but it did finish: the submission is published to BETA.")]
    [InlineData(ReviewDecisionKind.Reject, "pending", "rejected", "but it did finish: the submission is rejected")]
    [InlineData(ReviewDecisionKind.RequestChanges, "pending", "changes_requested", "but it did finish: the submission is returned to the contributor")]
    public void A_decision_that_landed_is_reported_as_done(ReviewDecisionKind kind, string before, string now, string expected)
    {
        Assert.Contains(expected, MaintainerWaitWording.DecisionAfterTimeout(kind, before, now), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // *** THE FIRST OF TWO APPROVALS ALREADY GIVEN IS NOT "RECORDED" AGAIN. *** A submission that
    // was already "approved" (by the other approver) and still is has NOT moved - this approval did
    // not land yet, and saying "recorded" would tell the maintainer to stop waiting for it.
    // ###########################################################################################
    [Fact]
    public void A_decision_that_did_not_move_the_submission_says_it_is_still_waiting()
    {
        string pending = MaintainerWaitWording.DecisionAfterTimeout(ReviewDecisionKind.Approve, "pending", "pending");
        string stillApproved = MaintainerWaitWording.DecisionAfterTimeout(ReviewDecisionKind.Approve, "approved", "approved");

        foreach (string text in new[] { pending, stillApproved })
        {
            Assert.Contains("the submission is still waiting", text, StringComparison.Ordinal);
            Assert.Contains("look again before trying again", text, StringComparison.Ordinal);
            Assert.DoesNotContain("did finish", text, StringComparison.Ordinal);
        }
    }

    // Somebody else decided it meanwhile: it says what it became, in CRT's words for states.
    [Fact]
    public void A_submission_decided_otherwise_meanwhile_says_what_it_became()
    {
        string text = MaintainerWaitWording.DecisionAfterTimeout(ReviewDecisionKind.Approve, "pending", "rejected");

        Assert.Contains($"\"{SubmissionReceiptPresenter.DescribeState("rejected")}\"", text, StringComparison.Ordinal);
        Assert.DoesNotContain("did finish", text, StringComparison.Ordinal);
    }

    [Fact]
    public void When_the_look_itself_fails_it_says_so_and_does_not_guess()
    {
        string text = MaintainerWaitWording.DecisionAfterTimeout(ReviewDecisionKind.Approve, "pending", null);

        Assert.Contains("whether it finished could not be checked", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(3, 4, "but it did finish: your changes are saved.")]
    [InlineData(3, 3, "and your changes are not saved yet.")]
    [InlineData(3, null, "could not be checked")]
    public void A_save_is_judged_by_whether_the_version_moved(int before, int? now, string expected)
    {
        Assert.Contains(expected, MaintainerWaitWording.SaveAfterTimeout(before, now), StringComparison.Ordinal);
    }

    // A publish and a push-back both take the system off the BETA list when they land.
    [Theory]
    [InlineData(false, "but it did finish")]
    [InlineData(true, "is still waiting in BETA")]
    [InlineData(null, "could not be checked")]
    public void A_publish_to_production_is_judged_by_whether_the_system_left_BETA(bool? stillInBeta, string expected)
    {
        Assert.Contains(expected, MaintainerWaitWording.PublishAfterTimeout("Commodore/C64/250407", stillInBeta), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, "is pushed back to the queue")]
    [InlineData(true, "is still in BETA")]
    public void A_push_back_is_judged_by_whether_the_system_left_BETA(bool stillInBeta, string expected)
    {
        Assert.Contains(expected, MaintainerWaitWording.PushBackAfterTimeout("Commodore/C64/250407", stillInBeta), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, "but it did finish: the files are removed.")]
    [InlineData(1, "1 of them is still there")]
    [InlineData(4, "4 of them are still there")]
    public void A_removal_counts_what_is_still_there(int stillThere, string expected)
    {
        Assert.Contains(expected, MaintainerWaitWording.RemovalAfterTimeout(stillThere), StringComparison.Ordinal);
    }

    [Fact]
    public void The_wait_sentences_name_what_is_happening()
    {
        Assert.Equal("Opening the submission to Commodore/C64/250407...", MaintainerWaitWording.OpeningSubmission("Commodore/C64/250407"));
        Assert.Equal("Removing 1 unused file...", MaintainerWaitWording.RemovingUnusedFiles(1));
        Assert.Equal("Removing 12 unused files...", MaintainerWaitWording.RemovingUnusedFiles(12));
        Assert.Equal("Adding Dennis as a maintainer of Commodore/C64/250407...", MaintainerWaitWording.AddingMaintainer("Dennis", "Commodore/C64/250407"));
        Assert.DoesNotContain("Working", MaintainerWaitWording.ReadingQueue, StringComparison.Ordinal);
    }

    // Nothing to read back: each says what to do next, never to repeat blindly.
    [Fact]
    public void The_sign_in_screen_waits_say_what_to_do_next()
    {
        Assert.Contains("wait a few minutes before asking again", MaintainerWaitWording.ResetCodeNoAnswer, StringComparison.Ordinal);
        Assert.Contains("try signing in with it", MaintainerWaitWording.PasswordNoAnswer, StringComparison.Ordinal);
        Assert.Contains("try signing in", MaintainerWaitWording.InvitationNoAnswer, StringComparison.Ordinal);
    }
}
