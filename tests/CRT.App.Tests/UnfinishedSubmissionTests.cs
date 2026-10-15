using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// UnfinishedSubmission - which receipt a cancelled send forgets (code review, 2026-10-13).
//
// *** THE BUG THIS EXISTS FOR. *** Cancelling while the finalise request was on its way deleted the
// receipt of a submission the server may already have queued and mailed the maintainers about: it
// sat in the review queue with no local trace - no badge, no "My submissions" row, no comments, no
// already-sent guard. Once the finalise request has gone, a cancel forgets nothing.
// ###########################################################################################
public sealed class UnfinishedSubmissionTests
{
    [Fact]
    public void A_cancel_during_the_upload_forgets_the_receipt_once()
    {
        var unfinished = new UnfinishedSubmission();
        unfinished.Created(42);

        Assert.Equal(42, unfinished.TakeForCancel());
        Assert.Equal(0, unfinished.TakeForCancel());
    }

    [Fact]
    public void A_cancel_after_the_finalise_request_has_gone_forgets_nothing()
    {
        var unfinished = new UnfinishedSubmission();
        unfinished.Created(42);
        unfinished.FinaliseSending();

        Assert.Equal(0, unfinished.TakeForCancel());
    }

    [Fact]
    public void Nothing_is_forgotten_before_the_server_hands_out_an_id_or_after_it_is_finished()
    {
        var unfinished = new UnfinishedSubmission();

        Assert.Equal(0, unfinished.TakeForCancel());

        unfinished.Created(7);
        unfinished.Finished();

        Assert.Equal(0, unfinished.TakeForCancel());
    }
}
