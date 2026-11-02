using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// UnfinishedSubmission - which receipt a cancelled send forgets (code review, 2026-10-13).
//
// *** THE BUG THIS EXISTS FOR. *** Cancelling while the finalise request was on its way deleted the
// receipt of a submission the server may already have queued and mailed the maintainers about: it
// sat in the review queue with no local trace - no badge, no "My submissions" row, no comments, no
// already-sent guard. Once the finalise request has gone, a cancel forgets nothing.
//
// In the store's collection: FinaliseSending notes the request on the receipt (code review, 2026-11-02).
// ###########################################################################################
[Collection("SubmissionReceipts")]
public sealed class UnfinishedSubmissionTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public UnfinishedSubmissionTests()
    {
        SubmissionReceiptStore.LoadFrom(Path.Combine(this.thisWorkspace.Root, "submissions.json"));
    }

    public void Dispose()
    {
        SubmissionReceiptStore.LoadFrom(string.Empty);
        this.thisWorkspace.Dispose();
    }

    // ###########################################################################################
    // *** THE RECEIPT SAYS THE FINALISE REQUEST WENT (code review, 2026-11-02). *** Noted as it goes, not
    // when it is answered: an answer lost on the way is the very case it is for. "Board: Name" and "My
    // submissions" tell by it a send that never reached anybody from one a maintainer may have.
    // ###########################################################################################
    [Fact]
    public void The_finalise_request_going_is_noted_on_the_receipt()
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt { SubmissionId = 42, UploadToken = "token" });

        var unfinished = new UnfinishedSubmission();
        unfinished.Created(42);

        Assert.Null(Assert.Single(SubmissionReceiptStore.All).FinaliseSentUtc);

        unfinished.FinaliseSending();

        Assert.NotNull(Assert.Single(SubmissionReceiptStore.All).FinaliseSentUtc);
    }

    // A cancel forgot the receipt first: nothing is written back for it.
    [Fact]
    public void Nothing_is_noted_for_a_submission_a_cancel_took()
    {
        SubmissionReceiptStore.Record(new SubmissionReceipt { SubmissionId = 42, UploadToken = "token" });

        var unfinished = new UnfinishedSubmission();
        unfinished.Created(42);
        Assert.Equal(42, unfinished.TakeForCancel());

        unfinished.FinaliseSending();

        Assert.Null(Assert.Single(SubmissionReceiptStore.All).FinaliseSentUtc);
    }

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
