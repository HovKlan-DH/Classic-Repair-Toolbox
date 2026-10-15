using System.Threading;

namespace Handlers.Online
{
    // ###########################################################################################
    // WHICH RECEIPT A CANCELLED SEND FORGETS (code review, 2026-10-13) - the Submit dialog's
    // bookkeeping, apart so the rule is tested.
    //
    // The receipt is written as soon as the server hands out the submission's id (so a crash still
    // leaves one), and a CANCEL during the upload forgets it again: the contributor stopped on
    // purpose, and "Never finished sending" in red over the earlier good submission would be news
    // about nothing. But once the FINALISE request has gone, the server may queue the submission and
    // mail the maintainers whatever is pressed afterwards - so from then on a cancel forgets
    // nothing, and the status check says what became of it.
    //
    // Set on the pool thread that sends, taken on the UI thread that cancels - hence Interlocked.
    // ###########################################################################################
    internal sealed class UnfinishedSubmission
    {
        private long thisSubmissionId;

        // The server handed out this submission's id; its receipt is written.
        public void Created(long submissionId) => Interlocked.Exchange(ref this.thisSubmissionId, submissionId);

        // The finalise request is about to go: the receipt is the server's to settle now.
        public void FinaliseSending() => Interlocked.Exchange(ref this.thisSubmissionId, 0);

        // The submission was finalised and its receipt updated.
        public void Finished() => Interlocked.Exchange(ref this.thisSubmissionId, 0);

        // The receipt a cancel forgets - 0 for none - taken once.
        public long TakeForCancel() => Interlocked.Exchange(ref this.thisSubmissionId, 0);
    }
}
