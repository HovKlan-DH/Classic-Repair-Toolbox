using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // REJECT AND REQUEST CHANGES - the two decisions that publish nothing (code review, 2026-10-13;
    // ReviewEndpoints.DecideAsync until then, which made every judgement in the endpoint itself).
    //
    // *** UNDER THE PUBLISH LOCK, AND THE WRITE IS CONDITIONAL. *** A Reject used to read the row,
    // check it was still pending, and then write `WHERE id = @id` with nothing else - while an
    // approval of the same submission could be publishing it, under the lock, at that moment. The
    // Reject landed after the approval's 'merged': the board was in BETA and the contributor was
    // told it had been refused. Now:
    //
    //   - the decision takes the PublishLock every approval and amendment takes, so none of them is
    //     half done while it looks;
    //   - the write is ISubmissionStore.TrySetDecisionAsync, which moves only a row still pending
    //     or approved - so even a writer outside the lock cannot be overwritten (409 instead).
    //
    // *** THE VERSION THE MAINTAINER SAW (ReviewDecisionRequest.ExpectedVersion). *** A decision on
    // a submission another maintainer changed in the table since this one opened it is refused,
    // 409, with ReviewDecisionRules.AmendedSinceMessage. Null, from an older CRT, is not checked.
    //
    // Pure apart from its two collaborators, so the race is a test against the fakes.
    // ###########################################################################################
    public static class ReviewDecisionFlow
    {
        // The states a decision may move a submission out of - ReviewDecisionRules' "undecided".
        private static readonly IReadOnlyList<string> Undecided = [SubmissionState.Pending, SubmissionState.Approved];

        public delegate bool DecisionRule(ReviewAccess? access, SubmissionRecord? submission, out string reason);

        public static async Task<DecisionOutcome> DecideAsync(
            ReviewAccess? access,
            long submissionId,
            ReviewDecisionRequest? request,
            string newState,
            DecisionRule rule,
            ISubmissionStore store,
            PublishLock publishLock,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(rule);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(publishLock);

            using IDisposable gate = await publishLock.EnterAsync(cancellationToken).ConfigureAwait(false);

            SubmissionRecord? record = await store.FindAsync(submissionId, cancellationToken).ConfigureAwait(false);

            if (record is null)
                return DecisionOutcome.NotFound();

            if (!rule(access, record, out string why))
            {
                // 409, not 403, when the account may review this board: what is wrong is the
                // SUBMISSION's state - somebody else decided it first.
                return ReviewAuthority.CanReview(access, record)
                    ? DecisionOutcome.Conflict(why)
                    : DecisionOutcome.Forbidden(why);
            }

            // *** THE REASON IS REQUIRED FOR BOTH OF THESE OUTCOMES. *** Contributors have no
            // account, so the contact email and this message are the entire feedback channel.
            if (!ReviewDecisionRules.IsUsableReason(request?.Comment, out string reasonProblem))
                return DecisionOutcome.Refused(reasonProblem);

            if (request!.ExpectedVersion is int expected)
            {
                int current = (await store.GetLatestAmendmentAsync(submissionId, cancellationToken).ConfigureAwait(false))?.Version ?? 0;

                if (current != expected)
                    return DecisionOutcome.Conflict(ReviewDecisionRules.AmendedSinceMessage);
            }

            string comment = request.Comment!.Trim();

            bool recorded = await store
                .TrySetDecisionAsync(submissionId, ReviewDecisionFlow.Undecided, newState, access!.Account.Id, comment, nowUtc, cancellationToken)
                .ConfigureAwait(false);

            return recorded
                ? DecisionOutcome.Decided(record, comment)
                : DecisionOutcome.Conflict(ReviewDecisionRules.DecidedMeanwhileMessage);
        }
    }

    // What a decision did: recorded (with the submission as it was, for the mail), or why not.
    public sealed record DecisionOutcome(
        bool IsDecided,
        string Error,
        SubmissionRecord? Submission = null,
        string? Comment = null,
        bool IsNotFound = false,
        bool IsForbidden = false,
        bool IsConflict = false)
    {
        public static DecisionOutcome Decided(SubmissionRecord submission, string comment) =>
            new(true, string.Empty, submission, comment);

        public static DecisionOutcome NotFound() => new(false, "No such submission.", IsNotFound: true);

        public static DecisionOutcome Forbidden(string error) => new(false, error, IsForbidden: true);

        public static DecisionOutcome Conflict(string error) => new(false, error, IsConflict: true);

        public static DecisionOutcome Refused(string error) => new(false, error);
    }
}
