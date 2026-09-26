using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE REVIEW QUEUE AS ONE ACCOUNT SEES IT (GET /api/review/queue) - which submissions, and what
    // each row says about itself.
    //
    // *** FILTERED TO WHAT THIS ACCOUNT MAY DECIDE. *** An administrator sees everything; a
    // maintainer sees their systems' submissions. The rule is ReviewAuthority's, applied row by row
    // - the queue is small, and one rule in one place beats a second copy of it in SQL.
    //
    // *** THE TWO BADGES (owner request, 2026-09-26) ARE THE DETAIL'S OWN ANSWERS. *** The queue
    // list in the maintainer application shows "New system" / "Published system" and "Awaiting
    // your review". A new system is one with no published board - PublishedBoardLocator, which the
    // detail's comparison reads through. Awaiting you is ApprovalStatus.CanApprove - the answer the
    // Approve button follows - judged with the stored shared-files flag: re-checking the tree needs
    // each submission's payload, which the queue deliberately never loads. The flag is only a
    // floor, but a first approval stores it, so the two answers part only for an unapproved
    // submission whose shared file changed since it arrived - and the detail, opened next, is right.
    //
    // One approvals read per row. The queue is at most ReviewEndpoints.DefaultQueueLimit rows.
    // ###########################################################################################
    public static class ReviewQueueFlow
    {
        public static async Task<ReviewQueueAnswer> BuildAsync(
            ReviewAccess access,
            IReadOnlyList<SubmissionRecord> queued,
            string? dataTreeRoot,
            ISubmissionStore store,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(queued);

            var entries = new List<ReviewQueueEntry>();

            foreach (SubmissionRecord record in queued)
            {
                if (!ReviewAuthority.CanReview(access, record))
                    continue;

                bool isNewSystem = !PublishedBoardLocator.LocateSystem(dataTreeRoot, record.SystemId).Exists;

                ApprovalStatus approval = await ApprovePublishFlow
                    .ApprovalStatusAsync(access, record, record.TouchesSharedFiles, store, accounts, cancellationToken)
                    .ConfigureAwait(false);

                bool awaitsYou = ReviewAuthority.CanPublish(access, record) && approval.CanApprove;

                entries.Add(ReviewQueueFlow.Entry(record, isNewSystem, awaitsYou));
            }

            // Everything in a filtered queue is publishable by the caller - CanPublish is kept for
            // review apps built when the answer could be false. IsAdministrator lets the app offer
            // the administrator's screens to the one person who may use them; the server refuses
            // everyone else regardless.
            return new ReviewQueueAnswer(
                CanPublish: true,
                IsAdministrator: access.Account.IsAdministrator,
                Count: entries.Count,
                Submissions: entries);
        }

        // ###########################################################################################
        // One submission as a queue row - also the `submission` of the detail answer, with the
        // detail's own two answers.
        // ###########################################################################################
        public static ReviewQueueEntry Entry(SubmissionRecord record, bool? isNewSystem, bool? awaitsYou)
        {
            ArgumentNullException.ThrowIfNull(record);

            return new ReviewQueueEntry(
                record.Id,
                record.SystemId,
                record.State,
                record.Summary,
                record.ContactEmail,
                record.BaseRevision,
                record.CreatedUtc,
                record.DecidedUtc,

                // So the administrator can see WHY a submission is theirs as well as a maintainer's.
                record.TouchesSharedFiles,
                isNewSystem,
                awaitsYou);
        }
    }
}
