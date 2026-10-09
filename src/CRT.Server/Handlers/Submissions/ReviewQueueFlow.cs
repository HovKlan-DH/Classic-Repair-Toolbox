using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE REVIEW QUEUE AS ONE ACCOUNT SEES IT (GET /api/review/queue) - which submissions, and what
    // each row says about itself.
    //
    // *** FILTERED TO WHAT THIS ACCOUNT MAY DECIDE. *** An administrator sees everything; a
    // maintainer sees their boards' submissions. The rule is ReviewAuthority's, applied row by row
    // - the queue is small, and one rule in one place beats a second copy of it in SQL.
    //
    // *** THE TWO BADGES (owner request, 2026-09-26) ARE THE DETAIL'S OWN ANSWERS. *** The queue
    // list in the Maintainer tab shows "New board" / "Published board" and "Awaiting
    // your review". A new board is one with no published board - PublishedBoardLocator, which the
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

            // Whose contributor discarded their own draft (2026-09-28) - one read for the whole queue.
            IReadOnlyDictionary<long, DateTimeOffset> discarded = await store
                .GetDraftDiscardsAsync(queued.Select(record => record.Id).ToList(), cancellationToken)
                .ConfigureAwait(false);

            foreach (SubmissionRecord record in queued)
            {
                if (!ReviewAuthority.CanReview(access, record))
                    continue;

                bool isNewBoard = !PublishedBoardLocator.LocateBoard(dataTreeRoot, record.BoardId).Exists;

                ApprovalStatus approval = await ApprovePublishFlow
                    .ApprovalStatusAsync(access, record, record.TouchesSharedFiles, store, accounts, cancellationToken)
                    .ConfigureAwait(false);

                bool awaitsYou = ReviewAuthority.CanPublish(access, record) && approval.CanApprove;

                entries.Add(ReviewQueueFlow.Entry(record, isNewBoard, awaitsYou, discarded.TryGetValue(record.Id, out DateTimeOffset at) ? at : null));
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
        public static ReviewQueueEntry Entry(SubmissionRecord record, bool? isNewBoard, bool? awaitsYou, DateTimeOffset? draftDiscardedUtc = null)
        {
            ArgumentNullException.ThrowIfNull(record);

            return new ReviewQueueEntry(
                record.Id,
                record.BoardId,
                record.State,
                record.Summary,
                record.ContactEmail,
                record.BaseRevision,
                record.CreatedUtc,
                record.DecidedUtc,

                // So the administrator can see WHY a submission is theirs as well as a maintainer's.
                record.TouchesSharedFiles,
                isNewBoard,
                awaitsYou,

                // The contributor discarded their own draft after sending it (2026-09-28).
                draftDiscardedUtc);
        }
    }
}
