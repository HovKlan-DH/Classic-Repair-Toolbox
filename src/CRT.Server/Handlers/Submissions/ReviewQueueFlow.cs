using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE REVIEW QUEUE AS ONE ACCOUNT SEES IT (GET /api/review/queue) - which submissions, and what
    // each row says about itself.
    //
    // *** FILTERED TO WHAT THIS ACCOUNT MAY DECIDE. *** An administrator sees everything; a
    // maintainer sees their boards' submissions. Since the code review of 2026-10-13 the store
    // filters a maintainer's queue by their boards IN SQL (ISubmissionStore.GetQueueForBoardsAsync):
    // reading the oldest hundred rows of every board and filtering afterwards left a maintainer of a
    // quiet board with an empty queue behind a hundred waiting elsewhere. ReviewAuthority's rule is
    // still applied row by row here, so a row the store let through by mistake is never shown.
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
    // *** A FIXED NUMBER OF QUERIES, HOWEVER LONG THE QUEUE (code review, 2026-10-13). *** The
    // approvals, the amendment versions, the discards and the boards are each read ONCE for the
    // whole list, and a board's maintainers once per board - it used to be one approvals read per
    // row, every minute, from every open Maintainer tab.
    //
    // Each row also carries its AMENDMENT VERSION and whether it was made from OLDER DATA than BETA
    // holds (ReviewQueueEntry, 2026-10-13) - see Entry.
    // ###########################################################################################
    public static class ReviewQueueFlow
    {
        // The queue the store reads for this account: everything for an administrator, a
        // maintainer's own boards otherwise - filtered in the query, before its limit.
        public static Task<IReadOnlyList<SubmissionRecord>> ReadAsync(
            ReviewAccess access,
            ISubmissionStore store,
            int limit,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(access);
            ArgumentNullException.ThrowIfNull(store);

            return access.Account.IsAdministrator
                ? store.GetQueueAsync(limit, cancellationToken)
                : store.GetQueueForBoardsAsync(access.MaintainerOf.ToList(), limit, cancellationToken);
        }

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

            List<SubmissionRecord> visible = queued.Where(record => ReviewAuthority.CanReview(access, record)).ToList();
            List<long> ids = visible.Select(record => record.Id).ToList();

            var entries = new List<ReviewQueueEntry>();

            if (visible.Count > 0)
            {
                // Whose contributor discarded their own draft (2026-09-28) - one read for the whole queue.
                IReadOnlyDictionary<long, DateTimeOffset> discarded = await store
                    .GetDraftDiscardsAsync(ids, cancellationToken)
                    .ConfigureAwait(false);

                IReadOnlyDictionary<long, IReadOnlyList<GivenApproval>> approvals = await store
                    .GetApprovalsForAsync(ids, cancellationToken)
                    .ConfigureAwait(false);

                IReadOnlyDictionary<long, int> versions = await store
                    .GetAmendmentVersionsAsync(ids, cancellationToken)
                    .ConfigureAwait(false);

                Dictionary<string, BoardRecord> boards = (await store.ListBoardsAsync(cancellationToken).ConfigureAwait(false))
                    .GroupBy(board => board.BoardId, StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

                // A board's maintainers matter only for a shared-file submission, and are read once
                // per board however many of its submissions wait.
                var hasMaintainers = new Dictionary<string, bool>(StringComparer.Ordinal);

                foreach (SubmissionRecord record in visible)
                {
                    PublishedBoardLocation beta = PublishedBoardLocator.LocateBoard(dataTreeRoot, record.BoardId);
                    bool isNewBoard = !beta.Exists;

                    bool boardHasMaintainers = false;

                    if (record.TouchesSharedFiles)
                    {
                        if (!hasMaintainers.TryGetValue(record.BoardId, out boardHasMaintainers))
                        {
                            boardHasMaintainers = (await accounts.GetMaintainersOfBoardAsync(record.BoardId, cancellationToken).ConfigureAwait(false))
                                .Any(ReviewAuthority.CanGiveMaintainerApproval);

                            hasMaintainers[record.BoardId] = boardHasMaintainers;
                        }
                    }

                    ApprovalStatus approval = ApprovePublishFlow.ApprovalStatusOf(
                        access,
                        record,
                        record.TouchesSharedFiles,
                        boardHasMaintainers,
                        approvals.TryGetValue(record.Id, out IReadOnlyList<GivenApproval>? given) ? given : []);

                    bool awaitsYou = ReviewAuthority.CanPublish(access, record) && approval.CanApprove;

                    entries.Add(ReviewQueueFlow.Entry(
                        record,
                        isNewBoard,
                        awaitsYou,
                        discarded.TryGetValue(record.Id, out DateTimeOffset at) ? at : null,
                        amendmentVersion: versions.TryGetValue(record.Id, out int version) ? version : 0,
                        basedOnOlderRevision: ReviewQueueFlow.IsBasedOnOlderRevision(
                            record.BaseRevision,
                            ReviewQueueFlow.BetaRevisionOf(beta, boards.TryGetValue(record.BoardId, out BoardRecord? board) ? board : null)),
                        newerFromSameContributor: SubmissionReplacementRules.NewerFromSameContributor(record, visible)));
                }
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
        // *** WAS IT MADE FROM OLDER DATA THAN BETA HOLDS NOW? (code review, 2026-10-13) *** A
        // contributor's draft is a whole copy of the board, so a submission made from an older
        // revision carries the older rows - and approving it removes every row published since,
        // with the table showing them as the submission's own removals. `currentRevision` is BETA's,
        // as BetaRevisionOf below decides it for the queue row and the detail alike.
        //
        // True only when both revisions parse and BETA's is LATER (DraftRevisionComparer's
        // OfficialIsNewer - CRT's own drift check); false when they agree or cannot be ordered;
        // null when it cannot be said at all - no base recorded, or a board BETA has no revision for
        // (a new board). A maintainer's change from the Boards screen takes BETA's own revision as
        // its base (BoardEditFlow), so it is never flagged. Two publishes on the SAME day share a
        // revision date, so a submission made between them is not caught - the date is all a draft
        // records.
        // ###########################################################################################
        public static bool? IsBasedOnOlderRevision(string? baseRevision, string? currentRevision)
        {
            if (string.IsNullOrWhiteSpace(baseRevision) || string.IsNullOrWhiteSpace(currentRevision))
                return null;

            return DraftRevisionComparer.Compare(baseRevision, currentRevision) == DraftDriftState.OfficialIsNewer;
        }

        // ###########################################################################################
        // *** THE ONE REVISION "MADE FROM OLDER DATA" IS JUDGED AGAINST (code review, 2026-10-14). ***
        // The queue row used `boards.current_revision` alone while the detail used the revision of
        // BETA's workbook as read, so the same submission could be flagged in one and not the other:
        // a shipped board no publish has recorded (no current_revision) was flagged only in the
        // detail, and a record lagging BETA's workbook could part them the other way. Both now ask
        // this: the revision date in BETA's own workbook - read once per version of the file
        // (BoardDataReader.ReadRevisionDateOnly, the full read's own scan) - or the record's, where
        // BETA has no board or its date cannot be read.
        // ###########################################################################################
        public static string? BetaRevisionOf(PublishedBoardLocation beta, BoardRecord? board)
        {
            ArgumentNullException.ThrowIfNull(beta);

            string read = beta.Exists ? BoardDataReader.ReadRevisionDateOnly(beta.WorkbookPath).Trim() : string.Empty;

            return read.Length > 0 ? read : board?.CurrentRevision;
        }

        // ###########################################################################################
        // One submission as a queue row - also the `submission` of the detail answer, with the
        // detail's own answers.
        // ###########################################################################################
        public static ReviewQueueEntry Entry(
            SubmissionRecord record,
            bool? isNewBoard,
            bool? awaitsYou,
            DateTimeOffset? draftDiscardedUtc = null,
            int? amendmentVersion = null,
            bool? basedOnOlderRevision = null,
            SubmissionRecord? newerFromSameContributor = null)
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
                draftDiscardedUtc,

                // What the maintainer's screen sends back with a decision (2026-10-13).
                amendmentVersion,
                basedOnOlderRevision,

                // The same contributor's newer submission waiting beside it (2026-10-13).
                newerFromSameContributor?.Id,
                newerFromSameContributor?.CreatedUtc);
        }
    }
}
