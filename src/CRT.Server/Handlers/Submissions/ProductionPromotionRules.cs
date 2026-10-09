using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The two small decisions the production promotion makes about ROWS rather than files
    // (2026-09-25). Pure, so both are unit tests.
    // ###########################################################################################
    public static class ProductionPromotionRules
    {
        // ###########################################################################################
        // Is this board's BETA state newer than what Production has?
        //
        // Only a board with a BETA content hash can be - that hash is written by every publish,
        // so a shipped board nobody has published through the pipeline is waiting for nothing.
        // Compared on the HASH rather than the revision date, because two publishes on one day
        // share a revision date and the second would otherwise never reach Production.
        // ###########################################################################################
        public static bool IsAwaitingProduction(BoardRecord? board)
        {
            if (board is null || string.IsNullOrWhiteSpace(board.ContentHash))
                return false;

            return !string.Equals(board.ContentHash, board.ProductionContentHash, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The state a CONTRIBUTOR is told about their submission.
        //
        // *** "merged" NOW MEANS "IN BETA". *** A merged submission is in the BETA data and reaches
        // everyone only when its board is promoted. So a merged submission decided at or before
        // the board's last promotion went out with it, and the contributor is told "published" -
        // which CRT shows as "Published to the stable source". Before that they are told "merged", which CRT
        // shows as "Published to the BETA source".
        //
        // The database state is NOT changed: "merged" stays the record of what the maintainer did,
        // and this is only how it reads from the outside.
        // ###########################################################################################
        public static string ContributorFacingState(
            string state,
            DateTimeOffset? decidedUtc,
            DateTimeOffset? boardProductionPublishedUtc,
            DateTimeOffset? returnedUtc = null)
        {
            // ###########################################################################################
            // *** "returned": TAKEN BACK OUT OF BETA (code review, 2026-09-27). *** A BETA rollback
            // moves a merged submission back to `pending` WITH the maintainer's reason. Reported as
            // plain "pending", CRT showed a contributor who had been told "Published to the BETA source"
            // a bare "Waiting for review" with no word for what had happened.
            //
            // *** RECORDED, NOT INFERRED (code review, 2026-09-29). *** It was read off "pending and
            // carrying a decision comment", which held only while nothing else ever left a comment
            // on a pending row - a manual fix, a future note to the contributor or an amendment
            // would have announced a rollback that never happened. The rollback now records WHEN it
            // returned the row (`returnedUtc`, migration 0015), at the same instant it writes as the
            // row's decided_utc. So a pending row is "returned" while that record is not older than
            // its latest decision: any decision since moves decidedUtc on and ends it.
            // ###########################################################################################
            if (string.Equals(state, SubmissionState.Pending, StringComparison.Ordinal) &&
                returnedUtc is DateTimeOffset returned &&
                (decidedUtc is null || returned >= decidedUtc.Value))
            {
                return ProductionPromotionRules.ReturnedState;
            }

            if (!string.Equals(state, SubmissionState.Merged, StringComparison.Ordinal))
                return state;

            if (decidedUtc is null || boardProductionPublishedUtc is null)
                return state;

            return boardProductionPublishedUtc.Value >= decidedUtc.Value
                ? ProductionPromotionRules.PublishedState
                : state;
        }

        // Not a database state - see ContributorFacingState. CRT.Data's SubmissionReceipt and
        // DraftRetirement both already treat it as "live".
        public const string PublishedState = "published";

        // Not a database state either - see ContributorFacingState. CRT.Data's
        // SubmissionReceiptPresenter reads it as "Taken back out of BETA".
        public const string ReturnedState = "returned";

        // ###########################################################################################
        // *** WHOSE WORK A PROMOTION WOULD CARRY, for the panel (owner request, 2026-09-27). ***
        //
        // The maintainer's production panel showed the file copy list and nothing else, which is the
        // mechanics of a copy rather than anything that can be judged - "I am not sure if the shown
        // information in the right-side panel is any helpful". The question the button really asks
        // is "whose accepted work am I about to push to every CRT user", and that was not on screen.
        //
        // *** THE WINDOW IS THE MAILS' WINDOW, and that is the point. ***
        // ProductionEndpoints.AfterPublishAsync emails exactly the submissions merged in
        // (lastProductionPublish, now] once a promotion succeeds. Deriving the panel's list from the
        // same bounds means the screen cannot name someone the mail will not reach, or stay silent
        // about someone it will. A null `after` is the FIRST promotion, which carries everything
        // merged so far.
        //
        // NEWEST FIRST: the most recently accepted submission is the one a maintainer is most
        // likely to be looking for, and an unbounded list is read from the top.
        // ###########################################################################################
        //
        // `discarded` (2026-09-28): when each contributor discarded their own draft since sending -
        // CarriedSubmission.DraftDiscardedUtc, so the panel can say so beside their name.
        //
        // `addresses` (2026-09-29): each submission's address as ContributorAddresses resolves it -
        // the account's for a signed-in contributor. Without it only the contact address is known,
        // which is empty for exactly those contributors.
        public static IReadOnlyList<CarriedSubmission> Carrying(
            IReadOnlyList<SubmissionRecord>? merged,
            IReadOnlyDictionary<long, DateTimeOffset>? discarded = null,
            IReadOnlyDictionary<long, string>? addresses = null)
        {
            if (merged is null || merged.Count == 0)
                return [];

            return merged
                .OrderByDescending(submission => submission.DecidedUtc ?? submission.CreatedUtc)
                .Select(submission => new CarriedSubmission(
                    submission.Id,

                    // Both are nullable on the record: a submission can be sent with no account,
                    // and saved with no description. The WORDING says so (ProductionDisplay), so
                    // they travel as they are rather than being substituted here.
                    addresses is not null && addresses.TryGetValue(submission.Id, out string? address)
                        ? address
                        : submission.ContactEmail ?? string.Empty,
                    submission.Summary,
                    submission.DecidedUtc,
                    discarded is not null && discarded.TryGetValue(submission.Id, out DateTimeOffset at) ? at : null))
                .ToList();
        }

        // ###########################################################################################
        // Does this board, waiting for production, wait for THIS account (owner request,
        // 2026-09-27 - the BETA button's badge counts "boards that need your attention")?
        //
        // Yes unless the account has already given its production approval for the BETA state on
        // offer: then the board waits for the OTHER approver a shared-file change needs, and there
        // is nothing for this account to do. The approvals are per BETA content hash, so one given
        // for an earlier BETA state never counts - new content needs a new look.
        //
        // Deliberately cheap: it does not work out whether the promotion needs two approvals at all,
        // which means hashing the board. An account that has not approved can always act - publish,
        // or give the first of two.
        //
        // *** WHILE ONLY ADMINISTRATORS PUBLISH TO STABLE, ALWAYS YES (code review, 2026-10-05). ***
        // Nobody else's approval is asked for then, so an administrator's own always completes the
        // publish - even one given BEFORE the setting was switched on, which left the board "with the
        // other approver", dimmed and off the badge, with nobody prompted to finish it. Asked only of
        // an account that may publish to stable (ReviewAuthority.CanPublishToProduction).
        // ###########################################################################################
        public static bool AwaitsAccount(IReadOnlyList<GivenApproval>? givenForThisBetaState, long accountId, bool administratorsOnly = false) =>
            administratorsOnly ||
            givenForThisBetaState is null ||
            !givenForThisBetaState.Any(approval => approval.AccountId == accountId);
    }
}
