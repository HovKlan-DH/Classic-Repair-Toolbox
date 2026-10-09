using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHO SENT A SUBMISSION, AND HOW THEIR OTHER SUBMISSIONS WENT (owner request, 2026-09-26) - the
    // submission detail's `contributor`, CRT.Data's ReviewContributorFacts.
    //
    // A maintainer deciding about a submission weighs who it is from: a first contribution is read
    // more carefully than the tenth from someone whose last nine were published. The contact email
    // was already on the queue row, but not said anywhere a maintainer would look, and nothing
    // counted a contributor's record at all.
    //
    // The same contributor as SubmissionReplacementRules means it - the same account, or the same
    // email among submissions sent without one - so a contributor is one person everywhere. Facts is
    // pure and tested; BuildAsync only fetches.
    //
    // *** AND THEIR WHOLE RECORD (owner request, 2026-09-30) *** - the Maintainer tab's Contributor
    // view: "all data we can see for whatever he/she has contributed, and which has been accepted
    // and rejected etc." - to judge whether a contributor can be trusted. The submissions LISTED
    // are exactly the ones COUNTED (Counted), so the list and the counts above it always agree;
    // each in the word its contributor is told (ProductionPromotionRules.ContributorFacingState),
    // the Boards screen's rule. Whether this submission came from an account is said too: an
    // address typed without one is not verified, which is worth knowing before trusting it.
    // ###########################################################################################
    public static class ContributorHistory
    {
        // The most submissions the Contributor view lists - the Boards screen's limit. Beyond it
        // the counts still say how many there are.
        public const int ListedSubmissions = 50;

        public static async Task<ReviewContributorFacts> BuildAsync(
            SubmissionRecord submission,
            ISubmissionStore store,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(accounts);

            AccountRecord? account = submission.AccountId is long accountId
                ? await accounts.FindByIdAsync(accountId, cancellationToken)
                : null;

            IReadOnlyList<ContributorSubmission> all = await store.GetContributorSubmissionsAsync(
                submission.AccountId, submission.ContactEmail, cancellationToken);

            IReadOnlyList<ContributorSubmission> listed = ContributorHistory.Listed(submission, all);

            // What each counted submission's board has reached production, and which a BETA
            // rollback returned - the two facts its contributor-facing word needs. The boards of
            // EVERY counted submission, not only the listed ones: PublishedToStable counts all of
            // them, and a record longer than the list would otherwise count a stable one as BETA.
            //
            // *** ONE QUERY, NOT ONE PER BOARD (code review, 2026-10-01). *** This asked
            // FindBoardAsync for each distinct board in a loop, so opening any submission cost as
            // many sequential round-trips as its contributor has boards - and a submission is opened
            // on every click in the queue. The boards table is one small row per board, so the whole
            // list is read once and looked up here; GetBetaReturnsAsync below is one query for the
            // same reason.
            var productionPublished = new Dictionary<string, DateTimeOffset?>(StringComparer.Ordinal);

            if (listed.Count > 0)
            {
                Dictionary<string, BoardRecord> boards = (await store.ListBoardsAsync(cancellationToken))
                    .GroupBy(board => board.BoardId, StringComparer.Ordinal)
                    .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);

                IEnumerable<string> countedBoards = all
                    .Where(other => other.Id != submission.Id && ContributorHistory.Counted(other))
                    .Select(other => other.BoardId)
                    .Distinct(StringComparer.Ordinal);

                foreach (string boardId in countedBoards)
                    productionPublished[boardId] = boards.TryGetValue(boardId, out BoardRecord? board) ? board.ProductionPublishedUtc : null;
            }

            IReadOnlyDictionary<long, DateTimeOffset> returns = listed.Count == 0
                ? new Dictionary<long, DateTimeOffset>()
                : await store.GetBetaReturnsAsync(listed.Select(other => other.Id).ToList(), cancellationToken);

            return ContributorHistory.Facts(submission, all, account) with
            {
                SignedIn = submission.AccountId is not null,
                AccountCreatedUtc = account?.CreatedUtc,
                Submissions = ContributorHistory.Entries(listed, productionPublished, returns),
                PublishedToStable = ContributorHistory.PublishedToStable(submission, all, productionPublished)
            };
        }

        // ###########################################################################################
        // How many of the contributor's other PUBLISHED ('merged') submissions have reached the
        // stable source (owner request, 2026-10-01: "[1] published to stable"). Published counts
        // the BETA data and the stable source together, since the database keeps 'merged' for
        // both; this is the part of it whose board was promoted after the decision - the same
        // rule as each listed submission's word (ProductionPromotionRules.ContributorFacingState),
        // so the count and the list below it cannot disagree. `productionPublished` names when
        // each board last reached production; a board missing from it has not.
        // ###########################################################################################
        public static int PublishedToStable(
            SubmissionRecord submission,
            IEnumerable<ContributorSubmission> all,
            IReadOnlyDictionary<string, DateTimeOffset?> productionPublished)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(all);
            ArgumentNullException.ThrowIfNull(productionPublished);

            return all.Count(other =>
                other.Id != submission.Id &&
                other.State == SubmissionState.Merged &&
                ProductionPromotionRules.ContributorFacingState(
                    other.State,
                    other.DecidedUtc,
                    productionPublished.GetValueOrDefault(other.BoardId)) == ProductionPromotionRules.PublishedState);
        }

        // ###########################################################################################
        // The facts, from the contributor's submissions - THIS one left out, since the maintainer is
        // looking at it. Published is 'merged' (the BETA data, and production, which keeps that
        // state); a rejection counts only when a maintainer made it; a submission replaced by the
        // contributor's newer one ('withdrawn') and one that never finished sending are not counted.
        // Each count is one of Counted's kinds, which is what keeps Listed in step with them.
        // ###########################################################################################
        public static ReviewContributorFacts Facts(
            SubmissionRecord submission,
            IEnumerable<ContributorSubmission> all,
            AccountRecord? account)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(all);

            List<ContributorSubmission> others = all.Where(other => other.Id != submission.Id).ToList();

            int Count(Func<ContributorSubmission, bool> which) => others.Count(which);

            string? email = account?.Email ?? submission.ContactEmail?.Trim();
            string? name = account is not null && !string.IsNullOrWhiteSpace(account.DisplayName)
                ? account.DisplayName.Trim()
                : null;

            return new ReviewContributorFacts(
                Email: string.IsNullOrWhiteSpace(email) ? null : email,
                Name: name,
                Published: Count(other => other.State == SubmissionState.Merged),
                Waiting: Count(other => other.State is SubmissionState.Pending or SubmissionState.Approved),
                ChangesRequested: Count(other => other.State == SubmissionState.ChangesRequested),
                Rejected: Count(other => other.State == SubmissionState.Rejected && other.DecidedByMaintainer));
        }

        // ###########################################################################################
        // Whether a submission is one the counts count: published, waiting, sent back, or turned
        // down BY A MAINTAINER. Not one the automatic checks refused (it never reached anyone), not
        // one replaced by the contributor's own newer submission (an earlier copy of the same work),
        // and not one that never finished sending.
        // ###########################################################################################
        public static bool Counted(ContributorSubmission submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return submission.State is SubmissionState.Merged or SubmissionState.Pending or SubmissionState.Approved or SubmissionState.ChangesRequested ||
                (submission.State == SubmissionState.Rejected && submission.DecidedByMaintainer);
        }

        // The contributor's OTHER counted submissions, newest first, at most ListedSubmissions.
        public static IReadOnlyList<ContributorSubmission> Listed(SubmissionRecord submission, IEnumerable<ContributorSubmission> all)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(all);

            return all
                .Where(other => other.Id != submission.Id && ContributorHistory.Counted(other))
                .OrderByDescending(other => other.Id)
                .Take(ContributorHistory.ListedSubmissions)
                .ToList();
        }

        // ###########################################################################################
        // The listed submissions as the Maintainer tab reads them, each in the word its contributor
        // is told - "published" once its board reached production after it was merged, "returned"
        // while a BETA rollback's return stands (ProductionPromotionRules.ContributorFacingState).
        // ###########################################################################################
        public static IReadOnlyList<ContributorSubmissionEntry> Entries(
            IEnumerable<ContributorSubmission> listed,
            IReadOnlyDictionary<string, DateTimeOffset?> productionPublished,
            IReadOnlyDictionary<long, DateTimeOffset> returns)
        {
            ArgumentNullException.ThrowIfNull(listed);
            ArgumentNullException.ThrowIfNull(productionPublished);
            ArgumentNullException.ThrowIfNull(returns);

            return listed
                .Select(other => new ContributorSubmissionEntry(
                    other.Id,
                    other.BoardId,
                    other.Summary,
                    ProductionPromotionRules.ContributorFacingState(
                        other.State,
                        other.DecidedUtc,
                        productionPublished.GetValueOrDefault(other.BoardId),
                        returns.TryGetValue(other.Id, out DateTimeOffset returned) ? returned : null),
                    other.CreatedUtc,
                    other.DecidedUtc,
                    other.DecisionComment))
                .ToList();
        }
    }
}
