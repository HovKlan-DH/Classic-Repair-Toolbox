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
    // ###########################################################################################
    public static class ContributorHistory
    {
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

            return ContributorHistory.Facts(submission, all, account);
        }

        // ###########################################################################################
        // The facts, from the contributor's submissions - THIS one left out, since the maintainer is
        // looking at it. Published is 'merged' (the BETA data, and production, which keeps that
        // state); a rejection counts only when a maintainer made it; a submission replaced by the
        // contributor's newer one ('withdrawn') and one that never finished sending are not counted.
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
    }
}
