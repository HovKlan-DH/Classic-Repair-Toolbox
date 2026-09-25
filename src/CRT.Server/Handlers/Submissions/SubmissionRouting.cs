using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHO IS TOLD that a submission is waiting (Phase 6 tasks 2, 3 and 11; 2026-09-25).
    //
    // The rule is the mirror of ApprovalRules, stated once for the people rather than for the
    // request: whoever must approve is told. The system's reviewers ordinarily; the reviewers AND
    // the administrators when a shared file changes (both must approve - maintainer decision,
    // 2026-09-25); the administrators alone when the system has no reviewers. "A new system routes
    // to the administrator, always" falls out of the last case, since a system nobody has reviewed
    // yet has an empty pool.
    //
    // Pure over the stores, so it is tested with the fakes.
    // ###########################################################################################
    public static class SubmissionRouting
    {
        public static async Task<IReadOnlyList<string>> RecipientsForAsync(
            SubmissionRecord submission,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(accounts);

            // Only reviewers who can approve as reviewers - the rule the approval itself counts by
            // (ReviewAuthority.CanGiveReviewerApproval), so who is asked and who is waited for agree.
            IReadOnlyList<ReviewerRecord> reviewers =
                (await accounts.GetReviewersOfSystemAsync(submission.SystemId, cancellationToken))
                    .Where(ReviewAuthority.CanGiveReviewerApproval)
                    .ToList();

            IReadOnlyList<ApproverRole> required = ApprovalRules.Required(submission.TouchesSharedFiles, reviewers.Count > 0);

            // Ordinary: any one approval, and the reviewers are the ones to ask - or the
            // administrators when there are none.
            if (required.Count == 0)
            {
                return reviewers.Count > 0
                    ? reviewers.Select(reviewer => reviewer.Email).ToList()
                    : await SubmissionRouting.AdministratorAddressesAsync(accounts, cancellationToken);
            }

            return await SubmissionRouting.RecipientsForRolesAsync(required, submission.SystemId, accounts, cancellationToken);
        }

        // ###########################################################################################
        // Everybody who approves in these roles for this system - for "your approval is needed"
        // once the other half has approved. De-duplicated, in role order.
        // ###########################################################################################
        public static async Task<IReadOnlyList<string>> RecipientsForRolesAsync(
            IEnumerable<ApproverRole> roles,
            string systemId,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(roles);
            ArgumentNullException.ThrowIfNull(accounts);

            var addresses = new List<string>();

            foreach (ApproverRole role in roles.Distinct())
            {
                if (role == ApproverRole.Reviewer)
                {
                    addresses.AddRange((await accounts.GetReviewersOfSystemAsync(systemId, cancellationToken))
                        .Where(ReviewAuthority.CanGiveReviewerApproval)
                        .Select(reviewer => reviewer.Email));
                }
                else
                {
                    addresses.AddRange(await SubmissionRouting.AdministratorAddressesAsync(accounts, cancellationToken));
                }
            }

            return addresses.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // A locked administrator is not somebody to write to; an unverified one has no proven
        // address.
        private static async Task<IReadOnlyList<string>> AdministratorAddressesAsync(
            IAccountStore accounts,
            CancellationToken cancellationToken) =>
            (await accounts.GetAdministratorsAsync(cancellationToken))
                .Where(account => account.IsVerified && !account.IsLocked)
                .Select(account => account.Email)
                .ToList();
    }
}
