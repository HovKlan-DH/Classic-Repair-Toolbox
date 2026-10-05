using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHO IS TOLD that a submission is waiting (Phase 6 tasks 2, 3 and 11; 2026-09-25).
    //
    // The rule is the mirror of ApprovalRules, stated once for the people rather than for the
    // request: whoever must approve is told. The system's maintainers ordinarily; the maintainers AND
    // the administrators when a shared file changes (both must approve - owner decision,
    // 2026-09-25); the administrators alone when the system has no maintainers. "A new system routes
    // to the administrator, always" falls out of the last case, since a system nobody has reviewed
    // yet has an empty pool.
    //
    // Each comes with the name on their account, to greet them by (2026-10-03).
    //
    // Pure over the stores, so it is tested with the fakes.
    // ###########################################################################################
    public static class SubmissionRouting
    {
        public static async Task<IReadOnlyList<MailRecipient>> RecipientsForAsync(
            SubmissionRecord submission,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(accounts);

            IReadOnlyList<MaintainerRecord> pool =
                await accounts.GetMaintainersOfSystemAsync(submission.SystemId, cancellationToken);

            // Only maintainers who can approve as maintainers decide WHICH approvals are needed - the
            // rule the approval itself counts by (ReviewAuthority.CanGiveMaintainerApproval), so who
            // is asked and who is waited for agree.
            bool hasMaintainers = pool.Any(ReviewAuthority.CanGiveMaintainerApproval);

            IReadOnlyList<ApproverRole> required = ApprovalRules.Required(submission.TouchesSharedFiles, hasMaintainers);

            // Ordinary: any one approval, and the system's maintainers are the ones to ask - or the
            // administrators when there are none. An ADMINISTRATOR in the pool (2026-10-05) is one of
            // the system's maintainers here: named as its maintainer, so told like one - and once,
            // however many roles they hold (Distinct, and SubmissionNotifier's own).
            if (required.Count == 0)
            {
                List<MailRecipient> maintainers = pool
                    .Where(SubmissionRouting.CanBeWrittenTo)
                    .Select(maintainer => new MailRecipient(maintainer.Email, maintainer.DisplayName))
                    .ToList();

                return maintainers.Count > 0
                    ? SubmissionRouting.Distinct(maintainers)
                    : await SubmissionRouting.AdministratorAddressesAsync(accounts, cancellationToken);
            }

            return await SubmissionRouting.RecipientsForRolesAsync(required, submission.SystemId, accounts, cancellationToken);
        }

        // ###########################################################################################
        // Everybody who approves in these roles for this system - for "your approval is needed"
        // once the other half has approved. De-duplicated, in role order.
        // ###########################################################################################
        public static async Task<IReadOnlyList<MailRecipient>> RecipientsForRolesAsync(
            IEnumerable<ApproverRole> roles,
            string systemId,
            IAccountStore accounts,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(roles);
            ArgumentNullException.ThrowIfNull(accounts);

            var recipients = new List<MailRecipient>();

            foreach (ApproverRole role in roles.Distinct())
            {
                if (role == ApproverRole.Maintainer)
                {
                    recipients.AddRange((await accounts.GetMaintainersOfSystemAsync(systemId, cancellationToken))
                        .Where(ReviewAuthority.CanGiveMaintainerApproval)
                        .Select(maintainer => new MailRecipient(maintainer.Email, maintainer.DisplayName)));
                }
                else
                {
                    recipients.AddRange(await SubmissionRouting.AdministratorAddressesAsync(accounts, cancellationToken));
                }
            }

            return SubmissionRouting.Distinct(recipients);
        }

        // One mail per address, in the order first named.
        private static IReadOnlyList<MailRecipient> Distinct(IEnumerable<MailRecipient> recipients) =>
            recipients
                .GroupBy(recipient => recipient.Email, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .ToList();

        // A pool member with a proven address who can still sign in - an administrator included.
        // The administrators' own rule, below.
        private static bool CanBeWrittenTo(MaintainerRecord maintainer) =>
            maintainer.IsVerified && !maintainer.IsLocked;

        // A locked administrator is not somebody to write to; an unverified one has no proven
        // address.
        private static async Task<IReadOnlyList<MailRecipient>> AdministratorAddressesAsync(
            IAccountStore accounts,
            CancellationToken cancellationToken) =>
            (await accounts.GetAdministratorsAsync(cancellationToken))
                .Where(account => account.IsVerified && !account.IsLocked)
                .Select(account => new MailRecipient(account.Email, account.DisplayName))
                .ToList();
    }
}
