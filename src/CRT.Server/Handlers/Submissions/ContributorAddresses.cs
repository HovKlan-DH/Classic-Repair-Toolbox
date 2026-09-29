using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHERE A CONTRIBUTOR IS WRITTEN TO (code review, 2026-09-29).
    //
    // A submission sent while SIGNED IN carries the account, and usually no contact address of its
    // own - so everything reading SubmissionRecord.ContactEmail alone found nobody for exactly the
    // account-holding contributors: the production panel listed "(no contact address)", and the
    // push-back and "now in production" mails were skipped for them, while the Systems screen beside
    // them did name the account's address. This is the one rule for all of them, the same one
    // ContributorHistory already applied: the account's address when there is an account, the
    // contact address otherwise.
    // ###########################################################################################
    public static class ContributorAddresses
    {
        // The address for one submission, given its account (null when it has none, or the account
        // is gone). Blank when neither says anything.
        public static string AddressOf(SubmissionRecord submission, AccountRecord? account)
        {
            ArgumentNullException.ThrowIfNull(submission);

            string? fromAccount = account?.Email?.Trim();

            return !string.IsNullOrWhiteSpace(fromAccount)
                ? fromAccount
                : submission.ContactEmail?.Trim() ?? string.Empty;
        }

        // The address for each of `submissions`, by submission id - one account lookup per account.
        public static async Task<IReadOnlyDictionary<long, string>> ResolveAsync(
            IAccountStore accounts,
            IEnumerable<SubmissionRecord> submissions,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(submissions);

            List<SubmissionRecord> list = submissions.ToList();
            var byAccount = new Dictionary<long, AccountRecord?>();

            foreach (long accountId in list.Select(submission => submission.AccountId).OfType<long>().Distinct())
                byAccount[accountId] = await accounts.FindByIdAsync(accountId, cancellationToken);

            return list.ToDictionary(
                submission => submission.Id,
                submission => ContributorAddresses.AddressOf(
                    submission,
                    submission.AccountId is long id ? byAccount.GetValueOrDefault(id) : null));
        }
    }
}
