using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A NEW BOARD'S CREATOR BECOMES ITS FIRST MAINTAINER (owner request, 2026-10-13: "I think we
    // need to (automatically) make the contributor of a NEW board then first/initial board
    // maintainer also").
    //
    // "Create board" in CRT asks the contributor to accept being the board's maintainer before
    // anything is created (NewBoardMaintainerWindow, owner request 2026-09-24). Until now nothing came
    // of that "Accept": the administrator had to invite the person by hand afterwards. This makes it
    // true, at the moment the board comes into being - the approval that first publishes it to BETA
    // (ApproveOutcome.CreatedBoard):
    //
    //   - the board was sent SIGNED IN - that account is put in the board's pool, as the
    //     administrator's "Add as maintainer" would, and mailed that it now maintains the board;
    //   - it was sent with a typed address that has NO account - the address is mailed an
    //     invitation, as the administrator's "Send invitation" would, worded for the board's
    //     creator. Accepting it creates the account and the pool row
    //     (MaintainerInvitationFlows.AcceptAsync); the code reaching the mailbox proves the address;
    //   - a typed address that DOES have an account - nobody is made maintainer (code review,
    //     2026-10-13). Nothing proves the sender owns that address (contributing needs no
    //     account, and the address is not checked), and the invitation, which would prove it, is
    //     refused for an address with an account. The approver is told, to settle by hand;
    //   - an account that may not maintain (unverified or locked - MaintainerAssignmentRules) is
    //     left alone, and the approver is told so, to settle by hand.
    //
    // *** THE APPROVER IS THE ONE WHO DID IT. *** The pool row and the audit rows name the account
    // that approved the board - the grant follows from that approval - in the same words the
    // administrator's own grant and invitation write, so the board's history reads the same.
    //
    // *** AFTER THE PUBLISH, AND IT NEVER FAILS IT. *** The board is already written: the caller
    // runs this guarded, and a fault costs the maintainer, never the approval (ReviewEndpoints).
    //
    // Pure, taking its stores and mailer as arguments, like every other flow.
    // ###########################################################################################
    public static class NewBoardMaintainerFlow
    {
        // ###########################################################################################
        // Makes the contributor of `submission` - a NEW board just published - its maintainer, or
        // invites them to be. Returns the sentence the approver reads.
        // ###########################################################################################
        public static async Task<string> MakeCreatorMaintainerAsync(
            SubmissionRecord submission,
            ReviewAccess approver,
            IAccountStore accounts,
            ISubmissionStore submissions,
            IEmailSender mailer,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(approver);
            ArgumentNullException.ThrowIfNull(accounts);
            ArgumentNullException.ThrowIfNull(submissions);
            ArgumentNullException.ThrowIfNull(mailer);

            string boardId = submission.BoardId;
            string boardName = await NewBoardMaintainerFlow.BoardNameAsync(boardId, submissions, cancellationToken);

            AccountRecord? account = submission.AccountId is long accountId
                ? await accounts.FindByIdAsync(accountId, cancellationToken)
                : null;

            if (account is null)
            {
                // A typed address belonging to an account: not proven to be the sender's.
                string typed = submission.ContactEmail?.Trim() ?? string.Empty;

                if (AccountRules.IsPlausibleEmail(typed) &&
                    await accounts.FindByNormalisedEmailAsync(AccountRules.NormaliseEmail(typed), cancellationToken) is not null)
                {
                    return NewBoardMaintainerWording.TypedAddressHasAccount(typed);
                }

                return await NewBoardMaintainerFlow.InviteAsync(submission, boardId, boardName, approver, accounts, mailer, now, cancellationToken);
            }

            if (account.IsLocked)
                return NewBoardMaintainerWording.Locked(account.Email);

            if (!account.IsVerified)
                return NewBoardMaintainerWording.Unverified(account.Email);

            bool already = (await accounts.GetMaintainersOfBoardAsync(boardId, cancellationToken))
                .Any(maintainer => maintainer.AccountId == account.Id);

            if (already)
                return NewBoardMaintainerWording.AlreadyMaintains(account.DisplayName, account.Email);

            await accounts.AddMaintainerAsync(boardId, account.Id, approver.Account.Id, now, cancellationToken);

            // The administrator's own grant writes these same words (MaintainerAssignmentFlows), which
            // the board's history reads.
            await accounts.WriteAuditAsync(
                new AuditEntry(
                    approver.Account.Id,
                    approver.Account.Email,
                    MaintainerAssignmentFlows.GrantedAction,
                    boardId,
                    $"account {account.Id} ({account.Email})",
                    now),
                cancellationToken);

            await mailer.SendAsync(EmailTemplates.NewBoardMaintainerAdded(account.Email, account.DisplayName, boardName), cancellationToken);

            return NewBoardMaintainerWording.Added(account.DisplayName, account.Email);
        }

        // No account: an invitation to the typed address, issued exactly as the administrator's own
        // (MaintainerInvitationFlows.IssueAsync, the one copy) - an earlier open one for the same
        // board and address stops working, the code is kept only as its hash, and it works for the
        // same 14 days. The board's row exists: the submission created it.
        private static async Task<string> InviteAsync(
            SubmissionRecord submission,
            string boardId,
            string boardName,
            ReviewAccess approver,
            IAccountStore accounts,
            IEmailSender mailer,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            string address = submission.ContactEmail?.Trim() ?? string.Empty;

            if (!AccountRules.IsPlausibleEmail(address))
                return NewBoardMaintainerWording.NoAddress;

            await MaintainerInvitationFlows.IssueAsync(
                boardId, boardName, address, approver, toItsCreator: true, accounts, mailer, now, cancellationToken);

            return NewBoardMaintainerWording.Invited(address);
        }

        // "Commodore / C64 / 250407", from the board's row - which the submission created - and the
        // id itself without one.
        private static async Task<string> BoardNameAsync(string boardId, ISubmissionStore submissions, CancellationToken cancellationToken)
        {
            (string Manufacturer, string Hardware, string Board)? identity =
                await MaintainerAssignmentFlows.FindBoardAsync(boardId, [], submissions, cancellationToken);

            return identity is { } known
                ? MaintainerInvitationRules.BoardDisplayName(known.Manufacturer, known.Hardware, known.Board)
                : boardId;
        }
    }

    // ###########################################################################################
    // What the approver is told, under "Published to BETA". Pure, so the sentences are tested. They
    // name the contributor or leave the pronoun out - never "their", "them" or "they" (owner rule).
    // ###########################################################################################
    public static class NewBoardMaintainerWording
    {
        public const string NoAddress =
            "Nobody was made this new board's maintainer: the submission has no email address to write to.";

        public const string Failed =
            "This new board's creator could not be made its maintainer - the server's log says why. Add the maintainer under " +
            MaintainerScreenWording.AccountQuoted + ", \"Maintainers\".";

        public static string Added(string name, string email) =>
            $"{NewBoardMaintainerWording.Who(name, email)}, who created this new board, is now its maintainer.";

        public static string AlreadyMaintains(string name, string email) =>
            $"{NewBoardMaintainerWording.Who(name, email)}, who created this new board, already maintains it.";

        public static string TypedAddressHasAccount(string email) =>
            $"This new board was sent without signing in, giving {email} - an address that has an account. Nobody was made " +
            "its maintainer: nothing proves the sender owns that address. If that account's owner created the board, add it under " +
            MaintainerScreenWording.AccountQuoted + ", \"Maintainers\".";

        public static string Invited(string email) =>
            $"{email}, who created this new board, has no account, so an invitation to become its maintainer was mailed there.";

        public static string Locked(string email) =>
            $"{email} created this new board, but that account is locked, so it was not made the board's maintainer.";

        public static string Unverified(string email) =>
            $"{email} created this new board, but that account's address is not verified yet, so it was not made the board's " +
            $"maintainer. Add it under {MaintainerScreenWording.AccountQuoted}, \"Maintainers\", once it is.";

        private static string Who(string name, string email) =>
            string.IsNullOrWhiteSpace(name) ? email : $"{name.Trim()} ({email})";
    }
}
