namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // LOCKING AN ACCOUNT FROM THE WEBSITE (owner decision, 2026-11-27: "yes, via web page where I as admin can do this" - by SQL
    // alone until then, INSTALLING.md "Accounts and sessions"). The administrator's alone, on "Your collections"
    // (CollectionPages.List). Pure, as the other flows are: the store handed in.
    //
    //   - A locked account cannot sign in or use a session (AccountFlows), its collections go offline (CollectionFlows
    //     .FindShownAsync - to everybody but the administrator) and its comments are shown to the administrator alone
    //     (CollectionCommentFlows.ShownTo) - so locking a spammer takes the spam down with it, and unlocking puts it back.
    //   - LOCKING AND UNLOCKING BOTH END EVERY SESSION THE ACCOUNT HAS, so nothing it held comes back to life by itself: a
    //     mistaken lock undone asks it to sign in again. On unlock too (code review, 2026-11-27, Security): an account locked
    //     by SQL kept its sessions - refused only while locked - and an Unlock here brought them all back, one stolen before
    //     the lock with them.
    //   - AN ADMINISTRATOR'S ACCOUNT IS NEVER LOCKED OR UNLOCKED HERE: that stays a hand's work in SQL, so no click locks the
    //     only one out - nor is one listed under "Locked accounts" (ListLockedAsync).
    //   - Every lock and unlock is audited, under the account it names - the revoke and the audit carried out once the lock is
    //     written, whatever the request does meanwhile.
    // ###########################################################################################
    public static class AccountLockFlows
    {
        public const string OnlyTheAdministrator = "Only the administrator locks or unlocks an account.";

        public const string AdministratorsStay = "An administrator's account is not locked or unlocked from here.";

        public const string NoSuchAccount = "There is no such account.";

        public const string LockedAction = "account.locked";

        public const string UnlockedAction = "account.unlocked";

        public static async Task<AccountLockOutcome> SetLockedAsync(
            AccountRecord? administrator,
            long accountId,
            bool locked,
            IAccountStore accounts,
            DateTimeOffset now,
            CancellationToken cancellationToken,
            ILogger? logger = null)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (administrator is not { IsAdministrator: true })
                return AccountLockOutcome.Refused(AccountLockVerdict.Forbidden, AccountLockFlows.OnlyTheAdministrator);

            AccountRecord? account = await accounts.FindByIdAsync(accountId, cancellationToken);

            if (account is null)
                return AccountLockOutcome.Refused(AccountLockVerdict.NotFound, AccountLockFlows.NoSuchAccount);

            if (account.IsAdministrator)
                return AccountLockOutcome.Refused(AccountLockVerdict.Refused, AccountLockFlows.AdministratorsStay);

            // Already so - a second press, or two administrators at once - is done, and nothing is written again.
            if (account.IsLocked == locked)
                return AccountLockOutcome.Done;

            await accounts.SetLockedAsync(account.Id, locked, cancellationToken);

            // Written: what follows is carried out whatever the request does now - a lock left with its sessions alive, or
            // unaudited, would be half a lock.
            await accounts.RevokeAllSessionsAsync(account.Id, locked ? "locked" : "unlocked", now, CancellationToken.None);

            try
            {
                await accounts.WriteAuditAsync(
                    new AuditEntry(
                        administrator.Id,
                        administrator.Email,
                        locked ? AccountLockFlows.LockedAction : AccountLockFlows.UnlockedAction,
                        $"account:{account.Id}",
                        account.DisplayName,
                        now),
                    CancellationToken.None);
            }
            catch (Exception ex)
            {
                // The lock stands - an audit that failed must not turn it into an error page.
                logger?.LogError(ex, "Locking account {Id} could not be audited.", account.Id);
            }

            return AccountLockOutcome.Done;
        }

        // The accounts locked, for the administrator to unlock one - nobody else is told of any. An administrator's account
        // locked by SQL is left out: it is not unlocked from here either.
        public static async Task<IReadOnlyList<AccountRecord>> ListLockedAsync(AccountRecord? viewer, IAccountStore accounts, CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            return viewer is { IsAdministrator: true }
                ? (await accounts.ListLockedAccountsAsync(cancellationToken)).Where(account => !account.IsAdministrator).ToList()
                : [];
        }

        // The account the administrator asked to lock (?lock=<id>), to be asked about first - only one SetLockedAsync would
        // lock: none for anybody but the administrator, for another administrator's, or for one locked already.
        public static async Task<AccountRecord?> ToAskAboutLockingAsync(AccountRecord? viewer, long? accountId, IAccountStore accounts, CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(accounts);

            if (viewer is not { IsAdministrator: true } || accountId is not { } id)
                return null;

            return await accounts.FindByIdAsync(id, cancellationToken) is { IsAdministrator: false, IsLocked: false } account ? account : null;
        }
    }

    public enum AccountLockVerdict
    {
        Done,
        Forbidden,
        NotFound,
        Refused
    }

    public sealed record AccountLockOutcome(AccountLockVerdict Verdict, string? Error)
    {
        public bool IsDone => this.Verdict == AccountLockVerdict.Done;

        public static readonly AccountLockOutcome Done = new(AccountLockVerdict.Done, null);

        public static AccountLockOutcome Refused(AccountLockVerdict verdict, string error) => new(verdict, error);
    }
}
