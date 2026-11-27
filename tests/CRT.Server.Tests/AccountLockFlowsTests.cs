using CRT.Server.Handlers.Accounts;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Locking an account from the website (AccountLockFlows; owner decision, 2026-11-27: "yes, via web page where I as admin can
    // do this"): the administrator alone, never an administrator's account, every session ended, every change audited.
    // ###########################################################################################
    public sealed class AccountLockFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 11, 27, 12, 0, 0, TimeSpan.Zero);

        private readonly FakeAccountStore thisAccounts = new();

        private static readonly AccountRecord Spammer = Account(2, "spam@example.com", "Spammer");

        private static readonly AccountRecord Bo = Account(3, "bo@example.com", "Bo");

        private static readonly AccountRecord Administrator = Account(9, "admin@example.com", "Dennis@CRT", administrator: true);

        private static readonly AccountRecord OtherAdministrator = Account(10, "other@example.com", "Other", administrator: true);

        public AccountLockFlowsTests()
        {
            foreach (AccountRecord account in new[] { Spammer, Bo, Administrator, OtherAdministrator })
                this.thisAccounts.Accounts[account.Id] = account;
        }

        private static AccountRecord Account(long id, string email, string name, bool administrator = false) =>
            new(id, email, email, "hash", name, true, administrator, IsLocked: false, Now.AddDays(-10), null);

        private Task<AccountLockOutcome> SetAsync(AccountRecord? who, long accountId, bool locked) =>
            AccountLockFlows.SetLockedAsync(who, accountId, locked, this.thisAccounts, Now, CancellationToken.None);

        private async Task<long> SessionOfAsync(AccountRecord account) =>
            await this.thisAccounts.CreateSessionAsync(new NewSession(account.Id, Guid.NewGuid().ToString("N"), Now.AddHours(-1), Now.AddDays(30), null, null));

        [Fact]
        public async Task The_administrator_locks_an_account_its_sessions_end_and_it_is_audited()
        {
            long session = await this.SessionOfAsync(Spammer);
            long somebodyElses = await this.SessionOfAsync(Bo);

            AccountLockOutcome outcome = await this.SetAsync(Administrator, Spammer.Id, locked: true);

            Assert.True(outcome.IsDone);
            Assert.True(this.thisAccounts.Accounts[Spammer.Id].IsLocked);
            Assert.Equal("locked", this.thisAccounts.Sessions[session].RevokedReason);
            Assert.Null(this.thisAccounts.Sessions[somebodyElses].RevokedUtc);

            AuditEntry audit = Assert.Single(this.thisAccounts.Audit);
            Assert.Equal((Administrator.Id, AccountLockFlows.LockedAction, $"account:{Spammer.Id}", "Spammer"), (audit.ActorAccountId, audit.Action, audit.Subject, audit.Detail));
        }

        // Unlocking brings back no session: a mistaken lock undone asks the account to sign in again.
        [Fact]
        public async Task Unlocking_lets_the_account_in_again_and_brings_back_no_session()
        {
            long session = await this.SessionOfAsync(Spammer);
            await this.SetAsync(Administrator, Spammer.Id, locked: true);

            AccountLockOutcome outcome = await this.SetAsync(Administrator, Spammer.Id, locked: false);

            Assert.True(outcome.IsDone);
            Assert.False(this.thisAccounts.Accounts[Spammer.Id].IsLocked);
            Assert.NotNull(this.thisAccounts.Sessions[session].RevokedUtc);
            Assert.Equal([AccountLockFlows.LockedAction, AccountLockFlows.UnlockedAction], this.thisAccounts.Audit.Select(entry => entry.Action));
        }

        // An account locked by SQL - INSTALLING.md's way until 15.0.0 - keeps its sessions, refused only while it is locked: an
        // Unlock here must not bring them back to life, one stolen before the lock with them (code review, 2026-11-27, Security).
        [Fact]
        public async Task Unlocking_an_account_locked_by_hand_ends_the_sessions_it_kept()
        {
            long session = await this.SessionOfAsync(Spammer);
            this.thisAccounts.Accounts[Spammer.Id] = Spammer with { IsLocked = true };

            AccountLockOutcome outcome = await this.SetAsync(Administrator, Spammer.Id, locked: false);

            Assert.True(outcome.IsDone);
            Assert.False(this.thisAccounts.Accounts[Spammer.Id].IsLocked);
            Assert.Equal("unlocked", this.thisAccounts.Sessions[session].RevokedReason);
        }

        // Once written, the lock's revoke and audit are carried out though the request goes away - half a lock is none.
        [Fact]
        public async Task A_lock_written_is_audited_and_its_sessions_end_though_the_request_is_cancelled_after_it()
        {
            long session = await this.SessionOfAsync(Spammer);
            using var cancelled = new CancellationTokenSource();
            this.thisAccounts.AfterSetLocked = cancelled.Cancel;

            AccountLockOutcome outcome = await AccountLockFlows.SetLockedAsync(Administrator, Spammer.Id, true, this.thisAccounts, Now, cancelled.Token);

            Assert.True(outcome.IsDone);
            Assert.NotNull(this.thisAccounts.Sessions[session].RevokedUtc);
            Assert.Single(this.thisAccounts.Audit);
        }

        [Fact]
        public async Task Nobody_but_the_administrator_locks_or_unlocks()
        {
            Assert.Equal(AccountLockVerdict.Forbidden, (await this.SetAsync(null, Spammer.Id, locked: true)).Verdict);
            Assert.Equal(AccountLockVerdict.Forbidden, (await this.SetAsync(Bo, Spammer.Id, locked: true)).Verdict);

            Assert.False(this.thisAccounts.Accounts[Spammer.Id].IsLocked);
            Assert.Empty(this.thisAccounts.Audit);
        }

        // No click locks an administrator out - not another one, not the one clicking: that stays a hand's work in SQL.
        [Fact]
        public async Task An_administrator_s_account_is_never_locked_from_here()
        {
            AccountLockOutcome other = await this.SetAsync(Administrator, OtherAdministrator.Id, locked: true);
            AccountLockOutcome self = await this.SetAsync(Administrator, Administrator.Id, locked: true);

            Assert.Equal((AccountLockVerdict.Refused, AccountLockFlows.AdministratorsStay), (other.Verdict, other.Error));
            Assert.Equal(AccountLockVerdict.Refused, self.Verdict);
            Assert.False(this.thisAccounts.Accounts[OtherAdministrator.Id].IsLocked);
            Assert.False(this.thisAccounts.Accounts[Administrator.Id].IsLocked);
        }

        [Fact]
        public async Task An_account_there_is_none_of_is_not_found() =>
            Assert.Equal(AccountLockVerdict.NotFound, (await this.SetAsync(Administrator, 404, locked: true)).Verdict);

        // A second press, or two administrators at once, is done - and nothing is written or audited again.
        [Fact]
        public async Task Locking_an_account_locked_already_writes_nothing_again()
        {
            await this.SetAsync(Administrator, Spammer.Id, locked: true);

            AccountLockOutcome again = await this.SetAsync(Administrator, Spammer.Id, locked: true);

            Assert.True(again.IsDone);
            Assert.Single(this.thisAccounts.Audit);
        }

        [Fact]
        public async Task The_lock_stands_when_its_audit_cannot_be_written()
        {
            this.thisAccounts.FailAuditWrites = true;

            Assert.True((await this.SetAsync(Administrator, Spammer.Id, locked: true)).IsDone);
            Assert.True(this.thisAccounts.Accounts[Spammer.Id].IsLocked);
        }

        [Fact]
        public async Task The_accounts_locked_are_listed_to_the_administrator_alone()
        {
            await this.SetAsync(Administrator, Spammer.Id, locked: true);

            Assert.Equal(["Spammer"], (await AccountLockFlows.ListLockedAsync(Administrator, this.thisAccounts, CancellationToken.None)).Select(account => account.DisplayName));
            Assert.Empty(await AccountLockFlows.ListLockedAsync(Bo, this.thisAccounts, CancellationToken.None));
            Assert.Empty(await AccountLockFlows.ListLockedAsync(null, this.thisAccounts, CancellationToken.None));
        }

        // An administrator's account locked by SQL is not unlocked from here, so it is not offered as if it were.
        [Fact]
        public async Task An_administrator_s_account_locked_by_hand_is_not_listed_to_unlock()
        {
            this.thisAccounts.Accounts[OtherAdministrator.Id] = OtherAdministrator with { IsLocked = true };
            await this.SetAsync(Administrator, Spammer.Id, locked: true);

            Assert.Equal(["Spammer"], (await AccountLockFlows.ListLockedAsync(Administrator, this.thisAccounts, CancellationToken.None)).Select(account => account.DisplayName));
        }

        // "?lock=<id>" asks only about an account SetLockedAsync would lock.
        [Fact]
        public async Task Only_an_account_the_administrator_may_lock_is_asked_about()
        {
            this.thisAccounts.Accounts[Bo.Id] = Bo with { IsLocked = true };

            Task<AccountRecord?> AskAsync(AccountRecord? who, long? id) => AccountLockFlows.ToAskAboutLockingAsync(who, id, this.thisAccounts, CancellationToken.None);

            Assert.Equal(Spammer.Id, (await AskAsync(Administrator, Spammer.Id))?.Id);
            Assert.Null(await AskAsync(Administrator, Bo.Id));                    // locked already
            Assert.Null(await AskAsync(Administrator, OtherAdministrator.Id));    // an administrator's
            Assert.Null(await AskAsync(Administrator, 404));                      // none
            Assert.Null(await AskAsync(Administrator, null));
            Assert.Null(await AskAsync(Spammer, Bo.Id));                          // not the administrator asking
            Assert.Null(await AskAsync(null, Spammer.Id));
        }
    }
}
