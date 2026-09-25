using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SubmissionRouting - who is told that a submission is waiting (Phase 6 tasks 2, 3
    // and 11).
    //
    // The rule mirrors ReviewAuthority for people rather than requests, and a mismatch between
    // the two is the failure worth pinning: a maintainer mailed about a submission they cannot
    // open, or - worse - nobody mailed at all for a submission only the administrator can
    // decide.
    // ###########################################################################################
    public sealed class SubmissionRoutingTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";

        private static SubmissionRecord Submission(bool touchesShared = false) =>
            new(
                Id: 1, SystemId: SubmissionRoutingTests.C64, AccountId: null, ContactEmail: "c@example.com",
                UploadTokenHash: "h", BaseRevision: "r1", State: SubmissionState.Pending, Summary: "x",
                FormatVersion: 1, CreatedUtc: SubmissionRoutingTests.Now, ExpiresUtc: null, DecidedUtc: null,
                DecisionComment: null, TouchesSharedFiles: touchesShared);

        private static async Task<long> AccountAsync(FakeAccountStore store, string email, bool admin = false, bool verified = true, bool locked = false)
        {
            long id = await store.CreateAccountAsync(new NewAccount(email, email, "hash", email, SubmissionRoutingTests.Now));
            store.Accounts[id] = store.Accounts[id] with { IsAdministrator = admin, IsVerified = verified, IsLocked = locked };
            return id;
        }

        [Fact]
        public async Task A_submission_goes_to_EVERY_maintainer_of_its_system()
        {
            // Phase 6 task 2: route to every maintainer of the pool; any one may act.
            var accounts = new FakeAccountStore();
            long anna = await SubmissionRoutingTests.AccountAsync(accounts, "anna@example.com");
            long bob = await SubmissionRoutingTests.AccountAsync(accounts, "bob@example.com");
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, anna));
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, bob));

            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(), accounts);

            Assert.Equal(["anna@example.com", "bob@example.com"], recipients);
        }

        [Fact]
        public async Task A_system_with_NO_maintainers_goes_to_the_administrators()
        {
            // Phase 6 task 3: "a new system routes to the administrator, always" - and so does a
            // shipped board nobody has been assigned to yet.
            var accounts = new FakeAccountStore();
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);
            await SubmissionRoutingTests.AccountAsync(accounts, "anna@example.com");

            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(), accounts);

            Assert.Equal(["admin@example.com"], recipients);
        }

        [Fact]
        public async Task A_submission_changing_SHARED_FILES_goes_to_the_maintainers_AND_the_administrators()
        {
            // Both must approve it (2026-09-25), so both are told.
            var accounts = new FakeAccountStore();
            long anna = await SubmissionRoutingTests.AccountAsync(accounts, "anna@example.com");
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, anna));

            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(touchesShared: true), accounts);

            Assert.Equal(["anna@example.com", "admin@example.com"], recipients);
        }

        [Fact]
        public async Task A_SHARED_FILES_submission_to_a_board_with_no_maintainers_goes_to_the_administrators_alone()
        {
            var accounts = new FakeAccountStore();
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);

            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(touchesShared: true), accounts);

            Assert.Equal(["admin@example.com"], recipients);
        }

        [Fact]
        public async Task The_other_half_of_an_approval_is_found_by_ROLE()
        {
            // After a maintainer approves, the administrators are the ones to tell - and the other way
            // round.
            var accounts = new FakeAccountStore();
            long anna = await SubmissionRoutingTests.AccountAsync(accounts, "anna@example.com");
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, anna));

            Assert.Equal(
                ["admin@example.com"],
                await SubmissionRouting.RecipientsForRolesAsync([ApproverRole.Administrator], SubmissionRoutingTests.C64, accounts));

            Assert.Equal(
                ["anna@example.com"],
                await SubmissionRouting.RecipientsForRolesAsync([ApproverRole.Maintainer], SubmissionRoutingTests.C64, accounts));
        }

        // ###########################################################################################
        // A pool row whose account cannot approve as a maintainer - made administrator by hand, locked
        // or unverified - is not "the board's maintainer" (code review, 2026-09-25). Routing counts
        // maintainers by the same rule the approval does (ReviewAuthority.CanGiveMaintainerApproval),
        // so a shared-file change goes to the administrators alone rather than waiting for a
        // maintainer half nobody can give, and nobody is mailed "as the maintainer" who cannot act as one.
        // ###########################################################################################
        [Fact]
        public async Task A_pool_row_that_cannot_approve_as_a_maintainer_is_not_asked_as_one()
        {
            var accounts = new FakeAccountStore();
            long promoted = await SubmissionRoutingTests.AccountAsync(accounts, "promoted@example.com", admin: true);
            long locked = await SubmissionRoutingTests.AccountAsync(accounts, "locked@example.com", locked: true);
            long unverified = await SubmissionRoutingTests.AccountAsync(accounts, "new@example.com", verified: false);
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, promoted));
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, locked));
            accounts.Maintainers.Add((SubmissionRoutingTests.C64, unverified));

            IReadOnlyList<string> shared = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(touchesShared: true), accounts);

            Assert.Equal(["admin@example.com", "promoted@example.com"], shared.Order(StringComparer.Ordinal));

            IReadOnlyList<string> asMaintainer = await SubmissionRouting.RecipientsForRolesAsync(
                [ApproverRole.Maintainer], SubmissionRoutingTests.C64, accounts);

            Assert.Empty(asMaintainer);
        }

        [Fact]
        public async Task A_locked_or_unverified_administrator_is_not_written_to()
        {
            var accounts = new FakeAccountStore();
            await SubmissionRoutingTests.AccountAsync(accounts, "locked@example.com", admin: true, locked: true);
            await SubmissionRoutingTests.AccountAsync(accounts, "unverified@example.com", admin: true, verified: false);
            await SubmissionRoutingTests.AccountAsync(accounts, "admin@example.com", admin: true);

            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(), accounts);

            Assert.Equal(["admin@example.com"], recipients);
        }

        [Fact]
        public async Task With_nobody_at_all_the_list_is_empty_rather_than_an_error()
        {
            IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForAsync(
                SubmissionRoutingTests.Submission(), new FakeAccountStore());

            Assert.Empty(recipients);
        }
    }
}
