using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewerAssignmentFlows and ReviewerAssignmentRules - the administrator putting
    // people into and out of a system's pool (Phase 6 roles, 2026-09-25).
    //
    // THE PROPERTY THAT MATTERS MOST is the last one: removal takes effect on the very next
    // request with the SAME token, because authority is read from the pool per request and never
    // cached in a session. Phase 6's definition of done asks for exactly that test.
    //
    // No database, no filesystem: the tree listing is handed in as a list, the way the endpoint
    // hands it in after PublishedSystemLister has read the real tree.
    // ###########################################################################################
    public sealed class ReviewerAssignmentFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";

        private static readonly IReadOnlyList<PublishedSystemLister.KnownSystem> Tree =
        [
            new(ReviewerAssignmentFlowsTests.C64, "Commodore", "C64", "250407"),
            new(ReviewerAssignmentFlowsTests.C128, "Commodore", "C128", "310378")
        ];

        private static async Task<(FakeAccountStore Accounts, ReviewAccess Admin, long AnnaId)> SetUpAsync()
        {
            var accounts = new FakeAccountStore();

            long adminId = await accounts.CreateAccountAsync(new NewAccount(
                "admin@example.com", "admin@example.com", "hash", "Admin", ReviewerAssignmentFlowsTests.Now));
            accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };

            long annaId = await accounts.CreateAccountAsync(new NewAccount(
                "anna@example.com", "anna@example.com", "hash", "Anna", ReviewerAssignmentFlowsTests.Now));
            accounts.Accounts[annaId] = accounts.Accounts[annaId] with { IsVerified = true };

            return (accounts, ReviewAccess.For(accounts.Accounts[adminId]), annaId);
        }

        private static ReviewAccess AccessFor(FakeAccountStore accounts, long accountId) =>
            ReviewAccess.For(
                accounts.Accounts[accountId],
                accounts.GetReviewedSystemIdsAsync(accountId).GetAwaiter().GetResult());

        // -----------------------------------------------------------------------------------
        // Adding
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task An_administrator_can_make_a_verified_account_a_reviewer_of_a_shipped_board()
        {
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            var submissions = new FakeSubmissionStore();

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, submissions, ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Contains((ReviewerAssignmentFlowsTests.C64, anna), accounts.Reviewers);

            // And Anna now has authority over exactly that board.
            ReviewAccess access = ReviewerAssignmentFlowsTests.AccessFor(accounts, anna);
            Assert.True(ReviewAuthority.CanPublish(access, ReviewerAssignmentFlowsTests.C64));
            Assert.False(ReviewAuthority.CanPublish(access, ReviewerAssignmentFlowsTests.C128));
        }

        [Fact]
        public async Task Assigning_to_a_board_with_no_systems_row_yet_creates_the_row_as_SHIPPED()
        {
            // The pool table's foreign key needs the row, and a shipped board nobody has submitted
            // to has none. It is 'shipped' because it was in the tree before any submission.
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            var submissions = new FakeSubmissionStore();

            await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, submissions, ReviewerAssignmentFlowsTests.Now);

            NewSubmission row = Assert.Single(submissions.Systems).Value;
            Assert.Equal("Commodore", row.Manufacturer);
            Assert.Equal("C64", row.Hardware);
            Assert.Equal("250407", row.Board);
        }

        [Fact]
        public async Task A_system_that_exists_NEITHER_in_the_database_NOR_in_the_tree_is_refused()
        {
            // A pool row for a system that does not exist would grant authority over something
            // that could only come into being through an unreviewed submission.
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, "Amstrad/CPC/464", anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.False(outcome.IsDone);
            Assert.True(outcome.IsNotFound);
            Assert.Empty(accounts.Reviewers);
        }

        [Fact]
        public async Task A_system_known_only_from_the_DATABASE_can_be_assigned()
        {
            // A contributed system has a row and, until published, no folder in the tree.
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            var submissions = new FakeSubmissionStore();

            await submissions.EnsureSystemAsync(
                "Amstrad/CPC/464", "Amstrad", "CPC", "464", "contributed", ReviewerAssignmentFlowsTests.Now);

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, "Amstrad/CPC/464", anna, [], accounts, submissions, ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
        }

        [Fact]
        public async Task Only_an_ADMINISTRATOR_may_assign_and_a_reviewer_cannot_promote_a_friend()
        {
            // Threat 3: no escalation path. A reviewer's token is useless for changing who reviews.
            (FakeAccountStore accounts, _, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            accounts.Reviewers.Add((ReviewerAssignmentFlowsTests.C64, anna));

            long friend = await accounts.CreateAccountAsync(new NewAccount(
                "friend@example.com", "friend@example.com", "hash", "Friend", ReviewerAssignmentFlowsTests.Now));
            accounts.Accounts[friend] = accounts.Accounts[friend] with { IsVerified = true };

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                ReviewerAssignmentFlowsTests.AccessFor(accounts, anna),
                ReviewerAssignmentFlowsTests.C64, friend, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsForbidden);
            Assert.DoesNotContain((ReviewerAssignmentFlowsTests.C64, friend), accounts.Reviewers);
        }

        [Fact]
        public async Task An_UNVERIFIED_account_cannot_be_made_a_reviewer()
        {
            // ReviewAuthority would refuse it anyway; granting it would make a reviewer who cannot
            // review and an administrator wondering why.
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            accounts.Accounts[anna] = accounts.Accounts[anna] with { IsVerified = false };

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.False(outcome.IsDone);
            Assert.Contains("verified", outcome.Error, StringComparison.OrdinalIgnoreCase);
            Assert.Empty(accounts.Reviewers);
        }

        [Fact]
        public async Task A_LOCKED_account_cannot_be_made_a_reviewer()
        {
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            accounts.Accounts[anna] = accounts.Accounts[anna] with { IsLocked = true };

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.False(outcome.IsDone);
            Assert.Contains("locked", outcome.Error, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task An_ADMINISTRATOR_is_not_put_in_a_pool_because_they_are_in_every_pool_already()
        {
            (FakeAccountStore accounts, ReviewAccess admin, _) = await ReviewerAssignmentFlowsTests.SetUpAsync();

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, admin.Account.Id, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.False(outcome.IsDone);
            Assert.Contains("administrator", outcome.Error, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task An_account_that_does_not_exist_is_NOT_FOUND()
        {
            (FakeAccountStore accounts, ReviewAccess admin, _) = await ReviewerAssignmentFlowsTests.SetUpAsync();

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, 999, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsNotFound);
        }

        [Fact]
        public async Task A_grant_is_written_to_the_audit_trail_naming_who_granted_whom_what()
        {
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();

            await ReviewerAssignmentFlows.AddAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, ReviewerAssignmentFlowsTests.Tree,
                accounts, new FakeSubmissionStore(), ReviewerAssignmentFlowsTests.Now);

            AuditEntry entry = Assert.Single(accounts.Audit);
            Assert.Equal(ReviewerAssignmentFlows.GrantedAction, entry.Action);
            Assert.Equal(admin.Account.Id, entry.ActorAccountId);
            Assert.Equal(ReviewerAssignmentFlowsTests.C64, entry.Subject);
            Assert.Contains("anna@example.com", entry.Detail, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // Removing - and that it bites at once
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Removing_a_reviewer_takes_effect_on_the_very_next_request()
        {
            // *** PHASE 6's DEFINITION OF DONE, verbatim: "removing takes effect immediately,
            // proven by a test using an already-issued token". *** The access built BEFORE the
            // removal still says yes; the one built after, from the same account and the same
            // store - which is what every request does - says no.
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            accounts.Reviewers.Add((ReviewerAssignmentFlowsTests.C64, anna));

            ReviewAccess before = ReviewerAssignmentFlowsTests.AccessFor(accounts, anna);
            Assert.True(ReviewAuthority.CanPublish(before, ReviewerAssignmentFlowsTests.C64));

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.RemoveAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, accounts, ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsDone);

            ReviewAccess after = ReviewerAssignmentFlowsTests.AccessFor(accounts, anna);
            Assert.False(ReviewAuthority.CanPublish(after, ReviewerAssignmentFlowsTests.C64));
            Assert.False(ReviewAuthority.CanReviewAnything(after));
        }

        [Fact]
        public async Task Removing_somebody_who_is_not_in_the_pool_is_not_an_error()
        {
            (FakeAccountStore accounts, ReviewAccess admin, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.RemoveAsync(
                admin, ReviewerAssignmentFlowsTests.C64, anna, accounts, ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsDone);
            Assert.Equal(ReviewerAssignmentFlows.RevokedAction, Assert.Single(accounts.Audit).Action);
        }

        [Fact]
        public async Task Only_an_ADMINISTRATOR_may_remove()
        {
            (FakeAccountStore accounts, _, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            accounts.Reviewers.Add((ReviewerAssignmentFlowsTests.C64, anna));

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.RemoveAsync(
                ReviewerAssignmentFlowsTests.AccessFor(accounts, anna),
                ReviewerAssignmentFlowsTests.C64, anna, accounts, ReviewerAssignmentFlowsTests.Now);

            Assert.True(outcome.IsForbidden);
            Assert.Contains((ReviewerAssignmentFlowsTests.C64, anna), accounts.Reviewers);
        }

        // -----------------------------------------------------------------------------------
        // The overview
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_overview_unions_the_tree_with_the_database_and_names_each_pool()
        {
            (FakeAccountStore accounts, _, long anna) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            var submissions = new FakeSubmissionStore();

            // A contributed system with a row and no folder, and a shipped one with a folder and
            // no row - both must appear, once each.
            await submissions.EnsureSystemAsync(
                "Amstrad/CPC/464", "Amstrad", "CPC", "464", "contributed", ReviewerAssignmentFlowsTests.Now);
            accounts.Reviewers.Add((ReviewerAssignmentFlowsTests.C64, anna));

            IReadOnlyList<SystemWithReviewers> systems = await ReviewerAssignmentFlows.ListSystemsAsync(
                ReviewerAssignmentFlowsTests.Tree, submissions, accounts);

            Assert.Equal(
                ["Amstrad/CPC/464", ReviewerAssignmentFlowsTests.C128, ReviewerAssignmentFlowsTests.C64],
                systems.Select(system => system.SystemId));

            SystemWithReviewers c64 = systems.Single(system => system.SystemId == ReviewerAssignmentFlowsTests.C64);
            Assert.Equal("Anna", Assert.Single(c64.Reviewers).DisplayName);
            Assert.Empty(systems.Single(system => system.SystemId == ReviewerAssignmentFlowsTests.C128).Reviewers);
        }

        [Fact]
        public async Task A_system_in_BOTH_places_is_listed_once_with_the_databases_revision()
        {
            (FakeAccountStore accounts, _, _) = await ReviewerAssignmentFlowsTests.SetUpAsync();
            var submissions = new FakeSubmissionStore();

            await submissions.EnsureSystemAsync(
                ReviewerAssignmentFlowsTests.C64, "Commodore", "C64", "250407", "shipped", ReviewerAssignmentFlowsTests.Now);
            await submissions.SetSystemPublishedAsync(ReviewerAssignmentFlowsTests.C64, "2026-May-14", "hash", ReviewerAssignmentFlowsTests.Now);

            IReadOnlyList<SystemWithReviewers> systems = await ReviewerAssignmentFlows.ListSystemsAsync(
                ReviewerAssignmentFlowsTests.Tree, submissions, accounts);

            SystemWithReviewers c64 = Assert.Single(systems, system => system.SystemId == ReviewerAssignmentFlowsTests.C64);
            Assert.Equal("2026-May-14", c64.CurrentRevision);
        }

        // -----------------------------------------------------------------------------------
        // The rules' sentences
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_verified_unlocked_ordinary_account_is_grantable()
        {
            var account = new AccountRecord(
                1, "a@example.com", "a@example.com", "hash", "A",
                IsVerified: true, IsAdministrator: false, IsLocked: false,
                ReviewerAssignmentFlowsTests.Now, null);

            Assert.Null(ReviewerAssignmentRules.WhyNotGrantable(account));
        }
    }
}
