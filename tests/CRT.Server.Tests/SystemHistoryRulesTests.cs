using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SystemHistoryRules - what has happened to a system, newest first, on the Systems screen
    // (owner request, 2026-09-27: "I would like to see the date, newest first, to understand what
    // has happened to a system").
    // ###########################################################################################
    public sealed class SystemHistoryRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";

        private static SubmissionRecord Record(long id, string state, DateTimeOffset created, DateTimeOffset? decided = null, string? comment = null, long? account = null) =>
            new(id, C64, account, $"sender{id}@example.com", "hash", "", state, $"Submission {id}", 1, created, null, decided, comment);

        private static AccountRecord Account(long id, string name) =>
            new(id, $"{name.ToLowerInvariant()}@example.com", $"{name.ToLowerInvariant()}@example.com", "hash", name, true, false, false, Now, null);

        // ###########################################################################################
        // The two sources in one list, NEWEST FIRST: a submission's sending and its decision (by
        // whom, and what they told the contributor), and the audit rows about the system.
        // ###########################################################################################
        [Fact]
        public void Submissions_and_system_events_make_one_list_newest_first()
        {
            var named = new Dictionary<long, AccountRecord> { [7] = SystemHistoryRulesTests.Account(7, "Anna") };

            SystemSubmissionRecord merged = new(
                SystemHistoryRulesTests.Record(9, SubmissionState.Merged, Now.AddDays(-3), decided: Now.AddDays(-2)),
                DecidedByMaintainer: true,
                DecidedByAccountId: 7);

            SystemSubmissionRecord rejected = new(
                SystemHistoryRulesTests.Record(10, SubmissionState.Rejected, Now.AddDays(-1), decided: Now.AddHours(-2), comment: "Wrong board."),
                DecidedByMaintainer: true,
                DecidedByAccountId: 7);

            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", SystemHistoryEvents.PublishedToProduction, C64, "revision 2026-September-26; 3 file(s) copied", Now.AddDays(-1).AddHours(-1)),
                new(1, "admin@example.com", SystemHistoryEvents.MaintainerAdded, C64, "account 7 (anna@example.com)", Now.AddDays(-4)),
            ];

            IReadOnlyList<SystemHistoryEntry> history = SystemHistoryRules.Build([rejected, merged], named, audit);

            Assert.Equal(
                [
                    (SystemHistoryEvents.Decided, (long?)10),
                    (SystemHistoryEvents.Sent, 10),
                    (SystemHistoryEvents.PublishedToProduction, null),
                    (SystemHistoryEvents.Decided, 9),
                    (SystemHistoryEvents.Sent, 9),
                    (SystemHistoryEvents.MaintainerAdded, null),
                ],
                history.Select(entry => (entry.Event, entry.SubmissionId)));

            Assert.True(history.Zip(history.Skip(1)).All(pair => pair.First.AtUtc >= pair.Second.AtUtc));

            SystemHistoryEntry decision = history[0];
            Assert.Equal("Anna", decision.Who);
            Assert.Equal(SubmissionState.Rejected, decision.Detail);
            Assert.Equal("Wrong board.", decision.Note);

            SystemHistoryEntry sending = history[1];
            Assert.Equal("sender10@example.com", sending.Who);
            Assert.Equal("Submission 10", sending.Detail);

            Assert.Equal("anna@example.com", history[5].Detail);
        }

        // Only a decision that ENDED somewhere is a line: a push-back leaves its submissions pending
        // (the push-back is its own line) and a first of two approvals publishes nothing.
        [Theory]
        [InlineData(SubmissionState.Pending)]
        [InlineData(SubmissionState.Approved)]
        public void A_submission_left_waiting_has_no_decision_line(string state)
        {
            SystemSubmissionRecord waiting = new(
                SystemHistoryRulesTests.Record(9, state, Now.AddDays(-1), decided: Now),
                DecidedByMaintainer: false);

            Assert.Equal([SystemHistoryEvents.Sent], SystemHistoryRules.Build([waiting], new Dictionary<long, AccountRecord>(), []).Select(entry => entry.Event));
        }

        // Sign-ins, unused-file removals and anything else not about the system stay out.
        [Fact]
        public void Audit_rows_that_are_not_about_the_system_are_left_out()
        {
            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", "account.login", C64, null, Now),
                new(1, "admin@example.com", SystemHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now),
            ];

            SystemHistoryEntry only = Assert.Single(SystemHistoryRules.Build([], new Dictionary<long, AccountRecord>(), audit));
            Assert.Equal(SystemHistoryEvents.Placed, only.Event);
            Assert.Equal("Commodore 64 / 250407", only.Detail);
            Assert.Equal("admin@example.com", only.Who);
        }

        [Theory]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "account 7 (anna@example.com)", "anna@example.com")]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "account 7", "account 7")]
        [InlineData(SystemHistoryEvents.Amended, "Commodore/C64/250407: amendment 2, 3 file(s)", "amendment 2, 3 file(s)")]
        [InlineData(SystemHistoryEvents.PushedBack, "RestoreProduction; 3 file(s) restored, 1 removed; 1 submission(s) returned to the queue; Not ready.", "3 file(s) restored, 1 removed; 1 submission(s) returned to the queue; Not ready.")]
        [InlineData(SystemHistoryEvents.RejectedFromBeta, "RestoreProduction; 3 file(s) restored, 1 removed; 1 submission(s) rejected; Not for this board.", "3 file(s) restored, 1 removed; 1 submission(s) rejected; Not for this board.")]
        [InlineData(SystemHistoryEvents.Invited, "new@example.com", "new@example.com")]
        public void An_audit_rows_detail_is_the_part_a_person_reads(string action, string detail, string expected)
        {
            Assert.Equal(expected, SystemHistoryRules.DetailOf(new AuditEntry(1, "admin@example.com", action, C64, detail, Now)));
        }

        // Beta > Prod's "Reject" (2026-09-28) is in the history like a push-back - an audit action
        // left out of ShownActions would vanish from the Systems screen without a word.
        [Fact]
        public void A_rejection_from_BETA_is_in_the_history()
        {
            AuditEntry[] audit =
            [
                new(1, "admin@example.com", SystemHistoryEvents.RejectedFromBeta, C64, "RestoreProduction; 1 file(s) restored, 0 removed; 1 submission(s) rejected; No.", Now),
            ];

            SystemHistoryEntry only = Assert.Single(SystemHistoryRules.Build([], new Dictionary<long, AccountRecord>(), audit));
            Assert.Equal(SystemHistoryEvents.RejectedFromBeta, only.Event);
            Assert.Equal("1 file(s) restored, 0 removed; 1 submission(s) rejected; No.", only.Detail);
        }

        // An amendment is recorded under its submission ("#41") - the line names that submission.
        [Fact]
        public void An_amendment_names_its_submission()
        {
            SystemHistoryEntry amended = Assert.Single(SystemHistoryRules.Build(
                [],
                new Dictionary<long, AccountRecord>(),
                [new AuditEntry(1, "anna@example.com", SystemHistoryEvents.Amended, "#41", "Commodore/C64/250407: amendment 1, 2 file(s)", Now)]));

            Assert.Equal(41, amended.SubmissionId);
        }

        // -----------------------------------------------------------------------------------
        // Without addresses - an account that does not maintain the system (owner request,
        // 2026-10-05: maintainers see email addresses only for their own systems)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // A pool change is written "account 7 (anna@example.com)"; the account number is how its
        // person is named without the address. Only the three pool actions carry one, and only in
        // that shape - anything else names nobody rather than a wrong account.
        // ###########################################################################################
        [Theory]
        [InlineData(SystemHistoryEvents.MaintainerAdded, "account 7 (anna@example.com)", 7L)]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "account 7 (anna@example.com)", 7L)]
        [InlineData(SystemHistoryEvents.InvitationAccepted, "account 12 (bo@example.com)", 12L)]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "  account 7  ", 7L)]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "anna@example.com", null)]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "account x (anna@example.com)", null)]
        [InlineData(SystemHistoryEvents.MaintainerRemoved, "", null)]
        [InlineData(SystemHistoryEvents.Invited, "account 7 (anna@example.com)", null)]
        [InlineData(SystemHistoryEvents.Placed, "account 7", null)]
        public void A_pool_change_names_its_account_by_number(string action, string detail, long? expected)
        {
            Assert.Equal(expected, SystemHistoryRules.AccountNamedBy(new AuditEntry(1, "admin@example.com", action, C64, detail, Now)));
        }

        // ###########################################################################################
        // The same history, with no "@" anywhere in it: a pool change names its person by their
        // account (or nobody, for an account no longer there), an invitation - whose detail IS an
        // address - says nothing, and a default detail that happens to carry an address is dropped.
        // ###########################################################################################
        [Fact]
        public void Without_addresses_a_pool_change_names_its_person_by_account_and_an_invitation_says_nothing()
        {
            var named = new Dictionary<long, AccountRecord>
            {
                [1] = SystemHistoryRulesTests.Account(1, "Dennis"),
                [7] = SystemHistoryRulesTests.Account(7, "Anna")
            };

            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", SystemHistoryEvents.MaintainerRemoved, C64, "account 7 (anna@example.com)", Now.AddMinutes(-5)),
                new(1, "admin@example.com", SystemHistoryEvents.MaintainerAdded, C64, "account 99 (gone@example.com)", Now.AddMinutes(-4)),
                new(1, "admin@example.com", SystemHistoryEvents.Invited, C64, "new@example.com", Now.AddMinutes(-3)),
                new(1, "admin@example.com", SystemHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now.AddMinutes(-2)),
                new(1, "admin@example.com", SystemHistoryEvents.Placed, C64, "named after someone@example.com", Now.AddMinutes(-1)),
            ];

            IReadOnlyList<SystemHistoryEntry> history = SystemHistoryRules.Build([], named, audit, showAddresses: false);

            Assert.Equal(
                [
                    (SystemHistoryEvents.Placed, (string?)null),
                    (SystemHistoryEvents.Placed, "Commodore 64 / 250407"),
                    (SystemHistoryEvents.Invited, null),
                    (SystemHistoryEvents.MaintainerAdded, null),
                    (SystemHistoryEvents.MaintainerRemoved, "Anna"),
                ],
                history.Select(entry => (entry.Event, entry.Detail)));

            Assert.All(history, entry => Assert.Equal("Dennis", entry.Who));
            Assert.DoesNotContain(history, entry => (entry.Who + entry.Detail + entry.Note).Contains('@'));
        }

        // ###########################################################################################
        // Who did it, without an address: the name on their account; else the label when it is no
        // address ("the contributor"); else nobody. A sending with no account behind it names nobody.
        // ###########################################################################################
        [Fact]
        public void Without_addresses_an_actor_is_named_by_account_or_by_a_label_that_is_no_address()
        {
            var named = new Dictionary<long, AccountRecord> { [7] = SystemHistoryRulesTests.Account(7, "Anna") };

            IReadOnlyList<AuditEntry> audit =
            [
                new(7, "anna@example.com", SystemHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now.AddMinutes(-3)),
                new(null, "the contributor", SystemHistoryEvents.DraftDiscarded, "#9", null, Now.AddMinutes(-2)),
                new(null, "someone@example.com", SystemHistoryEvents.DraftDiscarded, "#9", null, Now.AddMinutes(-1)),
            ];

            SystemSubmissionRecord sent = new(
                SystemHistoryRulesTests.Record(9, SubmissionState.Pending, Now.AddDays(-1)),
                DecidedByMaintainer: false);

            IReadOnlyList<SystemHistoryEntry> history = SystemHistoryRules.Build([sent], named, audit, showAddresses: false);

            Assert.Equal(
                [(string?)null, "the contributor", "Anna", null],
                history.Select(entry => entry.Who));

            // With addresses, the same rows say them.
            IReadOnlyList<SystemHistoryEntry> shown = SystemHistoryRules.Build([sent], named, audit);

            Assert.Equal(
                [(string?)"someone@example.com", "the contributor", "anna@example.com", "sender9@example.com"],
                shown.Select(entry => entry.Who));
        }

        // -----------------------------------------------------------------------------------
        // Through the real flows
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The detail carries the history built from the stores, and the events that used to be
        // recorded under no system - a removal without the address, an accepted invitation, a saved
        // placement - are now recorded so it can name them.
        // ###########################################################################################
        [Fact]
        public async Task A_systems_detail_tells_its_history_through_the_real_flows()
        {
            var accounts = new FakeAccountStore();
            var submissions = new FakeSubmissionStore();

            long adminId = await accounts.CreateAccountAsync(new NewAccount("admin@example.com", "admin@example.com", "hash", "Dennis", Now));
            accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };
            long annaId = await accounts.CreateAccountAsync(new NewAccount("anna@example.com", "anna@example.com", "hash", "Anna", Now));
            accounts.Accounts[annaId] = accounts.Accounts[annaId] with { IsVerified = true };

            ReviewAccess admin = ReviewAccess.For(accounts.Accounts[adminId]);
            IReadOnlyList<PublishedSystemLister.KnownSystem> tree = [new(C64, "Commodore", "C64", "250407")];

            await MaintainerAssignmentFlows.AddAsync(admin, C64, annaId, tree, accounts, submissions, Now.AddHours(-3));
            await MaintainerAssignmentFlows.RemoveAsync(admin, C64, annaId, accounts, Now.AddHours(-2));

            submissions.Submissions[5] = SystemHistoryRulesTests.Record(5, SubmissionState.Pending, Now.AddHours(-1));

            SystemDetailAnswer detail = (await SystemOverviewFlow.DetailAsync(
                admin, C64, tree, null, submissions, accounts, now: Now)).Detail!;

            Assert.Equal(
                [
                    (SystemHistoryEvents.Sent, "Submission 5"),
                    (SystemHistoryEvents.MaintainerRemoved, "anna@example.com"),
                    (SystemHistoryEvents.MaintainerAdded, "anna@example.com"),
                ],
                detail.History!.Select(entry => (entry.Event, entry.Detail)));
        }

        [Fact]
        public async Task A_saved_placement_is_recorded_under_its_system()
        {
            var accounts = new FakeAccountStore();
            var submissions = new FakeSubmissionStore();
            string beta = Path.Combine(Path.GetTempPath(), "crt-history", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(beta);

            try
            {
                DataTreeBuilder.ListingMaster(beta, ApprovePublishFlowTests.OtherListedBoard);

                long adminId = await accounts.CreateAccountAsync(new NewAccount("admin@example.com", "admin@example.com", "hash", "Dennis", Now));
                accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };

                await submissions.EnsureSystemAsync(C64, "Commodore", "C64", "250407", "contributed", Now);

                var flow = new SystemListingFlow(submissions, new PublishLock(), NullLogger<SystemListingFlow>.Instance, accounts);

                SetPlacementOutcome outcome = await flow.SetAsync(
                    ReviewAccess.For(accounts.Accounts[adminId]),
                    new SetPlacementRequest(C64, "Commodore 64", "250407", string.Empty, null),
                    beta,
                    Now);

                Assert.NotNull(outcome.Answer);

                AuditEntry placed = Assert.Single(accounts.Audit, entry => entry.Action == SystemHistoryEvents.Placed);
                Assert.Equal(C64, placed.Subject);
                Assert.Equal("Commodore 64 / 250407", placed.Detail);
            }
            finally
            {
                Directory.Delete(Path.GetDirectoryName(beta)!, recursive: true);
            }
        }
    }
}
