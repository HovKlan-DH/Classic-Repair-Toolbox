using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers BoardHistoryRules - what has happened to a board, newest first, on the Boards screen
    // (owner request, 2026-09-27: "I would like to see the date, newest first, to understand what
    // has happened to a system").
    // ###########################################################################################
    public sealed class BoardHistoryRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";

        private static SubmissionRecord Record(long id, string state, DateTimeOffset created, DateTimeOffset? decided = null, string? comment = null, long? account = null) =>
            new(id, C64, account, $"sender{id}@example.com", "hash", "", state, $"Submission {id}", 1, created, null, decided, comment);

        private static AccountRecord Account(long id, string name) =>
            new(id, $"{name.ToLowerInvariant()}@example.com", $"{name.ToLowerInvariant()}@example.com", "hash", name, true, false, false, Now, null);

        // ###########################################################################################
        // The two sources in one list, NEWEST FIRST: a submission's sending and its decision (by
        // whom, and what they told the contributor), and the audit rows about the board.
        // ###########################################################################################
        [Fact]
        public void Submissions_and_board_events_make_one_list_newest_first()
        {
            var named = new Dictionary<long, AccountRecord> { [7] = BoardHistoryRulesTests.Account(7, "Anna") };

            BoardSubmissionRecord merged = new(
                BoardHistoryRulesTests.Record(9, SubmissionState.Merged, Now.AddDays(-3), decided: Now.AddDays(-2)),
                DecidedByMaintainer: true,
                DecidedByAccountId: 7);

            BoardSubmissionRecord rejected = new(
                BoardHistoryRulesTests.Record(10, SubmissionState.Rejected, Now.AddDays(-1), decided: Now.AddHours(-2), comment: "Wrong board."),
                DecidedByMaintainer: true,
                DecidedByAccountId: 7);

            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", BoardHistoryEvents.PublishedToProduction, C64, "revision 2026-September-26; 3 file(s) copied", Now.AddDays(-1).AddHours(-1)),
                new(1, "admin@example.com", BoardHistoryEvents.MaintainerAdded, C64, "account 7 (anna@example.com)", Now.AddDays(-4)),
            ];

            IReadOnlyList<BoardHistoryEntry> history = BoardHistoryRules.Build([rejected, merged], named, audit);

            Assert.Equal(
                [
                    (BoardHistoryEvents.Decided, (long?)10),
                    (BoardHistoryEvents.Sent, 10),
                    (BoardHistoryEvents.PublishedToProduction, null),
                    (BoardHistoryEvents.Decided, 9),
                    (BoardHistoryEvents.Sent, 9),
                    (BoardHistoryEvents.MaintainerAdded, null),
                ],
                history.Select(entry => (entry.Event, entry.SubmissionId)));

            Assert.True(history.Zip(history.Skip(1)).All(pair => pair.First.AtUtc >= pair.Second.AtUtc));

            BoardHistoryEntry decision = history[0];
            Assert.Equal("Anna", decision.Who);
            Assert.Equal(SubmissionState.Rejected, decision.Detail);
            Assert.Equal("Wrong board.", decision.Note);

            BoardHistoryEntry sending = history[1];
            Assert.Equal("sender10@example.com", sending.Who);
            Assert.Equal("Submission 10", sending.Detail);

            // The person a pool change names, by their account's name (owner request, 2026-10-09:
            // "Please use name (if available, otherwise mail address) for all maintainers").
            Assert.Equal("Anna", history[5].Detail);
        }

        // Only a decision that ENDED somewhere is a line: a push-back leaves its submissions pending
        // (the push-back is its own line) and a first of two approvals publishes nothing.
        [Theory]
        [InlineData(SubmissionState.Pending)]
        [InlineData(SubmissionState.Approved)]
        public void A_submission_left_waiting_has_no_decision_line(string state)
        {
            BoardSubmissionRecord waiting = new(
                BoardHistoryRulesTests.Record(9, state, Now.AddDays(-1), decided: Now),
                DecidedByMaintainer: false);

            Assert.Equal([BoardHistoryEvents.Sent], BoardHistoryRules.Build([waiting], new Dictionary<long, AccountRecord>(), []).Select(entry => entry.Event));
        }

        // Sign-ins, unused-file removals and anything else not about the board stay out.
        [Fact]
        public void Audit_rows_that_are_not_about_the_board_are_left_out()
        {
            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", "account.login", C64, null, Now),
                new(1, "admin@example.com", BoardHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now),
            ];

            BoardHistoryEntry only = Assert.Single(BoardHistoryRules.Build([], new Dictionary<long, AccountRecord>(), audit));
            Assert.Equal(BoardHistoryEvents.Placed, only.Event);
            Assert.Equal("Commodore 64 / 250407", only.Detail);
            Assert.Equal("admin@example.com", only.Who);
        }

        [Theory]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "account 7 (anna@example.com)", "anna@example.com")]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "account 7", "account 7")]
        [InlineData(BoardHistoryEvents.Amended, "Commodore/C64/250407: amendment 2, 3 file(s)", "amendment 2, 3 file(s)")]
        [InlineData(BoardHistoryEvents.PushedBack, "RestoreProduction; 3 file(s) restored, 1 removed; 1 submission(s) returned to the queue; Not ready.", "3 file(s) restored, 1 removed; 1 submission(s) returned to the queue; Not ready.")]
        [InlineData(BoardHistoryEvents.RejectedFromBeta, "RestoreProduction; 3 file(s) restored, 1 removed; 1 submission(s) rejected; Not for this board.", "3 file(s) restored, 1 removed; 1 submission(s) rejected; Not for this board.")]
        [InlineData(BoardHistoryEvents.Invited, "new@example.com", "new@example.com")]
        public void An_audit_rows_detail_is_the_part_a_person_reads(string action, string detail, string expected)
        {
            Assert.Equal(expected, BoardHistoryRules.DetailOf(new AuditEntry(1, "admin@example.com", action, C64, detail, Now)));
        }

        // Beta > Prod's "Reject" (2026-09-28) is in the history like a push-back - an audit action
        // left out of ShownActions would vanish from the Boards screen without a word.
        [Fact]
        public void A_rejection_from_BETA_is_in_the_history()
        {
            AuditEntry[] audit =
            [
                new(1, "admin@example.com", BoardHistoryEvents.RejectedFromBeta, C64, "RestoreProduction; 1 file(s) restored, 0 removed; 1 submission(s) rejected; No.", Now),
            ];

            BoardHistoryEntry only = Assert.Single(BoardHistoryRules.Build([], new Dictionary<long, AccountRecord>(), audit));
            Assert.Equal(BoardHistoryEvents.RejectedFromBeta, only.Event);
            Assert.Equal("1 file(s) restored, 0 removed; 1 submission(s) rejected; No.", only.Detail);
        }

        // An amendment is recorded under its submission ("#41") - the line names that submission.
        [Fact]
        public void An_amendment_names_its_submission()
        {
            BoardHistoryEntry amended = Assert.Single(BoardHistoryRules.Build(
                [],
                new Dictionary<long, AccountRecord>(),
                [new AuditEntry(1, "anna@example.com", BoardHistoryEvents.Amended, "#41", "Commodore/C64/250407: amendment 1, 2 file(s)", Now)]));

            Assert.Equal(41, amended.SubmissionId);
        }

        // -----------------------------------------------------------------------------------
        // Without addresses - an account that does not maintain the board (owner request,
        // 2026-10-05: maintainers see email addresses only for their own boards)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // A pool change is written "account 7 (anna@example.com)"; the account number is how its
        // person is named without the address. Only the three pool actions carry one, and only in
        // that shape - anything else names nobody rather than a wrong account.
        // ###########################################################################################
        [Theory]
        [InlineData(BoardHistoryEvents.MaintainerAdded, "account 7 (anna@example.com)", 7L)]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "account 7 (anna@example.com)", 7L)]
        [InlineData(BoardHistoryEvents.InvitationAccepted, "account 12 (bo@example.com)", 12L)]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "  account 7  ", 7L)]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "anna@example.com", null)]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "account x (anna@example.com)", null)]
        [InlineData(BoardHistoryEvents.MaintainerRemoved, "", null)]
        [InlineData(BoardHistoryEvents.Invited, "account 7 (anna@example.com)", null)]
        [InlineData(BoardHistoryEvents.Placed, "account 7", null)]
        public void A_pool_change_names_its_account_by_number(string action, string detail, long? expected)
        {
            Assert.Equal(expected, BoardHistoryRules.AccountNamedBy(new AuditEntry(1, "admin@example.com", action, C64, detail, Now)));
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
                [1] = BoardHistoryRulesTests.Account(1, "Dennis"),
                [7] = BoardHistoryRulesTests.Account(7, "Anna")
            };

            IReadOnlyList<AuditEntry> audit =
            [
                new(1, "admin@example.com", BoardHistoryEvents.MaintainerRemoved, C64, "account 7 (anna@example.com)", Now.AddMinutes(-5)),
                new(1, "admin@example.com", BoardHistoryEvents.MaintainerAdded, C64, "account 99 (gone@example.com)", Now.AddMinutes(-4)),
                new(1, "admin@example.com", BoardHistoryEvents.Invited, C64, "new@example.com", Now.AddMinutes(-3)),
                new(1, "admin@example.com", BoardHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now.AddMinutes(-2)),
                new(1, "admin@example.com", BoardHistoryEvents.Placed, C64, "named after someone@example.com", Now.AddMinutes(-1)),
            ];

            IReadOnlyList<BoardHistoryEntry> history = BoardHistoryRules.Build([], named, audit, showAddresses: false);

            Assert.Equal(
                [
                    (BoardHistoryEvents.Placed, (string?)null),
                    (BoardHistoryEvents.Placed, "Commodore 64 / 250407"),
                    (BoardHistoryEvents.Invited, null),
                    (BoardHistoryEvents.MaintainerAdded, null),
                    (BoardHistoryEvents.MaintainerRemoved, "Anna"),
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
            var named = new Dictionary<long, AccountRecord> { [7] = BoardHistoryRulesTests.Account(7, "Anna") };

            IReadOnlyList<AuditEntry> audit =
            [
                new(7, "anna@example.com", BoardHistoryEvents.Placed, C64, "Commodore 64 / 250407", Now.AddMinutes(-3)),
                new(null, "the contributor", BoardHistoryEvents.DraftDiscarded, "#9", null, Now.AddMinutes(-2)),
                new(null, "someone@example.com", BoardHistoryEvents.DraftDiscarded, "#9", null, Now.AddMinutes(-1)),
            ];

            BoardSubmissionRecord sent = new(
                BoardHistoryRulesTests.Record(9, SubmissionState.Pending, Now.AddDays(-1)),
                DecidedByMaintainer: false);

            IReadOnlyList<BoardHistoryEntry> history = BoardHistoryRules.Build([sent], named, audit, showAddresses: false);

            Assert.Equal(
                [(string?)null, "the contributor", "Anna", null],
                history.Select(entry => entry.Who));

            // With addresses, the same rows say them - but a person with a name on their account
            // is named by it (owner request, 2026-10-09), so Anna stays Anna.
            IReadOnlyList<BoardHistoryEntry> shown = BoardHistoryRules.Build([sent], named, audit);

            Assert.Equal(
                [(string?)"someone@example.com", "the contributor", "Anna", "sender9@example.com"],
                shown.Select(entry => entry.Who));
        }

        // -----------------------------------------------------------------------------------
        // Through the real flows
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The detail carries the history built from the stores, and the events that used to be
        // recorded under no board - a removal without the address, an accepted invitation, a saved
        // placement - are now recorded so it can name them.
        // ###########################################################################################
        [Fact]
        public async Task A_boards_detail_tells_its_history_through_the_real_flows()
        {
            var accounts = new FakeAccountStore();
            var submissions = new FakeSubmissionStore();

            long adminId = await accounts.CreateAccountAsync(new NewAccount("admin@example.com", "admin@example.com", "hash", "Dennis", Now));
            accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };
            long annaId = await accounts.CreateAccountAsync(new NewAccount("anna@example.com", "anna@example.com", "hash", "Anna", Now));
            accounts.Accounts[annaId] = accounts.Accounts[annaId] with { IsVerified = true };

            ReviewAccess admin = ReviewAccess.For(accounts.Accounts[adminId]);
            IReadOnlyList<PublishedBoardLister.KnownBoard> tree = [new(C64, "Commodore", "C64", "250407")];

            await MaintainerAssignmentFlows.AddAsync(admin, C64, annaId, tree, accounts, submissions, Now.AddHours(-3));
            await MaintainerAssignmentFlows.RemoveAsync(admin, C64, annaId, accounts, Now.AddHours(-2));

            submissions.Submissions[5] = BoardHistoryRulesTests.Record(5, SubmissionState.Pending, Now.AddHours(-1));

            BoardDetailAnswer detail = (await BoardOverviewFlow.DetailAsync(
                admin, C64, tree, null, submissions, accounts, now: Now)).Detail!;

            Assert.Equal(
                [
                    (BoardHistoryEvents.Sent, "Submission 5"),
                    (BoardHistoryEvents.MaintainerRemoved, "Anna"),
                    (BoardHistoryEvents.MaintainerAdded, "Anna"),
                ],
                detail.History!.Select(entry => (entry.Event, entry.Detail)));
        }

        [Fact]
        public async Task A_saved_placement_is_recorded_under_its_board()
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

                await submissions.EnsureBoardAsync(C64, "Commodore", "C64", "250407", "contributed", Now);

                var flow = new BoardListingFlow(submissions, new PublishLock(), NullLogger<BoardListingFlow>.Instance, accounts);

                SetPlacementOutcome outcome = await flow.SetAsync(
                    ReviewAccess.For(accounts.Accounts[adminId]),
                    new SetPlacementRequest(C64, "Commodore 64", "250407", string.Empty, null),
                    beta,
                    Now);

                Assert.NotNull(outcome.Answer);

                AuditEntry placed = Assert.Single(accounts.Audit, entry => entry.Action == BoardHistoryEvents.Placed);
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
