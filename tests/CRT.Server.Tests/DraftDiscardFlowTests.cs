using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT (owner request, 2026-09-28: "if the contributor has
    // discarded his own data ... it should be clearly visible on the system and for the
    // maintainer(s) - both in the BETA to PROD queue, but also in the normal queue").
    //
    // Recorded once per submission (the FIRST time is kept), audited once under "#{id}" so the
    // board's history shows it, and carried to every screen a maintainer decides from: the queue,
    // the detail, "Beta > Prod" (its list and its plan) and the Boards screen.
    // ###########################################################################################
    public sealed class DraftDiscardFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 28, 12, 0, 0, TimeSpan.Zero);

        private const string C128 = "Commodore/C128/310378";

        private static SubmissionRecord Record(long id, string state = SubmissionState.Merged, string? email = "dennis@example.com") =>
            new(id, C128, null, email, "hash", "", state, "Changed \"a\" in short-line", 1, Now.AddDays(-1), null, Now.AddHours(-2));

        [Fact]
        public async Task A_discard_is_recorded_and_audited_under_the_submission()
        {
            var store = new FakeSubmissionStore();
            var accounts = new FakeAccountStore();

            bool recorded = await DraftDiscardFlow.RecordAsync(Record(42), store, accounts, Now);

            Assert.True(recorded);
            Assert.Equal(Now, store.DraftDiscards[42]);

            AuditEntry audit = Assert.Single(accounts.Audit);
            Assert.Equal(BoardHistoryEvents.DraftDiscarded, audit.Action);
            Assert.Equal("#42", audit.Subject);
            Assert.Equal("dennis@example.com", audit.ActorLabel);
            Assert.Equal(Now, audit.AtUtc);
        }

        // ###########################################################################################
        // *** A REPEAT KEEPS THE FIRST TIME AND WRITES NO SECOND HISTORY LINE. *** CRT tries again
        // until an answer arrives, so the same notice can land twice after a lost answer; the
        // history must still say it once, at the time it first happened.
        // ###########################################################################################
        [Fact]
        public async Task A_repeated_notice_keeps_the_first_time_and_writes_no_second_history_line()
        {
            var store = new FakeSubmissionStore();
            var accounts = new FakeAccountStore();

            await DraftDiscardFlow.RecordAsync(Record(42), store, accounts, Now);
            bool again = await DraftDiscardFlow.RecordAsync(Record(42), store, accounts, Now.AddHours(3));

            Assert.False(again);
            Assert.Equal(Now, store.DraftDiscards[42]);
            Assert.Single(accounts.Audit);
        }

        // Contributing needs no account, and a submission may carry no address at all.
        [Fact]
        public async Task A_submission_with_no_address_is_audited_as_the_contributor()
        {
            var accounts = new FakeAccountStore();

            await DraftDiscardFlow.RecordAsync(Record(7, email: null), new FakeSubmissionStore(), accounts, Now);

            Assert.Equal("the contributor", Assert.Single(accounts.Audit).ActorLabel);
        }

        // ###########################################################################################
        // THE QUEUE. The row itself says it, so the maintainer sees it before opening the submission.
        // ###########################################################################################
        [Fact]
        public async Task The_queue_row_carries_when_the_draft_was_discarded_and_only_for_that_row()
        {
            var store = new FakeSubmissionStore();
            AccountRecord administrator = new(6, "a@example.com", "a@example.com", "hash", "A", true, true, false, Now, null);

            long discarded = await AddPendingAsync(store);
            long kept = await AddPendingAsync(store);

            await DraftDiscardFlow.RecordAsync(store.Submissions[discarded], store, new FakeAccountStore(), Now);

            ReviewQueueAnswer answer = await ReviewQueueFlow.BuildAsync(
                ReviewAccess.For(administrator),
                await store.GetQueueAsync(100, CancellationToken.None),
                dataTreeRoot: null,
                store,
                new FakeAccountStore(),
                CancellationToken.None);

            Assert.Equal(Now, answer.Submissions.Single(entry => entry.Id == discarded).DraftDiscardedUtc);
            Assert.Null(answer.Submissions.Single(entry => entry.Id == kept).DraftDiscardedUtc);
        }

        // "Beta > Prod": the plan names the discarding contributor beside their work.
        [Fact]
        public void A_carried_submission_says_when_its_contributor_discarded_their_draft()
        {
            IReadOnlyList<CarriedSubmission> carrying = ProductionPromotionRules.Carrying(
                [Record(1), Record(2)],
                new Dictionary<long, DateTimeOffset> { [2] = Now });

            Assert.Equal(Now, carrying.Single(carried => carried.Id == 2).DraftDiscardedUtc);
            Assert.Null(carrying.Single(carried => carried.Id == 1).DraftDiscardedUtc);
        }

        [Fact]
        public void The_boards_screen_lists_the_discard_beside_the_submission()
        {
            IReadOnlyList<BoardSubmissionEntry> listed = BoardOverviewFlow.Submissions(
                [new BoardSubmissionRecord(Record(3), DecidedByMaintainer: true), new BoardSubmissionRecord(Record(4), DecidedByMaintainer: true)],
                boardProductionPublishedUtc: null,
                accounts: null,
                discarded: new Dictionary<long, DateTimeOffset> { [4] = Now });

            Assert.Equal(Now, listed.Single(entry => entry.Id == 4).DraftDiscardedUtc);
            Assert.Null(listed.Single(entry => entry.Id == 3).DraftDiscardedUtc);
        }

        // The audit row the flow writes is one the board's history shows, with who did it.
        [Fact]
        public async Task The_boards_history_shows_the_discard()
        {
            var accounts = new FakeAccountStore();
            SubmissionRecord record = Record(42);

            await DraftDiscardFlow.RecordAsync(record, new FakeSubmissionStore(), accounts, Now);

            IReadOnlyList<BoardHistoryEntry> history = BoardHistoryRules.Build(
                [new BoardSubmissionRecord(record, DecidedByMaintainer: true)],
                new Dictionary<long, AccountRecord>(),
                accounts.Audit);

            BoardHistoryEntry line = history.Single(entry => entry.Event == BoardHistoryEvents.DraftDiscarded);
            Assert.Equal(42, line.SubmissionId);
            Assert.Equal("dennis@example.com", line.Who);
        }

        private static async Task<long> AddPendingAsync(FakeSubmissionStore store)
        {
            long id = await store.CreateAsync(
                new NewSubmission(
                    C128, "Commodore", "C128", "310378",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "A change.",
                    1, [], Now, Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, SubmissionState.Pending, Now, CancellationToken.None);
            return id;
        }
    }
}
