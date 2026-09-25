using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ISubmissionStore.SetDecisionAsync - that a review decision is actually RECORDED
    // (Phase 5, task 5).
    //
    // *** THE COLUMNS EXISTED FROM THE START AND NOTHING EVER WROTE THEM. *** `submissions` has
    // carried `decided_by` and `decision_comment` since 0001_initial.sql, and until this task the
    // only writer was SetStateAsync, which sets neither. So the audit trail could say a submission
    // was rejected and could not say by whom or why.
    //
    // The comment half matters beyond the audit trail: contributing needs no account, so the
    // contact email and this comment are the ENTIRE feedback channel to the person who did the
    // work. A test asserting only that the state moved would pass happily while the comment was
    // being dropped - which is why the fake records what was written and these tests read it back.
    //
    // Run against FakeSubmissionStore, which is deliberately STRICTER than convenient: it throws
    // for a submission that does not exist, because the real statement is an UPDATE that would
    // otherwise affect zero rows and report success.
    // ###########################################################################################
    public sealed class ReviewDecisionRecordingTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

        private const long MaintainerAccountId = 7;

        private static async Task<(FakeSubmissionStore Store, long Id)> PendingAsync()
        {
            var store = new FakeSubmissionStore();

            long id = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "Corrected R12.",
                    1, [], ReviewDecisionRecordingTests.Now, ReviewDecisionRecordingTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(
                id, SubmissionState.Pending, ReviewDecisionRecordingTests.Now, CancellationToken.None);

            return (store, id);
        }

        [Fact]
        public async Task A_REJECTION_records_the_state_the_maintainer_and_the_reason()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id,
                SubmissionState.Rejected,
                ReviewDecisionRecordingTests.MaintainerAccountId,
                "The highlight coordinates for U8 look wrong.",
                ReviewDecisionRecordingTests.Now,
                CancellationToken.None);

            RecordedDecision decision = store.Decisions[id];

            Assert.Equal(SubmissionState.Rejected, decision.State);
            Assert.Equal(ReviewDecisionRecordingTests.MaintainerAccountId, decision.DecidedByAccountId);
            Assert.Equal("The highlight coordinates for U8 look wrong.", decision.Comment);
        }

        [Fact]
        public async Task The_SUBMISSION_ROW_itself_moves_to_the_new_state()
        {
            // The queue is filtered on `state`, so a decision that recorded the audit row without
            // moving the submission would leave it sitting in the queue forever, decided.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id, SubmissionState.ChangesRequested, 7, "Please add the region.",
                ReviewDecisionRecordingTests.Now, CancellationToken.None);

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

            Assert.Equal(SubmissionState.ChangesRequested, record!.State);
            Assert.Equal(ReviewDecisionRecordingTests.Now, record.DecidedUtc);
        }

        [Fact]
        public async Task A_DECIDED_submission_leaves_the_queue()
        {
            // The end-to-end consequence, asserted rather than assumed: the queue is `pending`
            // only, so deciding must remove it. A maintainer who decides something and still sees it
            // next refresh will decide it again.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            Assert.Single(await store.GetQueueAsync(100, CancellationToken.None));

            await store.SetDecisionAsync(
                id, SubmissionState.Rejected, 7, "Not this time, sorry.",
                ReviewDecisionRecordingTests.Now, CancellationToken.None);

            Assert.Empty(await store.GetQueueAsync(100, CancellationToken.None));
        }

        [Fact]
        public async Task Deciding_a_submission_that_does_NOT_EXIST_throws()
        {
            // The fake is stricter than convenient on purpose. The real statement is an UPDATE
            // against an existing row: against MariaDB a bad id updates nothing and reports
            // success, so a permissive fake would certify a caller that silently does nothing.
            var store = new FakeSubmissionStore();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                store.SetDecisionAsync(
                    999, SubmissionState.Rejected, 7, "A reason.",
                    ReviewDecisionRecordingTests.Now, CancellationToken.None));
        }

        [Fact]
        public async Task A_BLANK_comment_is_stored_as_NULL_rather_than_an_empty_string()
        {
            // So "no comment was given" and "an empty comment was given" stay distinguishable in
            // the audit trail. The endpoints refuse a blank reason for the two outcomes that need
            // one, but an approval may legitimately carry none.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id, SubmissionState.Approved, 7, "   ",
                ReviewDecisionRecordingTests.Now, CancellationToken.None);

            Assert.True(string.IsNullOrWhiteSpace(store.Decisions[id].Comment));
        }

        [Fact]
        public async Task The_LATEST_decision_wins_when_one_is_recorded_twice()
        {
            // Not a supported flow - ReviewDecisionRules refuses a second decision - but the store
            // must behave predictably if it ever happens, rather than keeping the first silently.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id, SubmissionState.ChangesRequested, 7, "First look.",
                ReviewDecisionRecordingTests.Now, CancellationToken.None);

            await store.SetDecisionAsync(
                id, SubmissionState.Rejected, 8, "Second look.",
                ReviewDecisionRecordingTests.Now.AddHours(1), CancellationToken.None);

            Assert.Equal(SubmissionState.Rejected, store.Decisions[id].State);
            Assert.Equal(8, store.Decisions[id].DecidedByAccountId);
        }

        // -----------------------------------------------------------------------------------
        // What the CONTRIBUTOR can read back.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_maintainers_COMMENT_reaches_the_contributors_own_record()
        {
            // *** THE POINT OF THE WHOLE "REQUEST CHANGES" OUTCOME. *** Contributing needs no
            // account, so there is no inbox and no thread - the contact email and this sentence
            // are the entire channel back to the person who did the work. A comment recorded only
            // in an audit table would reach nobody.
            //
            // The contributor's status endpoint reads it off this record and returns it as
            // `maintainerComment`, a field reserved from the start for exactly this moment.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id,
                SubmissionState.ChangesRequested,
                ReviewDecisionRecordingTests.MaintainerAccountId,
                "The highlight for U8 looks like it is on the wrong pin - could you check?",
                ReviewDecisionRecordingTests.Now,
                CancellationToken.None);

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

            Assert.Equal(
                "The highlight for U8 looks like it is on the wrong pin - could you check?",
                record!.DecisionComment);
        }

        [Fact]
        public async Task An_UNDECIDED_submission_carries_NO_comment()
        {
            // So the contributor's screen can tell "not looked at yet" from "looked at and
            // returned without a word" - the second would be a defect worth noticing.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

            Assert.Null(record!.DecisionComment);
        }

        // -----------------------------------------------------------------------------------
        // The rules and the recording, together - what an endpoint actually does.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_submission_decided_ONCE_cannot_be_decided_AGAIN()
        {
            // *** THE INTERLOCK, exercised as a sequence rather than as a table of states. *** Two
            // maintainers with the queue open both act on the same row; the second must be refused
            // on the state the first one wrote, not on the state their screen was showing.
            (FakeSubmissionStore store, long id) = await ReviewDecisionRecordingTests.PendingAsync();

            await store.SetDecisionAsync(
                id, SubmissionState.Rejected, 7, "Already handled.",
                ReviewDecisionRecordingTests.Now, CancellationToken.None);

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

            Assert.False(ReviewDecisionRules.CanReject(
                ReviewDecisionRecordingTests.Maintainer(), record, out string why));

            Assert.False(ReviewDecisionRules.CanRequestChanges(
                ReviewDecisionRecordingTests.Maintainer(), record, out _));

            Assert.NotEmpty(why);
        }

        // A maintainer OF THE FIXTURE'S SYSTEM - authority is per system since Phase 6.
        private static ReviewAccess Maintainer() =>
            ReviewAccess.For(
                new Handlers.Accounts.AccountRecord(
                    Id: ReviewDecisionRecordingTests.MaintainerAccountId,
                    Email: "maintainer@example.com",
                    NormalisedEmail: "maintainer@example.com",
                    PasswordHash: "hash",
                    DisplayName: "Maintainer",
                    IsVerified: true,
                    IsAdministrator: false,
                    IsLocked: false,
                    CreatedUtc: ReviewDecisionRecordingTests.Now,
                    LastLoginUtc: null),
                ["Commodore/C64/250407"]);
    }
}
