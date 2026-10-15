using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewDecisionFlow - Reject and Request changes (code review, 2026-10-13).
    //
    // The two properties that were missing: a decision may not land on a submission that an approval
    // is publishing (the write was `WHERE id = @id` with no state, outside the publish lock), and a
    // decision made on an amendment version the maintainer never saw is refused.
    // ###########################################################################################
    public sealed class ReviewDecisionFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 13, 12, 0, 0, TimeSpan.Zero);

        private const string BoardId = "Commodore/C64/250407";

        private const string Reason = "The part number on U8 is the C128's, not the C64's.";

        private static ReviewAccess Maintainer(params string[] boards) =>
            ReviewAccess.For(
                new AccountRecord(7, "maintainer@example.com", "maintainer@example.com", "hash", "Maintainer",
                    IsVerified: true, IsAdministrator: false, IsLocked: false, ReviewDecisionFlowTests.Now, null),
                boards);

        private static async Task<(FakeSubmissionStore Store, long Id)> PendingAsync()
        {
            var store = new FakeSubmissionStore();

            long id = await store.CreateAsync(new NewSubmission(
                ReviewDecisionFlowTests.BoardId, "Commodore", "C64", "250407",
                null, "someone@example.com", "192.0.2.1", "hash", "2026-September-1", "U8 fix", 1, [],
                ReviewDecisionFlowTests.Now, ReviewDecisionFlowTests.Now.AddDays(1)));

            await store.SetStateAsync(id, SubmissionState.Pending, ReviewDecisionFlowTests.Now);

            return (store, id);
        }

        private static Task<DecisionOutcome> RejectAsync(
            FakeSubmissionStore store,
            long id,
            PublishLock? publishLock = null,
            int? expectedVersion = null,
            string comment = ReviewDecisionFlowTests.Reason,
            ReviewAccess? access = null) =>
            ReviewDecisionFlow.DecideAsync(
                access ?? ReviewDecisionFlowTests.Maintainer(ReviewDecisionFlowTests.BoardId),
                id,
                new ReviewDecisionRequest(comment, ExpectedVersion: expectedVersion),
                SubmissionState.Rejected,
                ReviewDecisionRules.CanReject,
                store,
                publishLock ?? new PublishLock(),
                ReviewDecisionFlowTests.Now);

        [Fact]
        public async Task A_rejection_is_recorded_with_who_and_why()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id);

            Assert.True(outcome.IsDecided, outcome.Error);
            Assert.Equal(SubmissionState.Rejected, store.Submissions[id].State);
            Assert.Equal(ReviewDecisionFlowTests.Reason, store.Decisions[id].Comment);
            Assert.Equal(7, store.Decisions[id].DecidedByAccountId);
        }

        // ###########################################################################################
        // *** THE RACE THE CONDITIONAL WRITE CLOSES. *** An approval publishes the submission between
        // the decision's own check and its write. The write used to land anyway - 'merged' became
        // 'rejected' with the board already in BETA, and the contributor was told it was refused.
        // ###########################################################################################
        [Fact]
        public async Task A_rejection_landing_while_an_approval_published_the_submission_changes_nothing()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();

            store.BeforeConditionalWrite = () =>
            {
                store.Submissions[id] = store.Submissions[id] with { State = SubmissionState.Merged };
                return Task.CompletedTask;
            };

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id);

            Assert.False(outcome.IsDecided);
            Assert.True(outcome.IsConflict);
            Assert.Equal(ReviewDecisionRules.DecidedMeanwhileMessage, outcome.Error);
            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);
            Assert.False(store.Decisions.ContainsKey(id));
        }

        // The decision waits for a publish in progress - the lock every approval holds while it
        // writes - and only then looks at the submission.
        [Fact]
        public async Task A_decision_waits_for_the_publish_lock()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();
            var publishLock = new PublishLock();

            Task<DecisionOutcome> deciding;

            using (await publishLock.EnterAsync())
            {
                deciding = ReviewDecisionFlowTests.RejectAsync(store, id, publishLock);

                await Task.Delay(100);
                Assert.False(deciding.IsCompleted);

                // The approval finishes while the decision waits.
                store.Submissions[id] = store.Submissions[id] with { State = SubmissionState.Merged };
            }

            DecisionOutcome outcome = await deciding;

            Assert.True(outcome.IsConflict);
            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_decision_on_an_amendment_version_never_shown_is_refused_as_a_conflict()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();
            store.Amendments.Add((id, 1, new SubmissionManifest(), "Another Maintainer", ReviewDecisionFlowTests.Now));

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id, expectedVersion: 0);

            Assert.True(outcome.IsConflict);
            Assert.Equal(ReviewDecisionRules.AmendedSinceMessage, outcome.Error);
            Assert.Equal(SubmissionState.Pending, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_decision_on_the_version_shown_is_recorded()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();
            store.Amendments.Add((id, 1, new SubmissionManifest(), "Another Maintainer", ReviewDecisionFlowTests.Now));

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id, expectedVersion: 1);

            Assert.True(outcome.IsDecided, outcome.Error);
        }

        // An older CRT sends no version - nothing is checked, as before.
        [Fact]
        public async Task A_decision_without_a_version_is_not_checked_against_one()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();
            store.Amendments.Add((id, 2, new SubmissionManifest(), "Another Maintainer", ReviewDecisionFlowTests.Now));

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id, expectedVersion: null);

            Assert.True(outcome.IsDecided, outcome.Error);
        }

        [Fact]
        public async Task A_rejection_without_a_reason_is_refused_and_nothing_is_recorded()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(store, id, comment: "  ");

            Assert.False(outcome.IsDecided);
            Assert.False(outcome.IsConflict);
            Assert.Equal(SubmissionState.Pending, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_maintainer_of_another_board_is_forbidden()
        {
            (FakeSubmissionStore store, long id) = await ReviewDecisionFlowTests.PendingAsync();

            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(
                store, id, access: ReviewDecisionFlowTests.Maintainer("Commodore/C128/310378"));

            Assert.True(outcome.IsForbidden);
            Assert.Equal(SubmissionState.Pending, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_submission_that_does_not_exist_is_not_found()
        {
            DecisionOutcome outcome = await ReviewDecisionFlowTests.RejectAsync(new FakeSubmissionStore(), 404);

            Assert.True(outcome.IsNotFound);
        }
    }
}
