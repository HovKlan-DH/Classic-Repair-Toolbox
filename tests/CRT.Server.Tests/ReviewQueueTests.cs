using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the review QUEUE - which submissions are waiting, and in what order.
    //
    // Both properties are real decisions rather than incidental behaviour, and both are the kind
    // that fail quietly: a queue that includes the wrong states puts a maintainer in front of a
    // half-uploaded contribution, and a queue in the wrong order lets a steady trickle of new
    // submissions bury the one that has been waiting a month. Neither throws.
    //
    // These run against FakeSubmissionStore, which mirrors the real SQL's filter and ordering -
    // see its own header on why a fake that is more permissive than the real store certifies bugs.
    // ###########################################################################################
    public sealed class ReviewQueueTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private static async Task<long> AddAsync(
            FakeSubmissionStore store, string state, string summary = "A change.", string boardId = "Commodore/C64/250407")
        {
            string[] parts = boardId.Split('/');

            long id = await store.CreateAsync(
                new NewSubmission(
                    boardId, parts[0], parts[1], parts[2],
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", summary,
                    1, [], ReviewQueueTests.Now, ReviewQueueTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, state, ReviewQueueTests.Now, CancellationToken.None);

            return id;
        }

        [Fact]
        public async Task Only_PENDING_and_HALF_APPROVED_submissions_are_queued()
        {
            // *** THE FILTER IS A DECISION. *** 'uploading' is a contribution still arriving - a
            // maintainer acting on one would be deciding about a half-delivered submission. 'pending'
            // means "arrived intact and is somebody's to decide", which is what a queue is.
            //
            // 'approved' joined it on 2026-09-25: it is where the FIRST of the two approvals a
            // shared-file change needs leaves a submission. Leaving it out would hide it from the
            // very person whose approval it now waits for.
            var store = new FakeSubmissionStore();

            long pending = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Uploading);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Abandoned);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Rejected);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Merged);
            long half = await ReviewQueueTests.AddAsync(store, SubmissionState.Approved);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Withdrawn);
            await ReviewQueueTests.AddAsync(store, SubmissionState.ChangesRequested);

            IReadOnlyList<SubmissionRecord> queue = await store.GetQueueAsync(100, CancellationToken.None);

            Assert.Equal([pending, half], queue.Select(record => record.Id));
        }

        // ###########################################################################################
        // THE QUEUE A MAINTAINER SEES IS FILTERED BY ReviewAuthority (Phase 6 roles). The endpoint
        // applies CanReview row by row; this pins the rule over the rows the store returns, so the
        // two halves are tested against the same records.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainer_sees_only_their_own_boards_INCLUDING_a_shared_files_one()
        {
            var store = new FakeSubmissionStore();

            long mine = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Mine.");
            long theirs = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Theirs.", "Commodore/C128/310378");
            long shared = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Shared.");
            store.Submissions[shared] = store.Submissions[shared] with { TouchesSharedFiles = true };

            var maintainer = ReviewAccess.For(
                new Handlers.Accounts.AccountRecord(
                    5, "r@example.com", "r@example.com", "hash", "R", true, false, false, ReviewQueueTests.Now, null),
                ["Commodore/C64/250407"]);

            IReadOnlyList<SubmissionRecord> queue = await store.GetQueueAsync(100, CancellationToken.None);

            // The shared-files one too, since 2026-09-25: the maintainer's approval is one of the two
            // it needs.
            Assert.Equal(
                [mine, shared],
                queue.Where(record => ReviewAuthority.CanReview(maintainer, record)).Select(record => record.Id));

            // The administrator sees all three.
            var admin = ReviewAccess.For(
                new Handlers.Accounts.AccountRecord(
                    6, "a@example.com", "a@example.com", "hash", "A", true, true, false, ReviewQueueTests.Now, null));

            Assert.Equal(
                [mine, theirs, shared],
                queue.Where(record => ReviewAuthority.CanReview(admin, record)).Select(record => record.Id));
        }

        [Fact]
        public async Task The_queue_is_OLDEST_FIRST()
        {
            // *** THE ORDERING IS A DECISION TOO, and the opposite of "my submissions". *** A work
            // queue is worked from the front, so the longest-waiting submission must lead.
            // Newest-first would let a steady trickle of new contributions keep burying the one
            // that has been waiting longest - which is how a contribution quietly never gets
            // reviewed, and the contributor concludes the project ignored them.
            var store = new FakeSubmissionStore();

            long first = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Oldest.");
            long second = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Middle.");
            long third = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Newest.");

            IReadOnlyList<SubmissionRecord> queue = await store.GetQueueAsync(100, CancellationToken.None);

            Assert.Equal([first, second, third], queue.Select(record => record.Id));
        }

        [Fact]
        public async Task An_empty_queue_is_empty_rather_than_an_error()
        {
            // The ordinary state of a healthy project, and what the window shows on first launch.
            var store = new FakeSubmissionStore();

            Assert.Empty(await store.GetQueueAsync(100, CancellationToken.None));
        }

        [Fact]
        public async Task A_queue_of_only_decided_submissions_is_empty()
        {
            var store = new FakeSubmissionStore();

            await ReviewQueueTests.AddAsync(store, SubmissionState.Merged);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Rejected);

            Assert.Empty(await store.GetQueueAsync(100, CancellationToken.None));
        }

        [Fact]
        public async Task The_limit_takes_the_OLDEST_submissions_not_the_newest()
        {
            // Truncating from the wrong end would hide exactly the submissions the ordering above
            // exists to surface - the ones that have waited longest.
            var store = new FakeSubmissionStore();

            long first = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);
            long second = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);

            IReadOnlyList<SubmissionRecord> queue = await store.GetQueueAsync(2, CancellationToken.None);

            Assert.Equal([first, second], queue.Select(record => record.Id));
        }

        [Fact]
        public async Task A_nonsense_limit_does_not_return_everything_or_nothing()
        {
            // The store clamps rather than trusting its caller: a zero would make the queue look
            // empty when it is not, which reads as "nothing to review" and is the worst possible
            // wrong answer here.
            var store = new FakeSubmissionStore();

            await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);

            Assert.Single(await store.GetQueueAsync(0, CancellationToken.None));
            Assert.Single(await store.GetQueueAsync(-5, CancellationToken.None));
        }

        [Fact]
        public async Task A_queued_row_carries_what_the_maintainer_needs_to_reply()
        {
            // Contributors have no account, so the contact address is the ENTIRE channel back to
            // them. A queue row without it would leave a maintainer able to reject something with no
            // way to say why.
            var store = new FakeSubmissionStore();

            await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Corrected R12.");

            SubmissionRecord row = Assert.Single(await store.GetQueueAsync(100, CancellationToken.None));

            Assert.Equal("someone@example.com", row.ContactEmail);
            Assert.Equal("Corrected R12.", row.Summary);
            Assert.Equal("Commodore/C64/250407", row.BoardId);
        }
    }
}
