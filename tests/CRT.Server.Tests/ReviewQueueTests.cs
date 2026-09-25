using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the review QUEUE - which submissions are waiting, and in what order.
    //
    // Both properties are real decisions rather than incidental behaviour, and both are the kind
    // that fail quietly: a queue that includes the wrong states puts a reviewer in front of a
    // half-uploaded contribution, and a queue in the wrong order lets a steady trickle of new
    // submissions bury the one that has been waiting a month. Neither throws.
    //
    // These run against FakeSubmissionStore, which mirrors the real SQL's filter and ordering -
    // see its own header on why a fake that is more permissive than the real store certifies bugs.
    // ###########################################################################################
    public sealed class ReviewQueueTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private static async Task<long> AddAsync(FakeSubmissionStore store, string state, string summary = "A change.")
        {
            long id = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", summary,
                    1, [], ReviewQueueTests.Now, ReviewQueueTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, state, ReviewQueueTests.Now, CancellationToken.None);

            return id;
        }

        [Fact]
        public async Task Only_PENDING_submissions_are_queued()
        {
            // *** THE FILTER IS A DECISION. *** 'uploading' is a contribution still arriving - a
            // reviewer acting on one would be deciding about a half-delivered submission. Every
            // other state has already been decided. 'pending' is the only state meaning "arrived
            // intact and is somebody's to decide", which is what a queue is.
            var store = new FakeSubmissionStore();

            long pending = await ReviewQueueTests.AddAsync(store, SubmissionState.Pending);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Uploading);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Abandoned);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Rejected);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Merged);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Approved);
            await ReviewQueueTests.AddAsync(store, SubmissionState.Withdrawn);
            await ReviewQueueTests.AddAsync(store, SubmissionState.ChangesRequested);

            IReadOnlyList<SubmissionRecord> queue = await store.GetQueueAsync(100, CancellationToken.None);

            Assert.Equal(pending, Assert.Single(queue).Id);
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
        public async Task A_queued_row_carries_what_the_reviewer_needs_to_reply()
        {
            // Contributors have no account, so the contact address is the ENTIRE channel back to
            // them. A queue row without it would leave a reviewer able to reject something with no
            // way to say why.
            var store = new FakeSubmissionStore();

            await ReviewQueueTests.AddAsync(store, SubmissionState.Pending, "Corrected R12.");

            SubmissionRecord row = Assert.Single(await store.GetQueueAsync(100, CancellationToken.None));

            Assert.Equal("someone@example.com", row.ContactEmail);
            Assert.Equal("Corrected R12.", row.Summary);
            Assert.Equal("Commodore/C64/250407", row.SystemId);
        }
    }
}
