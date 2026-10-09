using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewQueueFlow - the review queue as one account sees it: which rows, and the two
    // badges each carries (owner request, 2026-09-26): "New system" when the system has no
    // published board, and "Awaiting your review" while this account's approval is still needed.
    //
    // The badges are the detail's own answers (PublishedBoardLocator, ApprovalStatus.CanApprove),
    // so a row cannot promise what opening the submission then contradicts. A real temp folder
    // stands in for the published tree - which workbook is there is what is under test.
    // ###########################################################################################
    public sealed class ReviewQueueFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 26, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisTree;

        public ReviewQueueFlowTests()
        {
            this.thisTree = Path.Combine(Path.GetTempPath(), "crt-queue-flow-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(this.thisTree);
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(this.thisTree))
                    Directory.Delete(this.thisTree, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private static readonly AccountRecord AdministratorAccount =
            new(6, "a@example.com", "a@example.com", "hash", "A", true, true, false, ReviewQueueFlowTests.Now, null);

        private static readonly AccountRecord MaintainerAccount =
            new(5, "m@example.com", "m@example.com", "hash", "M", true, false, false, ReviewQueueFlowTests.Now, null);

        private static ReviewAccess Administrator() => ReviewAccess.For(ReviewQueueFlowTests.AdministratorAccount);

        private static ReviewAccess MaintainerOfC64() => ReviewAccess.For(ReviewQueueFlowTests.MaintainerAccount, ["Commodore/C64/250407"]);

        private void Publish(string boardId)
        {
            string folder = Path.Combine([this.thisTree, .. boardId.Split('/')]);
            Directory.CreateDirectory(folder);
            File.WriteAllText(Path.Combine(folder, $"Data {boardId.Split('/')[1]} {boardId.Split('/')[2]}.xlsx"), "a workbook");
        }

        private static async Task<long> AddAsync(FakeSubmissionStore store, string boardId = "Commodore/C64/250407", bool touchesSharedFiles = false)
        {
            string[] parts = boardId.Split('/');

            long id = await store.CreateAsync(
                new NewSubmission(
                    boardId, parts[0], parts[1], parts[2],
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "A change.",
                    1, [], ReviewQueueFlowTests.Now, ReviewQueueFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, SubmissionState.Pending, ReviewQueueFlowTests.Now, CancellationToken.None);

            if (touchesSharedFiles)
                store.Submissions[id] = store.Submissions[id] with { TouchesSharedFiles = true };

            return id;
        }

        private async Task<ReviewQueueAnswer> QueueAsync(ReviewAccess access, FakeSubmissionStore store, FakeAccountStore? accounts = null) =>
            await ReviewQueueFlow.BuildAsync(
                access,
                await store.GetQueueAsync(100, CancellationToken.None),
                this.thisTree,
                store,
                accounts ?? new FakeAccountStore(),
                CancellationToken.None);

        [Fact]
        public async Task A_maintainer_is_given_only_their_own_boards_submissions()
        {
            var store = new FakeSubmissionStore();

            long mine = await ReviewQueueFlowTests.AddAsync(store);
            await ReviewQueueFlowTests.AddAsync(store, "Commodore/C128/310378");

            ReviewQueueAnswer answer = await this.QueueAsync(ReviewQueueFlowTests.MaintainerOfC64(), store);

            Assert.Equal([mine], answer.Submissions.Select(entry => entry.Id));
            Assert.Equal(1, answer.Count);
            Assert.False(answer.IsAdministrator);
            Assert.True((await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store)).IsAdministrator);
        }

        // "New board" is a board with no published board - the highest-risk submission there is.
        [Fact]
        public async Task A_board_with_no_published_board_is_new_and_one_with_a_board_is_not()
        {
            var store = new FakeSubmissionStore();
            this.Publish("Commodore/C64/250407");

            long published = await ReviewQueueFlowTests.AddAsync(store);
            long brandNew = await ReviewQueueFlowTests.AddAsync(store, "Commodore/C65/Prototype");

            ReviewQueueAnswer answer = await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store);

            Assert.False(answer.Submissions.Single(entry => entry.Id == published).IsNewBoard);
            Assert.True(answer.Submissions.Single(entry => entry.Id == brandNew).IsNewBoard);
        }

        // ###########################################################################################
        // "Awaiting your review" while this account's approval is still needed - and not once it
        // has given it, while the shared-file change waits for the board's maintainer.
        // ###########################################################################################
        [Fact]
        public async Task It_awaits_you_until_you_have_approved_and_then_awaits_the_other_approver()
        {
            var store = new FakeSubmissionStore();
            var accounts = new FakeAccountStore();

            accounts.Accounts[ReviewQueueFlowTests.MaintainerAccount.Id] = ReviewQueueFlowTests.MaintainerAccount;
            accounts.Maintainers.Add(("Commodore/C64/250407", ReviewQueueFlowTests.MaintainerAccount.Id));

            long id = await ReviewQueueFlowTests.AddAsync(store, touchesSharedFiles: true);

            Assert.True(Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store, accounts)).Submissions).AwaitsYou);

            // The administrator approves - the first of the two.
            await store.AddApprovalAsync(id, ApproverRole.Administrator, ReviewQueueFlowTests.AdministratorAccount.Id, "A", ReviewQueueFlowTests.Now);
            await store.SetStateAsync(id, SubmissionState.Approved, ReviewQueueFlowTests.Now, CancellationToken.None);

            Assert.False(Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store, accounts)).Submissions).AwaitsYou);
            Assert.True(Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.MaintainerOfC64(), store, accounts)).Submissions).AwaitsYou);
        }

        // An ordinary submission - one approval publishes it - awaits whoever may decide it.
        [Fact]
        public async Task An_ordinary_submission_awaits_its_boards_maintainer()
        {
            var store = new FakeSubmissionStore();
            await ReviewQueueFlowTests.AddAsync(store);

            Assert.True(Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.MaintainerOfC64(), store)).Submissions).AwaitsYou);
        }

        // *** THE UPLOAD TOKEN HASH IS NEVER IN A ROW. *** It is the contributor's capability for the
        // submission; the answer's type has nowhere to put it.
        [Fact]
        public void A_queue_row_has_no_place_for_the_contributors_upload_token()
        {
            Assert.DoesNotContain(
                typeof(ReviewQueueEntry).GetProperties(),
                property => property.Name.Contains("Token", StringComparison.OrdinalIgnoreCase));
        }
    }
}
