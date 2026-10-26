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
            File.WriteAllText(Path.Combine(folder, DataTreeLayout.BoardFileName), "a board file");
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

        // ###########################################################################################
        // *** THE AMENDMENT VERSION IS ON THE ROW (code review, 2026-10-13). *** Another maintainer's
        // amendment changed nothing a row carried, so CRT never read the open submission again -
        // and an Approve published an amendment its approver had never seen.
        // ###########################################################################################
        [Fact]
        public async Task A_row_carries_its_amendment_version_and_0_when_nobody_amended_it()
        {
            var store = new FakeSubmissionStore();
            long untouched = await ReviewQueueFlowTests.AddAsync(store);
            long amended = await ReviewQueueFlowTests.AddAsync(store);

            store.Amendments.Add((amended, 1, new SubmissionManifest(), "Anna", ReviewQueueFlowTests.Now));
            store.Amendments.Add((amended, 2, new SubmissionManifest(), "Bo", ReviewQueueFlowTests.Now));

            ReviewQueueAnswer answer = await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store);

            Assert.Equal(0, answer.Submissions.Single(entry => entry.Id == untouched).AmendmentVersion);
            Assert.Equal(2, answer.Submissions.Single(entry => entry.Id == amended).AmendmentVersion);
        }

        // ###########################################################################################
        // *** MADE FROM OLDER DATA THAN BETA HAS NOW (code review, 2026-10-13). *** A draft is a whole
        // copy of the board; one made before the board's latest publish carries the older rows and
        // would remove what was published since.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_made_from_an_older_revision_than_BETAs_is_flagged()
        {
            var store = new FakeSubmissionStore();
            long id = await ReviewQueueFlowTests.AddAsync(store);

            store.Submissions[id] = store.Submissions[id] with { BaseRevision = "2026-September-1" };
            store.PublishedBoards["Commodore/C64/250407"] = new PublishedBoardRow("2026-October-4", "hash", ReviewQueueFlowTests.Now);

            ReviewQueueEntry entry = Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store)).Submissions);

            Assert.True(entry.BasedOnOlderRevision);
        }

        // ###########################################################################################
        // *** THE SAME CONTRIBUTOR'S NEWER ONE BESIDE A HALF-APPROVED ONE (owner decision,
        // 2026-10-13). *** The older one stays (the replacement rule); its row names the newer one,
        // and the newer one's row names nothing.
        // ###########################################################################################
        [Fact]
        public async Task A_half_approved_submission_names_the_newer_one_its_contributor_sent()
        {
            var store = new FakeSubmissionStore();
            long older = await ReviewQueueFlowTests.AddAsync(store);
            store.Submissions[older] = store.Submissions[older] with { State = SubmissionState.Approved };
            long newer = await ReviewQueueFlowTests.AddAsync(store);

            ReviewQueueAnswer answer = await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store);

            ReviewQueueEntry olderEntry = answer.Submissions.Single(entry => entry.Id == older);
            ReviewQueueEntry newerEntry = answer.Submissions.Single(entry => entry.Id == newer);

            Assert.Equal(newer, olderEntry.NewerSubmissionId);
            Assert.Equal(store.Submissions[newer].CreatedUtc, olderEntry.NewerSubmissionUtc);
            Assert.Null(newerEntry.NewerSubmissionId);
            Assert.Null(newerEntry.NewerSubmissionUtc);
        }

        // ###########################################################################################
        // *** THE ROW AND THE DETAIL JUDGE AGAINST ONE REVISION (code review, 2026-10-14). *** The row
        // read boards.current_revision alone and the detail BETA's workbook, so a shipped board no
        // publish has recorded (no current_revision) was flagged in the detail and not on its row.
        // Both now take BETA's workbook's own date, or the record's without one.
        // ###########################################################################################
        [Fact]
        public async Task A_shipped_board_no_publish_recorded_is_judged_against_BETAs_workbook_on_its_row()
        {
            this.PublishWorkbook("Commodore/C64/250407", "2026-October-4");

            var store = new FakeSubmissionStore();
            long id = await ReviewQueueFlowTests.AddAsync(store);
            store.Submissions[id] = store.Submissions[id] with { BaseRevision = "2026-September-1" };

            ReviewQueueEntry entry = Assert.Single((await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store)).Submissions);

            Assert.True(entry.BasedOnOlderRevision);
        }

        [Fact]
        public void BETAs_own_workbook_outranks_the_record_and_the_record_stands_in_without_one()
        {
            this.PublishWorkbook("Commodore/C64/250407", "2026-October-4");

            var lagging = new BoardRecord("Commodore/C64/250407", "Commodore", "C64", "250407", CurrentRevision: "2026-September-1", IsAccepting: true);

            Assert.Equal(
                "2026-October-4",
                ReviewQueueFlow.BetaRevisionOf(PublishedBoardLocator.LocateBoard(this.thisTree, "Commodore/C64/250407"), lagging));

            Assert.Equal(
                "2026-September-1",
                ReviewQueueFlow.BetaRevisionOf(PublishedBoardLocator.LocateBoard(this.thisTree, "Commodore/C128/310378"), lagging));

            Assert.Null(ReviewQueueFlow.BetaRevisionOf(PublishedBoardLocation.None, board: null));
        }

        // A real board file in the BETA tree, carrying a revision date.
        private void PublishWorkbook(string boardId, string revision)
        {
            string[] parts = boardId.Split('/');
            string path = Path.Combine([this.thisTree, .. parts, DataTreeLayout.BoardFileName]);
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);

            BoardWorkbookWriter.Write(path, new BoardData
            {
                RevisionDate = revision,
                Components = [new ComponentEntry { BoardLabel = "U8" }],
            });
        }

        [Theory]
        [InlineData("2026-October-4", "2026-October-4", false)]
        [InlineData("2026-September-1", "2026-October-4", true)]
        [InlineData("2026-October-4", "2026-September-1", false)]
        [InlineData("", "2026-October-4", null)]
        [InlineData("2026-October-4", null, null)]
        [InlineData("see changelog", "2026-October-4", false)]
        public void Older_data_is_said_only_when_BETAs_revision_is_known_to_be_later(string baseRevision, string? current, bool? expected)
        {
            Assert.Equal(expected, ReviewQueueFlow.IsBasedOnOlderRevision(baseRevision, current));
        }

        // ###########################################################################################
        // *** A MAINTAINER'S QUEUE IS FILTERED IN THE QUERY (code review, 2026-10-13). *** It read the
        // oldest hundred rows of EVERY board and filtered them afterwards, so a quiet board's
        // submission behind a hundred others was never shown to its maintainer.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainers_submission_behind_a_hundred_others_is_still_in_their_queue()
        {
            var store = new FakeSubmissionStore();

            for (int i = 0; i < ReviewEndpoints.DefaultQueueLimit; i++)
                await ReviewQueueFlowTests.AddAsync(store, "Commodore/C128/310378");

            long mine = await ReviewQueueFlowTests.AddAsync(store);

            ReviewAccess maintainer = ReviewQueueFlowTests.MaintainerOfC64();

            ReviewQueueAnswer answer = await ReviewQueueFlow.BuildAsync(
                maintainer,
                await ReviewQueueFlow.ReadAsync(maintainer, store, ReviewEndpoints.DefaultQueueLimit, CancellationToken.None),
                this.thisTree,
                store,
                new FakeAccountStore(),
                CancellationToken.None);

            Assert.Equal([mine], answer.Submissions.Select(entry => entry.Id));
        }

        // One approvals read for the whole queue, however long - it was one per row, every minute,
        // from every open Maintainer tab.
        [Fact]
        public async Task The_approvals_of_a_whole_queue_are_read_at_once()
        {
            var store = new FakeSubmissionStore();

            for (int i = 0; i < 5; i++)
                await ReviewQueueFlowTests.AddAsync(store);

            await this.QueueAsync(ReviewQueueFlowTests.Administrator(), store);

            Assert.Equal(1, store.ApprovalReads);
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
