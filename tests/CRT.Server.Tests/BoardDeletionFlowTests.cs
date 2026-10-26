using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers deleting a board completely (owner request, 2026-10-03: "remove it from everywhere -
    // including BETA and stable sources"), the administrator's Account > "Delete a board".
    //
    // Asserted on the DISK and in the stores, like BetaRollbackFlowTests: what matters is what the
    // two trees, their drop-down lists, their manifests and the database hold afterwards - and,
    // for every refusal, that NOTHING moved.
    //
    // The board under test is "Commodore/C64/999999", beside a board that stays ("Commodore/C64/
    // 250407", DataTreeBuilder's) in both trees - so a delete taking too much is seen too.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class BoardDeletionFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 3, 12, 0, 0, TimeSpan.Zero);

        private const string BoardId = "Commodore/C64/999999";
        private const string BoardFile = "Commodore/C64/999999/board.json";
        private const string ImageName = "Images/sheet.png";
        private const string ImageContent = "the sheet image";
        private const string SharedName = "Manual.pdf";
        private const string SharedContent = "the manual both boards list";
        private const string OtherBoard = "Commodore/C64/250407";

        // Where the two files' bytes are stored ("Files/<ab>/<sha256>.<ext>"): the image only this
        // board lists, the manual both boards list.
        private static readonly string Image = DataTreeBuilder.Stored(BoardDeletionFlowTests.ImageContent, BoardDeletionFlowTests.ImageName);
        private static readonly string Shared = DataTreeBuilder.Stored(BoardDeletionFlowTests.SharedContent, BoardDeletionFlowTests.SharedName);

        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisProduction;

        public BoardDeletionFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-board-delete", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(this.thisRoot, "app-data-BETA", "Data");
            this.thisProduction = Path.Combine(this.thisRoot, "app-data", "Data");

            Directory.CreateDirectory(this.thisBeta);
            Directory.CreateDirectory(this.thisProduction);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless.
            }
        }

        private ServerOptions Options(bool production = true) => new()
        {
            DataTreeRoot = this.thisBeta,
            ManifestPath = this.BetaManifest,
            PublicDataBaseUrl = "https://example.com/app-data-BETA/Data",
            ProductionTreeRoot = Path.Combine(this.thisRoot, "app-data"),
            ProductionDataTreeRoot = production ? this.thisProduction : null,
            ProductionManifestPath = production ? this.ProductionManifest : null,
            ProductionPublicDataBaseUrl = production ? "https://example.com/app-data/Data" : null
        };

        private string BetaManifest => Path.Combine(this.thisRoot, "app-data-BETA", "dataChecksums.json");

        private string ProductionManifest => Path.Combine(this.thisRoot, "app-data", "dataChecksums.json");

        private static AccountRecord Account(long id = 1, bool administrator = true, string? email = null) =>
            new(id, email ?? $"a{id}@example.com", email ?? $"a{id}@example.com", "hash", $"A{id}",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false,
                BoardDeletionFlowTests.Now, null);

        private static ReviewAccess Admin() => ReviewAccess.For(BoardDeletionFlowTests.Account());

        private sealed record Fixture(
            FakeSubmissionStore Store,
            FakeAccountStore Accounts,
            FakeBoardViewStore Views,
            FakeEmailSender Mail,
            BoardDeletionFlow Flow);

        private string Blobs => Path.Combine(this.thisRoot, "blobs");

        private Fixture Make(FakeSubmissionStore? store = null)
        {
            var submissions = store ?? new FakeSubmissionStore();
            var accounts = new FakeAccountStore();
            var views = new FakeBoardViewStore();
            var mail = new FakeEmailSender();

            var flow = new BoardDeletionFlow(
                submissions,
                accounts,
                views,
                new SubmissionNotifier(mail, NullLogger<SubmissionNotifier>.Instance),
                new PublishLock(),
                new BlobStore(this.Blobs, NullLogger<BlobStore>.Instance),
                NullLogger<BoardDeletionFlow>.Instance);

            return new Fixture(submissions, accounts, views, mail, flow);
        }

        // A partial upload of `submissionId` on disk, as an upload that stopped half way leaves it.
        private string PartialUpload(long submissionId)
        {
            Assert.True(BlobStorePaths.TryGetPartialPath(this.Blobs, submissionId, new string('a', 64), out string path));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, [1, 2, 3]);

            Assert.True(BlobStorePaths.TryGetPartialFolder(this.Blobs, submissionId, out string folder));
            return folder;
        }

        private string In(string root, string relative) =>
            Path.Combine(root, relative.Replace('/', Path.DirectorySeparatorChar));

        // ###########################################################################################
        // Both trees as a published board leaves them (CRT 4.0, 2026-10-25): the main file listing
        // the board that stays and (when `listed`) this one, both board files - this one listing its
        // own image and a manual the board that stays lists too, so its bytes are stored once for
        // both.
        // ###########################################################################################
        private void BothTrees(bool listed = true)
        {
            foreach (string root in new[] { this.thisBeta, this.thisProduction })
            {
                var rows = new List<MasterListingRow> { new("C64", "250407", DataTreeBuilder.BoardFile, string.Empty) };

                if (listed)
                    rows.Add(new MasterListingRow("C64", "999999", BoardDeletionFlowTests.BoardFile, string.Empty));

                DataTreeBuilder.ListingMaster(root, [.. rows]);
                DataTreeBuilder.Board(root, DataTreeBuilder.BoardId, (BoardDeletionFlowTests.SharedName, BoardDeletionFlowTests.SharedContent));
                this.ThisBoard(root);
            }
        }

        // This board's file in `root`, listing its image, the shared manual and `more`.
        private void ThisBoard(string root, params (string Name, string Content)[] more) =>
            DataTreeBuilder.Board(
                root,
                BoardDeletionFlowTests.BoardId,
                [
                    (BoardDeletionFlowTests.ImageName, BoardDeletionFlowTests.ImageContent),
                    (BoardDeletionFlowTests.SharedName, BoardDeletionFlowTests.SharedContent),
                    .. more
                ]);

        // A submission of `boardId`, then set to `state` at `decided` (left as sent when null).
        private static async Task<long> SubmissionAsync(
            FakeSubmissionStore store,
            string? state,
            string email = "contributor@example.com",
            string boardId = BoardDeletionFlowTests.BoardId,
            long? accountId = null,
            DateTimeOffset? decided = null)
        {
            string[] parts = boardId.Split('/');

            long id = await store.CreateAsync(
                new NewSubmission(
                    boardId, parts[0], parts[1], parts[2],
                    accountId, email, "192.0.2.1", "hash", "r0", "A change.", 1, [],
                    BoardDeletionFlowTests.Now.AddDays(-3), BoardDeletionFlowTests.Now.AddDays(-2)),
                CancellationToken.None);

            if (state is not null)
                await store.SetStateAsync(id, state, decided ?? BoardDeletionFlowTests.Now.AddDays(-1), CancellationToken.None);

            return id;
        }

        private async Task<BoardDeletionOutcome> PlanAsync(Fixture fixture, ReviewAccess? access = null, ServerOptions? options = null) =>
            await fixture.Flow.PlanAsync(
                access ?? BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, options ?? this.Options(), BoardDeletionFlowTests.Now);

        // Plans, then deletes what the plan showed - the Account screen's two calls.
        private async Task<BoardDeletionOutcome> PlanAndDeleteAsync(
            Fixture fixture,
            string? reason = "It was a test board.",
            Func<string, bool>? canWriteFolder = null)
        {
            BoardDeletionOutcome planned = await this.PlanAsync(fixture);
            Assert.True(planned.IsPlanned, planned.Error);

            return await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(),
                BoardDeletionFlowTests.BoardId,
                planned.Plan!.Fingerprint,
                reason,
                this.Options(),
                BoardDeletionFlowTests.Now,
                canWriteFolder: canWriteFolder ?? (_ => true));
        }

        // -----------------------------------------------------------------------------------
        // What goes
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE CASE THAT PROMPTED IT: "Not published yet", with nothing behind it. *** A record
        // and its submissions, in neither tree. All of it goes, and nothing of the other board.
        // ###########################################################################################
        [Fact]
        public async Task A_board_only_in_the_database_is_deleted_with_every_submission_of_it()
        {
            Fixture fixture = this.Make();
            long rejected = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);
            long abandoned = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Abandoned);
            long other = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, boardId: BoardDeletionFlowTests.OtherBoard);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Null(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.False(fixture.Store.Submissions.ContainsKey(rejected));
            Assert.False(fixture.Store.Submissions.ContainsKey(abandoned));

            Assert.True(fixture.Store.Submissions.ContainsKey(other));
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.OtherBoard));
        }

        // ###########################################################################################
        // *** "EVERYWHERE": both trees' folders and both drop-down lists. *** The board that stays
        // keeps its folder, its files and its row in each list.
        // ###########################################################################################
        [Fact]
        public async Task Its_folder_leaves_both_trees_and_its_row_leaves_both_drop_down_lists()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(2, outcome.BetaFilesRemoved);
            Assert.Equal(2, outcome.ProductionFilesRemoved);

            foreach (string root in new[] { this.thisBeta, this.thisProduction })
            {
                Assert.False(Directory.Exists(this.In(root, BoardDeletionFlowTests.BoardId)));
                Assert.False(File.Exists(this.In(root, BoardDeletionFlowTests.Image)));
                Assert.True(File.Exists(this.In(root, DataTreeBuilder.BoardFile)));

                MasterListingRow row = Assert.Single(DataTreeBuilder.ListedIn(root));
                Assert.Equal(BoardDeletionFlowTests.OtherBoard, row.BoardId);
            }
        }

        // ###########################################################################################
        // *** BYTES ANOTHER BOARD LISTS STAY (CRT 4.0, 2026-10-25). *** A file's bytes are stored
        // once for every board naming them: the manual the board that stays lists too is still its
        // manual afterwards, in both trees.
        // ###########################################################################################
        [Fact]
        public async Task Stored_bytes_another_board_also_lists_stay_in_both_trees()
        {
            this.BothTrees();
            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.True(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.Shared)));
            Assert.True(File.Exists(this.In(this.thisProduction, BoardDeletionFlowTests.Shared)));
        }

        // ###########################################################################################
        // Each tree is asked on its own: bytes the board that stays lists in the STABLE tree only
        // stay there, and go from BETA, where nothing else lists them.
        // ###########################################################################################
        [Fact]
        public async Task Bytes_another_board_lists_in_one_tree_stay_there_and_go_from_the_other()
        {
            this.BothTrees();
            DataTreeBuilder.Board(
                this.thisProduction,
                DataTreeBuilder.BoardId,
                (BoardDeletionFlowTests.SharedName, BoardDeletionFlowTests.SharedContent),
                ("sheet.png", BoardDeletionFlowTests.ImageContent));

            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(2, outcome.BetaFilesRemoved);
            Assert.Equal(1, outcome.ProductionFilesRemoved);
            Assert.False(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.Image)));
            Assert.True(File.Exists(this.In(this.thisProduction, BoardDeletionFlowTests.Image)));
        }

        // The hardware folder stays while another board is in it, and the manufacturer's with it.
        [Fact]
        public async Task An_emptied_folder_goes_but_one_still_holding_another_board_stays()
        {
            this.BothTrees();
            File.WriteAllText(DataTreeBuilder.Full(this.thisBeta, "Commodore/C64/999999/Docs/Old/notes.txt"), "notes");
            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.False(Directory.Exists(this.In(this.thisBeta, "Commodore/C64/999999")));
            Assert.True(Directory.Exists(this.In(this.thisBeta, "Commodore/C64")));
            Assert.True(Directory.Exists(this.In(this.thisBeta, "Commodore")));
        }

        // ###########################################################################################
        // *** BOTH MANIFESTS ARE REBUILT AT ONCE. *** Every other writer of a tree does it: a manifest
        // still listing the deleted files would have every CRT try to download them, and 404.
        // ###########################################################################################
        [Fact]
        public async Task Both_checksum_manifests_stop_listing_its_files()
        {
            this.BothTrees();
            Assert.True(DataChecksumManifest.Write(this.thisBeta, "https://example.com/app-data-BETA/Data", this.BetaManifest) > 0);
            Assert.True(DataChecksumManifest.Write(this.thisProduction, "https://example.com/app-data/Data", this.ProductionManifest) > 0);
            Assert.Contains("999999", File.ReadAllText(this.BetaManifest), StringComparison.Ordinal);

            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.DoesNotContain("999999", File.ReadAllText(this.BetaManifest), StringComparison.Ordinal);
            Assert.DoesNotContain("999999", File.ReadAllText(this.ProductionManifest), StringComparison.Ordinal);
            Assert.Contains("250407", File.ReadAllText(this.ProductionManifest), StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The stable row is taken out first, so when BETA's then fails, the stable main Excel data
        // file has already been rewritten - and a manifest still naming its old checksum makes every
        // CRT that downloads it throw it away (code review, 2026-10-04). The manifests follow on
        // that way out too.
        // ###########################################################################################
        [Fact]
        public async Task When_BETAs_row_cannot_be_removed_after_the_stable_one_the_manifests_still_follow()
        {
            this.BothTrees();
            Assert.True(DataChecksumManifest.Write(this.thisProduction, "https://example.com/app-data/Data", this.ProductionManifest) > 0);

            Fixture fixture = this.Make();
            BoardDeletionOutcome planned = await this.PlanAsync(fixture);
            Assert.True(planned.IsPlanned, planned.Error);

            BoardDeletionOutcome outcome = await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(),
                BoardDeletionFlowTests.BoardId,
                planned.Plan!.Fingerprint,
                "It was a test board.",
                this.Options(),
                BoardDeletionFlowTests.Now,
                canWriteFolder: _ => true,
                removeListing: survey => survey.TreeName == BoardDeletionRules.BetaTreeName
                    ? MasterListingEdit.Failed("locked by another program")
                    : BoardDeletionFiles.RemoveListing(survey, BoardDeletionFlowTests.BoardId));

            Assert.False(outcome.IsDone);
            Assert.Equal(BoardDeletionFlowTests.OtherBoard, Assert.Single(DataTreeBuilder.ListedIn(this.thisProduction)).BoardId);

            DataChecksumManifest.Entry master = Assert.Single(
                DataChecksumManifest.Scan(this.thisProduction, "https://example.com/app-data/Data"),
                entry => string.Equals(Path.GetFileName(entry.File), DataTreeLayout.MainFileName, StringComparison.Ordinal));
            Assert.Contains(master.Checksum, File.ReadAllText(this.ProductionManifest), StringComparison.Ordinal);
        }

        // The board views go with it (owner decision, 2026-10-03: "Delete them"); another board's stay.
        [Fact]
        public async Task Its_board_views_are_deleted_and_another_boards_are_not()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);
            fixture.Views.Rows.Add(FakeBoardViewStore.Row(BoardDeletionFlowTests.BoardId, BoardDeletionFlowTests.Now.AddDays(-1)));
            fixture.Views.Rows.Add(FakeBoardViewStore.Row(BoardDeletionFlowTests.OtherBoard, BoardDeletionFlowTests.Now.AddDays(-1)));

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(BoardDeletionFlowTests.OtherBoard, Assert.Single(fixture.Views.Rows).BoardId);
        }

        // A view count that cannot be deleted is statistics - the delete still completes.
        [Fact]
        public async Task A_board_view_store_that_fails_does_not_fail_the_delete()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);
            fixture.Views.FailDeletes = true;

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Null(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
        }

        // Audited under the board id, with who and why - the audit table is never deleted from.
        [Fact]
        public async Task The_delete_is_audited_under_the_board_id_with_the_reason()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture, "A test board.");

            Assert.True(outcome.IsDone, outcome.Error);
            AuditEntry entry = Assert.Single(fixture.Accounts.Audit, row => row.Action == BoardHistoryEvents.Deleted);
            Assert.Equal(BoardDeletionFlowTests.BoardId, entry.Subject);
            Assert.Equal(1, entry.ActorAccountId);
            Assert.Contains("A test board.", entry.Detail, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** ITS SUBMISSIONS' PARTIAL UPLOADS GO TOO (code review, 2026-10-04). *** The abandoned-
        // upload sweeper finds partials through the database rows, which the record's cascade has
        // just deleted - so a half-finished upload left here would stay on the disk for good.
        // Another board's upload is left alone.
        // ###########################################################################################
        [Fact]
        public async Task The_partial_uploads_of_its_submissions_are_cleared_and_another_boards_stay()
        {
            Fixture fixture = this.Make();
            long uploading = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Uploading);
            long other = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Uploading, boardId: BoardDeletionFlowTests.OtherBoard);

            string partial = this.PartialUpload(uploading);
            string otherPartial = this.PartialUpload(other);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.False(Directory.Exists(partial));
            Assert.True(Directory.Exists(otherPartial));
        }

        // ###########################################################################################
        // *** ONCE THE RECORD IS GONE, NOTHING MAY STOP THE MAILS (code review, 2026-10-04). *** An
        // audit write that fails after it - a database blip - answered 500, and a retry found "There
        // is no board": the contributors of the open submissions were never told. The delete is
        // done, so it is answered as done and they are mailed.
        // ###########################################################################################
        [Fact]
        public async Task An_audit_that_fails_after_the_record_is_deleted_still_mails_the_contributors()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, "waiting@example.com");
            fixture.Accounts.FailAuditWrites = true;

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Null(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.Equal("waiting@example.com", Assert.Single(fixture.Mail.Sent).ToAddress);
        }

        // ###########################################################################################
        // A request given up once the record is deleted - the long manifest rebuild before it is
        // where CRT's wait limit runs out - still finishes the views, the audit and the mails.
        // ###########################################################################################
        [Fact]
        public async Task A_request_cancelled_once_the_record_is_deleted_still_mails_the_contributors()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, "waiting@example.com");

            BoardDeletionOutcome planned = await this.PlanAsync(fixture);
            Assert.True(planned.IsPlanned, planned.Error);

            using var cancellation = new CancellationTokenSource();
            fixture.Store.AfterBoardDeleted = cancellation.Cancel;

            BoardDeletionOutcome outcome = await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(),
                BoardDeletionFlowTests.BoardId,
                planned.Plan!.Fingerprint,
                "It was a test board.",
                this.Options(),
                BoardDeletionFlowTests.Now,
                cancellation.Token,
                canWriteFolder: _ => true);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Single(fixture.Accounts.Audit, row => row.Action == BoardHistoryEvents.Deleted);
            Assert.Single(fixture.Mail.Sent);
        }

        // -----------------------------------------------------------------------------------
        // The contributors
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** OPEN SUBMISSIONS ARE DELETED AND THEIR CONTRIBUTORS MAILED (owner decision,
        // 2026-10-03). *** Open: waiting for review, approved once, in BETA and not yet in the stable
        // source. Finished ones - rejected, or merged and since promoted ("published") - are deleted
        // without a mail: there is nothing pending to tell anybody about.
        // ###########################################################################################
        [Fact]
        public async Task Only_the_contributors_of_open_submissions_are_mailed_with_the_reason()
        {
            Fixture fixture = this.Make();
            DateTimeOffset promoted = BoardDeletionFlowTests.Now.AddDays(-10);

            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, "waiting@example.com");
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Approved, "approved@example.com");
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Merged, "inbeta@example.com", decided: BoardDeletionFlowTests.Now.AddDays(-1));
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Merged, "published@example.com", decided: promoted.AddDays(-1));
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected, "rejected@example.com");

            fixture.Store.ProductionBoards[BoardDeletionFlowTests.BoardId] = new PublishedBoardRow("2026-September-23", "hash", promoted);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture, "It was a test board.");

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(3, outcome.ContributorsMailed);

            Assert.Equal(
                ["approved@example.com", "inbeta@example.com", "waiting@example.com"],
                fixture.Mail.Sent.Select(message => message.ToAddress).Order(StringComparer.Ordinal));

            Assert.All(fixture.Mail.Sent, message => Assert.Contains("It was a test board.", message.Body, StringComparison.Ordinal));
        }

        // A contributor who sent while signed in is mailed at the ACCOUNT's address, by name.
        [Fact]
        public async Task A_signed_in_contributor_is_mailed_at_the_accounts_address()
        {
            Fixture fixture = this.Make();
            fixture.Accounts.Accounts[7] = BoardDeletionFlowTests.Account(7, administrator: false, email: "anna@example.com") with { DisplayName = "Anna" };
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, email: string.Empty, accountId: 7);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            EmailMessage message = Assert.Single(fixture.Mail.Sent);
            Assert.Equal("anna@example.com", message.ToAddress);
            Assert.StartsWith("Hi Anna,", message.Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** THE REASON IS REQUIRED WHILE A SUBMISSION IS OPEN *** - it is the mail's only content,
        // the BETA rollback's rule. Refused before anything moves. Without open submissions nobody
        // is told anything, so none is needed.
        // ###########################################################################################
        [Fact]
        public async Task A_reason_is_required_when_a_submission_is_open_and_nothing_moves_without_one()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture, reason: "   ");

            Assert.False(outcome.IsDone);
            Assert.Equal(BoardDeletionRules.ReasonRequiredMessage, outcome.Error);
            Assert.True(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.BoardFile)));
            Assert.Equal(2, DataTreeBuilder.ListedIn(this.thisProduction).Count);
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.Empty(fixture.Mail.Sent);
        }

        [Fact]
        public async Task No_reason_is_needed_when_nothing_is_open()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture, reason: null);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Empty(fixture.Mail.Sent);
        }

        // -----------------------------------------------------------------------------------
        // The plan
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // What the confirmation shows is counted by the code that then deletes: files per tree, the
        // rows, the record, every submission, the open ones with the address that is mailed, the
        // pool and the open invitations.
        // ###########################################################################################
        [Fact]
        public async Task The_plan_counts_everything_that_would_go()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            long open = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending, "waiting@example.com");
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            fixture.Accounts.Accounts[5] = BoardDeletionFlowTests.Account(5, administrator: false);
            fixture.Accounts.Maintainers.Add((BoardDeletionFlowTests.BoardId, 5));

            BoardDeletionOutcome outcome = await this.PlanAsync(fixture);

            Assert.True(outcome.IsPlanned, outcome.Error);
            BoardDeletePlanAnswer plan = outcome.Plan!.ToAnswer();

            Assert.Equal(("Commodore", "C64", "999999"), (plan.Manufacturer, plan.Hardware, plan.Board));
            Assert.Equal(2, plan.BetaFiles);
            Assert.Equal(2, plan.ProductionFiles);
            Assert.True(plan.ListedInBeta);
            Assert.True(plan.ListedInProduction);
            Assert.True(plan.HasRecord);
            Assert.Equal(2, plan.Submissions);
            Assert.Equal(1, plan.Maintainers);
            Assert.Null(plan.BlockedBecause);
            Assert.False(string.IsNullOrWhiteSpace(plan.Fingerprint));

            BoardDeleteOpenSubmission listed = Assert.Single(plan.OpenSubmissions);
            Assert.Equal((open, SubmissionState.Pending, "waiting@example.com"), (listed.Id, listed.State, listed.Contributor));
        }

        // A plan writes nothing at all.
        [Fact]
        public async Task Planning_changes_nothing()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending);

            await this.PlanAsync(fixture);

            Assert.True(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.Image)));
            Assert.Equal(2, DataTreeBuilder.ListedIn(this.thisBeta).Count);
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.Empty(fixture.Accounts.Audit);
        }

        // ###########################################################################################
        // *** WHAT WAS CONFIRMED IS WHAT IS DELETED. *** A file copied in, a publish, or a new
        // submission between the plan and the delete changes the fingerprint, and the delete is
        // refused with nothing moved - the administrator never deletes what they were not shown.
        // ###########################################################################################
        [Fact]
        public async Task A_board_that_changed_since_the_plan_is_not_deleted()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            BoardDeletionOutcome planned = await this.PlanAsync(fixture);

            // A publish since: the board lists one more file.
            this.ThisBoard(this.thisBeta, ("Images/new.png", "a new image"));

            BoardDeletionOutcome outcome = await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, planned.Plan!.Fingerprint, "Test.",
                this.Options(), BoardDeletionFlowTests.Now, canWriteFolder: _ => true);

            Assert.True(outcome.IsConflict);
            Assert.Equal(BoardDeletionRules.ChangedMessage, outcome.Error);
            Assert.True(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.BoardFile)));
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
        }

        [Fact]
        public async Task A_submission_sent_since_the_plan_stops_the_delete()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            BoardDeletionOutcome planned = await this.PlanAsync(fixture);

            long arrived = await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending);

            BoardDeletionOutcome outcome = await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, planned.Plan!.Fingerprint, "Test.",
                this.Options(), BoardDeletionFlowTests.Now);

            Assert.True(outcome.IsConflict);
            Assert.True(fixture.Store.Submissions.ContainsKey(arrived));
        }

        // -----------------------------------------------------------------------------------
        // What it refuses
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // A main file that cannot be read says nothing about whether it lists the board: refused,
        // with nothing moved, never a guess.
        // ###########################################################################################
        [Fact]
        public async Task A_main_file_that_cannot_be_read_stops_the_delete()
        {
            this.BothTrees();
            File.WriteAllText(this.In(this.thisProduction, DataTreeLayout.MainFileName), "{ not json");

            Fixture fixture = this.Make();

            BoardDeletionOutcome planned = await this.PlanAsync(fixture);

            Assert.True(planned.IsPlanned, planned.Error);
            Assert.StartsWith($"[{DataTreeLayout.MainFileName}] in the stable data could not be read", planned.Plan!.BlockedBecause, StringComparison.Ordinal);

            BoardDeletionOutcome outcome = await fixture.Flow.DeleteAsync(
                BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, planned.Plan.Fingerprint, "Test.",
                this.Options(), BoardDeletionFlowTests.Now, canWriteFolder: _ => true);

            Assert.True(outcome.IsConflict);
            Assert.True(File.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.BoardFile)));
            Assert.Equal(2, DataTreeBuilder.ListedIn(this.thisBeta).Count);
        }

        // ###########################################################################################
        // Its KiCad data and anything else in its folder are the board's, and go with it: the board
        // file, the stored bytes only it lists (the image and the KiCad file) and the folder's note.
        // ###########################################################################################
        [Fact]
        public async Task Its_own_KiCad_data_and_whatever_its_folder_holds_go_with_it()
        {
            this.BothTrees();
            this.ThisBoard(this.thisBeta, ("KiCad data/board.kicad_pcb", "(kicad_pcb)"));
            File.WriteAllText(DataTreeBuilder.Full(this.thisBeta, "Commodore/C64/999999/!README.txt"), "read me");

            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(4, outcome.BetaFilesRemoved);
            Assert.False(Directory.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.BoardId)));
            Assert.False(File.Exists(this.In(this.thisBeta, DataTreeBuilder.Stored("(kicad_pcb)", "KiCad data/board.kicad_pcb"))));
        }

        // A symbolic link on the way to a stored file it lists: deleting through it could reach
        // outside the tree.
        [Fact]
        public async Task A_link_on_the_way_to_a_stored_file_stops_the_delete()
        {
            this.BothTrees();
            string link = this.In(this.thisBeta, BoardDeletionFlowTests.Image);

            Fixture fixture = this.Make();

            BoardDeletionOutcome planned = await fixture.Flow.PlanAsync(
                BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, this.Options(), BoardDeletionFlowTests.Now,
                isLink: path => string.Equals(path, link, StringComparison.Ordinal));

            Assert.Equal(
                BoardDeletionRules.LinkMessage(BoardDeletionRules.BetaTreeName, this.thisBeta, link),
                planned.Plan!.BlockedBecause);
        }

        // A symbolic link anywhere in the folder: deleting through it could reach outside the tree.
        [Fact]
        public async Task A_link_in_its_folder_stops_the_delete()
        {
            this.BothTrees();
            File.WriteAllText(DataTreeBuilder.Full(this.thisBeta, "Commodore/C64/999999/Images/notes.txt"), "notes");
            string link = this.In(this.thisBeta, "Commodore/C64/999999/Images");

            Fixture fixture = this.Make();

            BoardDeletionOutcome planned = await fixture.Flow.PlanAsync(
                BoardDeletionFlowTests.Admin(), BoardDeletionFlowTests.BoardId, this.Options(), BoardDeletionFlowTests.Now,
                isLink: path => string.Equals(path, link, StringComparison.Ordinal));

            Assert.Equal(
                BoardDeletionRules.LinkMessage(BoardDeletionRules.BetaTreeName, this.thisBeta, link),
                planned.Plan!.BlockedBecause);

            // Named from the data's root - never the server's own absolute path (code review,
            // 2026-10-13: the administrator's screen is no place for the server's folders).
            Assert.Contains("[Commodore/C64/999999/Images]", planned.Plan.BlockedBecause);
            Assert.DoesNotContain(this.thisBeta, planned.Plan.BlockedBecause);
        }

        // ###########################################################################################
        // *** EVERY FOLDER IS PROBED BEFORE ANYTHING MOVES (TreeWriteAccess). *** A folder copied into
        // the stable data by hand as root refuses the delete with the BETA side untouched too - not
        // half a board gone.
        // ###########################################################################################
        [Fact]
        public async Task A_folder_the_service_may_not_write_refuses_before_anything_changes()
        {
            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            // The folder the image's bytes are stored in ("Files/<ab>").
            string storedFolder = BoardDeletionFlowTests.Image[..BoardDeletionFlowTests.Image.LastIndexOf('/')];
            string refusing = this.In(this.thisProduction, storedFolder);

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(
                fixture, canWriteFolder: folder => !string.Equals(folder, refusing, StringComparison.Ordinal));

            Assert.False(outcome.IsDone);
            Assert.Contains(storedFolder, outcome.Error, StringComparison.Ordinal);

            foreach (string root in new[] { this.thisBeta, this.thisProduction })
            {
                Assert.True(File.Exists(this.In(root, BoardDeletionFlowTests.Image)));
                Assert.Equal(2, DataTreeBuilder.ListedIn(root).Count);
            }

            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
        }

        // ###########################################################################################
        // *** A FILE THAT WILL NOT GO KEEPS THE RECORD. *** The record, the submissions and the mails
        // are the LAST step, so pressing Delete again finishes the job. Windows only: a file held
        // open cannot be deleted there, and is everywhere else.
        // ###########################################################################################
        [Fact]
        public async Task A_file_that_will_not_go_keeps_the_record_and_a_second_delete_finishes()
        {
            Assert.SkipUnless(OperatingSystem.IsWindows(), "Only Windows refuses to delete a file held open.");

            this.BothTrees();
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Pending);

            BoardDeletionOutcome first;

            using (new FileStream(this.In(this.thisBeta, BoardDeletionFlowTests.Image), FileMode.Open, FileAccess.Read, FileShare.None))
                first = await this.PlanAndDeleteAsync(fixture);

            Assert.False(first.IsDone);
            Assert.Equal(BoardDeletionRules.FilesLeftMessage(1), first.Error);
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.Empty(fixture.Mail.Sent);

            // Out of both lists already; the rest goes on the second press.
            Assert.Single(DataTreeBuilder.ListedIn(this.thisBeta));

            BoardDeletionOutcome second = await this.PlanAndDeleteAsync(fixture);

            Assert.True(second.IsDone, second.Error);
            Assert.False(Directory.Exists(this.In(this.thisBeta, BoardDeletionFlowTests.BoardId)));
            Assert.Null(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
            Assert.Single(fixture.Mail.Sent);
        }

        [Fact]
        public async Task Only_an_administrator_may_plan_or_delete()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            // A maintainer of this very board is still not an administrator.
            ReviewAccess maintainer = ReviewAccess.For(
                BoardDeletionFlowTests.Account(2, administrator: false), [BoardDeletionFlowTests.BoardId]);

            BoardDeletionOutcome planned = await this.PlanAsync(fixture, maintainer);
            BoardDeletionOutcome deleted = await fixture.Flow.DeleteAsync(
                maintainer, BoardDeletionFlowTests.BoardId, "x", "Test.", this.Options(), BoardDeletionFlowTests.Now);

            Assert.True(planned.IsForbidden);
            Assert.True(deleted.IsForbidden);
            Assert.NotNull(await fixture.Store.FindBoardAsync(BoardDeletionFlowTests.BoardId));
        }

        // "Everywhere" includes the stable data, which the service may only write once it is set up.
        [Fact]
        public async Task Nothing_happens_while_publishing_to_the_stable_source_is_off()
        {
            Fixture fixture = this.Make();
            await BoardDeletionFlowTests.SubmissionAsync(fixture.Store, SubmissionState.Rejected);

            BoardDeletionOutcome outcome = await this.PlanAsync(fixture, options: this.Options(production: false));

            Assert.True(outcome.IsNotConfigured);
            Assert.Equal(ProductionPromotionFlow.NotConfiguredMessage, outcome.Error);
        }

        // Nothing of it anywhere - no record, no folder, no row - is not a board.
        [Fact]
        public async Task A_board_nothing_knows_is_not_found()
        {
            this.BothTrees(listed: false);
            Directory.Delete(this.In(this.thisBeta, BoardDeletionFlowTests.BoardId), recursive: true);
            Directory.Delete(this.In(this.thisProduction, BoardDeletionFlowTests.BoardId), recursive: true);

            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAsync(fixture);

            Assert.True(outcome.IsNotFound);
            Assert.Equal(BoardDeletionRules.NoSuchBoardMessage(BoardDeletionFlowTests.BoardId), outcome.Error);
        }

        // ###########################################################################################
        // A board only in the trees - shipped, never submitted to, so no record - is still a board,
        // and goes like one.
        // ###########################################################################################
        [Fact]
        public async Task A_board_with_no_database_record_is_deleted_from_both_trees()
        {
            this.BothTrees();
            Fixture fixture = this.Make();

            BoardDeletionOutcome outcome = await this.PlanAndDeleteAsync(fixture, reason: null);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.False(outcome.Plan!.HasRecord);
            Assert.False(Directory.Exists(this.In(this.thisProduction, BoardDeletionFlowTests.BoardId)));
        }
    }
}
