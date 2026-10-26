using System.Security.Cryptography;
using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers rolling a BETA board back to what production holds, and returning its submissions to
    // the queue (owner decision, 2026-09-27) - the production window's "push back to queue".
    //
    // Asserted on the DISK, like ProductionPromotionFlowTests: the point of a rollback is what the
    // BETA tree holds afterwards, and the only honest way to test it is to look. The trees are CRT
    // 4.0 ones (2026-10-25): a board is its board file and the stored files it lists, so "BETA
    // holds production's sheet" is read through BETA's board file (BetaHolds).
    //
    // *** THE TWO FACTS THAT MAKE IT POSSIBLE *** are pinned elsewhere and named here so a reader
    // knows where to look: a MERGED submission keeps its blobs (SubmissionCollectionStates.Live
    // contains Merged - SubmissionCollectionStatesTests), and production holds a COMPLETE board
    // (ProductionPromotionPlan). It was first judged impossible on the belief that the old bytes
    // were collected, which was wrong.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class BetaRollbackFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string BoardId = "Commodore/C64/250407";
        private const string OtherBoardId = "Commodore/C128/310378";
        private const string Sheet = "Images/sheet1.png";
        private const string NewImage = "Images/new.png";

        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisProduction;
        private readonly string thisBlobRoot;

        public BetaRollbackFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-rollback", Guid.NewGuid().ToString("N"));
            this.thisBeta = Path.Combine(this.thisRoot, "app-data-BETA", "Data");
            this.thisProduction = Path.Combine(this.thisRoot, "app-data", "Data");
            this.thisBlobRoot = Path.Combine(this.thisRoot, "blobs");

            Directory.CreateDirectory(this.thisBeta);
            Directory.CreateDirectory(this.thisProduction);
            Directory.CreateDirectory(this.thisBlobRoot);

            DataTreeBuilder.Master(this.thisBeta, BetaRollbackFlowTests.BoardId);
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

        private ServerOptions Options() => new()
        {
            DataTreeRoot = this.thisBeta,
            ManifestPath = this.BetaManifest,
            PublicDataBaseUrl = "https://example.com/app-data-BETA/Data",
            ProductionTreeRoot = Path.Combine(this.thisRoot, "app-data"),
            ProductionDataTreeRoot = this.thisProduction,
            ProductionManifestPath = Path.Combine(this.thisRoot, "app-data", "dataChecksums.json"),
            ProductionPublicDataBaseUrl = "https://example.com/app-data/Data"
        };

        private string BetaManifest => Path.Combine(this.thisRoot, "app-data-BETA", "dataChecksums.json");

        private static string Sha(string content) =>
            Convert.ToHexStringLower(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(content)));

        private static AccountRecord Account(long id = 1, bool administrator = true) =>
            new(id, $"a{id}@example.com", $"a{id}@example.com", "hash", $"A{id}",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false,
                BetaRollbackFlowTests.Now, null);

        private static ReviewAccess Admin() => ReviewAccess.For(BetaRollbackFlowTests.Account());

        private BetaRollbackFlow Flow(FakeSubmissionStore store, FakeAccountStore? accounts = null) =>
            new(store, accounts ?? new FakeAccountStore(), new PublishLock(), NullLogger<BetaRollbackFlow>.Instance);

        private string InBeta(string relative) =>
            Path.Combine(this.thisBeta, relative.Replace('/', Path.DirectorySeparatorChar));

        // Each tree's copy of the board: its board file listing these files, their bytes stored.
        private void Production(params (string Name, string Content)[] files) =>
            DataTreeBuilder.Board(this.thisProduction, BetaRollbackFlowTests.BoardId, files);

        private void Beta(params (string Name, string Content)[] files) =>
            DataTreeBuilder.Board(this.thisBeta, BetaRollbackFlowTests.BoardId, files);

        // The usual case: production holds the published sheet; BETA the submitted one and, when
        // `withNewImage`, an image the submission added.
        private void BetaAheadOfProduction(bool withNewImage = false)
        {
            this.Production((BetaRollbackFlowTests.Sheet, "PUBLISHED"));

            if (withNewImage)
                this.Beta((BetaRollbackFlowTests.Sheet, "SUBMITTED"), (BetaRollbackFlowTests.NewImage, "ONLY IN BETA"));
            else
                this.Beta((BetaRollbackFlowTests.Sheet, "SUBMITTED"));
        }

        // What BETA's board file says `name` holds, read through its stored bytes - null when it
        // lists no such file, or its bytes are not stored.
        private string? BetaHolds(string name)
        {
            BoardDocument? document = BoardFiles.ReadDocument(this.InBeta(DataTreeLayout.BoardFilePath(BetaRollbackFlowTests.BoardId)));
            BoardDocument.StoredFileRow? row = document?.Files.FirstOrDefault(file => string.Equals(file.Name, name, StringComparison.Ordinal));

            if (row is null)
                return null;

            string stored = this.InBeta(DataTreeLayout.StoredFilePath(row.Sha256, row.Name));

            return File.Exists(stored) ? File.ReadAllText(stored) : null;
        }

        private bool StoredInBeta(string content, string name) =>
            File.Exists(this.InBeta(DataTreeBuilder.Stored(content, name)));

        // ###########################################################################################
        // A board whose BETA state is ahead of production, with one merged submission behind it.
        // The trees are written directly rather than published through ApprovePublishFlow: this
        // class is about what a rollback WRITES, and a real publish would only add noise.
        // ###########################################################################################
        private FakeSubmissionStore StoreAheadOfProduction(bool everPromoted = true)
        {
            var store = new FakeSubmissionStore();

            store.Boards[BetaRollbackFlowTests.BoardId] = new NewSubmission(
                BetaRollbackFlowTests.BoardId, "Commodore", "C64", "250407",
                null, "contributor@example.com", "192.0.2.1", "hash", "r0", "A change.", 1, [],
                BetaRollbackFlowTests.Now, BetaRollbackFlowTests.Now.AddHours(24));

            // The fake derives BoardRecord from these two rows, as the real store derives it from
            // the `boards` columns - so the state is set the way a publish and a promotion set it.
            store.PublishedBoards[BetaRollbackFlowTests.BoardId] =
                new PublishedBoardRow("2026-September-27", "beta-hash", BetaRollbackFlowTests.Now.AddDays(-1));

            if (everPromoted)
            {
                store.ProductionBoards[BetaRollbackFlowTests.BoardId] =
                    new PublishedBoardRow("2026-May-14", "production-hash", BetaRollbackFlowTests.Now.AddDays(-30));
            }

            return store;
        }

        // `files`: what the submission carried - path and the hash of the bytes it put there.
        private async Task<long> MergedSubmissionAsync(
            FakeSubmissionStore store,
            string email = "contributor@example.com",
            params SubmissionFile[] files)
        {
            long id = await store.CreateAsync(
                new NewSubmission(
                    BetaRollbackFlowTests.BoardId, "Commodore", "C64", "250407",
                    null, email, "192.0.2.1", "hash", "r0", "Corrected U8.", 1, [.. files],
                    BetaRollbackFlowTests.Now.AddDays(-2), BetaRollbackFlowTests.Now.AddDays(-1)),
                CancellationToken.None);

            await store.SetStateAsync(id, SubmissionState.Merged, BetaRollbackFlowTests.Now.AddDays(-1), CancellationToken.None);

            return id;
        }

        // -----------------------------------------------------------------------------------
        // The restore
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** PRODUCTION'S BOARD GOES BACK, AND WHAT ONLY THE SUBMISSIONS ADDED LEAVES. *** Both
        // halves matter: restoring alone would leave the bytes the rolled-back submission ADDED
        // sitting in BETA, listed by nothing and still served to every BETA user.
        // ###########################################################################################
        [Fact]
        public async Task A_rollback_puts_productions_board_back_and_removes_what_only_the_submissions_added()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction(withNewImage: true);

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "The U8 pinout is wrong.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);

            Assert.Equal("PUBLISHED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.Null(this.BetaHolds(BetaRollbackFlowTests.NewImage));
            Assert.Equal(
                File.ReadAllBytes(Path.Combine(this.thisProduction, "Commodore", "C64", "250407", "board.json")),
                File.ReadAllBytes(this.InBeta(DataTreeLayout.BoardFilePath(BetaRollbackFlowTests.BoardId))));

            Assert.False(this.StoredInBeta("SUBMITTED", BetaRollbackFlowTests.Sheet));
            Assert.False(this.StoredInBeta("ONLY IN BETA", BetaRollbackFlowTests.NewImage));
        }

        // ###########################################################################################
        // *** EVERY SUBMISSION MERGED SINCE THE LAST PROMOTION GOES BACK TO THE QUEUE (owner
        // decision). *** A rollback is per BOARD - the board's file holds all of them - so picking
        // one out is impossible. All of them return to `pending` with the reason.
        // ###########################################################################################
        [Fact]
        public async Task Every_merged_submission_goes_back_to_the_queue_with_the_reason()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            long first = await this.MergedSubmissionAsync(store, "one@example.com");
            long second = await this.MergedSubmissionAsync(store, "two@example.com");
            this.BetaAheadOfProduction();

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Needs the revision date.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);

            foreach (long id in new[] { first, second })
            {
                SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

                Assert.Equal(SubmissionState.Pending, record!.State);
                Assert.Equal("Needs the revision date.", record.DecisionComment);

                // RECORDED as returned (migration 0015, code review 2026-09-29), at the instant of
                // its decision - what the contributor's "Taken back out of BETA" is now read from.
                Assert.Equal(BetaRollbackFlowTests.Now, store.BetaReturns[id]);
                Assert.Equal(
                    ProductionPromotionRules.ReturnedState,
                    ProductionPromotionRules.ContributorFacingState(record.State, record.DecidedUtc, null, store.BetaReturns[id]));
            }
        }

        // ###########################################################################################
        // *** "REJECT" ON BETA > PROD IS THE SAME ROLLBACK (owner request, 2026-09-28: "a direct
        // 'Reject' button also - just like the normal queue"). *** BETA moves exactly as for a
        // push-back; every merged submission is REJECTED with the reason instead of returning to the
        // queue, and the history names it as a rejection.
        // ###########################################################################################
        [Fact]
        public async Task A_rejection_rolls_BETA_back_the_same_and_rejects_every_merged_submission()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            long first = await this.MergedSubmissionAsync(store, "one@example.com");
            long second = await this.MergedSubmissionAsync(store, "two@example.com");
            var accounts = new FakeAccountStore();
            this.BetaAheadOfProduction(withNewImage: true);

            BetaRollbackOutcome outcome = await this.Flow(store, accounts).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Not what this board needs.",
                this.Options(), BetaRollbackFlowTests.Now, reject: true);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.True(outcome.Rejected);

            Assert.Equal("PUBLISHED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.False(this.StoredInBeta("ONLY IN BETA", BetaRollbackFlowTests.NewImage));
            Assert.Equal(("2026-May-14", "production-hash"), store.BetaStates[BetaRollbackFlowTests.BoardId]);

            foreach (long id in new[] { first, second })
            {
                SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

                Assert.Equal(SubmissionState.Rejected, record!.State);
                Assert.Equal("Not what this board needs.", record.DecisionComment);

                // A rejection is its own final state, never "returned".
                Assert.False(store.BetaReturns.ContainsKey(id));
            }

            AuditEntry audit = Assert.Single(accounts.Audit);
            Assert.Equal(BoardHistoryEvents.RejectedFromBeta, audit.Action);
            Assert.Contains("2 submission(s) rejected; Not what this board needs.", audit.Detail, StringComparison.Ordinal);
        }

        // A rejection without a reason is refused like a push-back, in its own words.
        [Fact]
        public async Task A_rejection_with_no_reason_is_refused_and_writes_nothing()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            long id = await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction();

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, " ",
                this.Options(), BetaRollbackFlowTests.Now, reject: true);

            Assert.False(outcome.IsDone);
            Assert.Contains("why this is being rejected", outcome.Error, StringComparison.Ordinal);
            Assert.Equal("SUBMITTED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // ###########################################################################################
        // *** THE RECORDED BETA STATE FOLLOWS THE TREE. *** Left alone, `content_hash` would still
        // name the state that was just removed - the production window would keep offering to
        // promote a board BETA no longer holds, and the next publish would compare against a hash
        // describing nothing.
        // ###########################################################################################
        [Fact]
        public async Task A_restored_board_is_recorded_as_level_with_production()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction();

            await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong data.",
                this.Options(), BetaRollbackFlowTests.Now);

            (string? revision, string? hash) = store.BetaStates[BetaRollbackFlowTests.BoardId];

            Assert.Equal("2026-May-14", revision);
            Assert.Equal("production-hash", hash);
        }

        // -----------------------------------------------------------------------------------
        // A board never promoted
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** NOTHING IN PRODUCTION MEANS THE BOARD LEAVES BETA (owner decision, 2026-09-27). ***
        // There is no state to restore to, and "restore nothing" would leave the bad board exactly
        // as it is while reporting success.
        // ###########################################################################################
        [Fact]
        public async Task A_board_never_promoted_is_removed_from_beta_entirely()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction(everPromoted: false);
            await this.MergedSubmissionAsync(store);
            this.Beta((BetaRollbackFlowTests.Sheet, "THE NEW BOARD"));

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Not ready.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal(BetaRollbackKind.RemoveFromBeta, outcome.Plan!.Kind);

            Assert.False(File.Exists(this.InBeta(DataTreeLayout.BoardFilePath(BetaRollbackFlowTests.BoardId))));
            Assert.False(this.StoredInBeta("THE NEW BOARD", BetaRollbackFlowTests.Sheet));

            // The emptied folder goes too, so the tree does not keep a husk of a board.
            Assert.False(Directory.Exists(this.InBeta("Commodore/C64/250407")));

            // And it has no BETA state at all any more.
            (string? revision, string? hash) = store.BetaStates[BetaRollbackFlowTests.BoardId];
            Assert.Null(revision);
            Assert.Null(hash);
        }

        // ###########################################################################################
        // *** A PROMOTED BOARD WHOSE PRODUCTION COPY IS GONE IS REFUSED, NOT REMOVED (code review,
        // 2026-09-27). *** An empty production listing used to mean "never promoted", so a mount
        // that was down or a folder renamed by hand wiped a board that IS in production out of
        // BETA. The record decides now: the plan is refused, and BETA is not touched.
        // ###########################################################################################
        [Fact]
        public async Task A_promoted_board_whose_production_folder_cannot_be_read_is_refused_and_beta_kept()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction(everPromoted: true);
            long id = await this.MergedSubmissionAsync(store);

            // BETA has the board; production holds no board file of it at all.
            this.Beta((BetaRollbackFlowTests.Sheet, "THE BOARD"));

            BetaRollbackOutcome planned = await this.Flow(store).PlanAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, this.Options());

            Assert.True(planned.IsConflict);
            Assert.Equal(BetaRollbackFlow.ProductionUnreadableMessage, planned.Error);

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Not ready.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsConflict);
            Assert.False(outcome.IsDone);
            Assert.Equal("THE BOARD", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // ###########################################################################################
        // *** THE BETA MANIFEST DESCRIBES THE ROLLED-BACK TREE (code review, 2026-09-27). *** The
        // first version rewrote the tree and left the manifest advertising the post-publish hashes:
        // a CRT holding the bad bytes saw them match and never re-downloaded, a fresh one refused
        // the restored file on its checksum, and the removed file 404'd on every sync.
        // ###########################################################################################
        [Fact]
        public async Task The_BETA_checksum_manifest_is_rebuilt_to_match_the_rolled_back_tree()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction(withNewImage: true);

            // What the publish left: a manifest naming the submitted bytes and the new image.
            DataChecksumManifest.Write(this.thisBeta, this.Options().PublicDataBaseUrl!, this.BetaManifest);
            Assert.Contains(BetaRollbackFlowTests.Sha("ONLY IN BETA"), File.ReadAllText(this.BetaManifest), StringComparison.OrdinalIgnoreCase);

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong data.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);

            string manifest = File.ReadAllText(this.BetaManifest);

            Assert.DoesNotContain(BetaRollbackFlowTests.Sha("ONLY IN BETA"), manifest, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(BetaRollbackFlowTests.Sha("PUBLISHED"), manifest, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(BetaRollbackFlowTests.Sha("SUBMITTED"), manifest, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // *** A RETURNED SUBMISSION STARTS ITS REVIEW AGAIN WITH NO APPROVALS (code review,
        // 2026-09-27). *** Left in place, a submission approved by BOTH roles was republished by ONE
        // approval, and an administrator who had approved an ordinary one was answered "you have
        // already approved" and could never approve it again.
        // ###########################################################################################
        [Fact]
        public async Task A_returned_submissions_earlier_approvals_are_cleared()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            long id = await this.MergedSubmissionAsync(store);

            await store.AddApprovalAsync(id, ApproverRole.Maintainer, 5, "M", BetaRollbackFlowTests.Now.AddDays(-2), CancellationToken.None);
            await store.AddApprovalAsync(id, ApproverRole.Administrator, 1, "A", BetaRollbackFlowTests.Now.AddDays(-1), CancellationToken.None);

            this.BetaAheadOfProduction();

            await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Needs another look.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.Empty(await store.GetApprovalsAsync(id, CancellationToken.None));
        }

        // ###########################################################################################
        // *** A FAILED RECORD IS SAID, AND PUSHING BACK AGAIN FINISHES IT (code review,
        // 2026-09-27). *** The bookkeeping is one transaction after the tree has moved. A failure
        // there used to escape as a 500 - "it did not happen" - with some submissions flipped and
        // never mailed. Now nothing is recorded, the maintainer is told the truth, and the second
        // attempt finds the tree already level, moves no file, and records it.
        // ###########################################################################################
        [Fact]
        public async Task A_rollback_whose_record_fails_says_so_and_a_second_push_back_finishes_it()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            long id = await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction();

            store.FailRollbackRecord = true;

            BetaRollbackOutcome first = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong data.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.False(first.IsDone);
            Assert.Equal(BetaRollbackFlow.NotRecordedMessage, first.Error);

            // The tree moved; the database did not, at all.
            Assert.Equal("PUBLISHED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);

            store.FailRollbackRecord = false;

            BetaRollbackOutcome second = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong data.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(second.IsDone, second.Error);
            Assert.Equal(0, second.FilesRestored);
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
            Assert.Equal(id, Assert.Single(second.Plan!.Returning).Id);
        }

        // -----------------------------------------------------------------------------------
        // Bytes stored once for several boards (CRT 4.0 tree, 2026-10-25)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** BYTES ANOTHER BETA BOARD LISTS STAY. *** In the 4.0 tree a file's bytes are stored once
        // for every board naming them, so what the rolled-back submission added may be another
        // board's file too. Only bytes no BETA board lists once production's board file is back go.
        // ###########################################################################################
        [Fact]
        public async Task Bytes_another_BETA_board_lists_stay_when_the_board_is_rolled_back()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction(withNewImage: true);

            DataTreeBuilder.Master(this.thisBeta, BetaRollbackFlowTests.BoardId, BetaRollbackFlowTests.OtherBoardId);
            DataTreeBuilder.Board(this.thisBeta, BetaRollbackFlowTests.OtherBoardId, ("Images/the same picture.png", "ONLY IN BETA"));

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong picture.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Null(this.BetaHolds(BetaRollbackFlowTests.NewImage));
            Assert.True(this.StoredInBeta("ONLY IN BETA", BetaRollbackFlowTests.NewImage));
            Assert.False(this.StoredInBeta("SUBMITTED", BetaRollbackFlowTests.Sheet));

            // Nothing in the 4.0 tree is a shared file.
            Assert.Empty(outcome.Plan!.SharedRestored);
            Assert.False(outcome.Plan.TouchesSharedFiles);
        }

        // ###########################################################################################
        // *** A TREE THAT CANNOT BE READ COMPLETELY KEEPS EVERY STORED FILE. *** Whether another board
        // lists the bytes is then not known, so none are removed - failing closed, as every writer
        // does. The board itself is still put back.
        // ###########################################################################################
        [Fact]
        public async Task A_BETA_tree_that_cannot_be_read_completely_keeps_the_stored_files_and_still_restores_the_board()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction(withNewImage: true);

            DataTreeBuilder.Master(this.thisBeta, BetaRollbackFlowTests.BoardId, BetaRollbackFlowTests.OtherBoardId);
            File.WriteAllText(DataTreeBuilder.Full(this.thisBeta, DataTreeLayout.BoardFilePath(BetaRollbackFlowTests.OtherBoardId)), "{ not a board file");

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Wrong data.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal("PUBLISHED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
            Assert.True(this.StoredInBeta("SUBMITTED", BetaRollbackFlowTests.Sheet));
            Assert.True(this.StoredInBeta("ONLY IN BETA", BetaRollbackFlowTests.NewImage));
        }

        // -----------------------------------------------------------------------------------
        // Refusals
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE REASON IS REQUIRED. *** It is the contributor's ONLY feedback - contributing needs
        // no account, so there is no inbox and no thread. Telling someone their accepted work left
        // BETA without saying why is worse than not offering the button.
        // ###########################################################################################
        [Fact]
        public async Task A_rollback_with_no_reason_is_refused_and_writes_nothing()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction();

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "   ",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.False(outcome.IsDone);
            Assert.Contains("Say why", outcome.Error, StringComparison.Ordinal);

            // BETA is untouched - the only honest test of a refusal.
            Assert.Equal("SUBMITTED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
        }

        // Someone who does not maintain this board cannot roll it back, exactly as they cannot
        // publish it.
        [Fact]
        public async Task An_account_without_authority_over_the_board_is_refused()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store);
            this.BetaAheadOfProduction();

            ReviewAccess outsider = ReviewAccess.For(BetaRollbackFlowTests.Account(2, administrator: false));

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                outsider, BetaRollbackFlowTests.BoardId, "Not mine.", this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsForbidden);
            Assert.Equal("SUBMITTED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
        }

        // A board already level with production has nothing to roll back - the same question
        // "is BETA ahead?" the production list is built on.
        [Fact]
        public async Task A_board_level_with_production_is_refused()
        {
            var store = new FakeSubmissionStore();

            store.Boards[BetaRollbackFlowTests.BoardId] = new NewSubmission(
                BetaRollbackFlowTests.BoardId, "Commodore", "C64", "250407",
                null, "c@example.com", "192.0.2.1", "hash", "r0", "A change.", 1, [],
                BetaRollbackFlowTests.Now, BetaRollbackFlowTests.Now.AddHours(24));

            store.PublishedBoards[BetaRollbackFlowTests.BoardId] =
                new PublishedBoardRow("2026-May-14", "same", BetaRollbackFlowTests.Now.AddDays(-1));

            store.ProductionBoards[BetaRollbackFlowTests.BoardId] =
                new PublishedBoardRow("2026-May-14", "same", BetaRollbackFlowTests.Now.AddDays(-1));

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Why not.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsConflict);
        }

        // -----------------------------------------------------------------------------------
        // The plan, shown before anything happens
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_plan_says_what_would_change_and_who_goes_back_without_touching_anything()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            await this.MergedSubmissionAsync(store, "hest@mailscan.dk");
            this.BetaAheadOfProduction(withNewImage: true);

            BetaRollbackOutcome outcome = await this.Flow(store).PlanAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, this.Options());

            Assert.True(outcome.IsPlanned, outcome.Error);
            Assert.Equal(BetaRollbackKind.RestoreFromProduction, outcome.Plan!.Kind);

            // Production's board file and the published sheet's bytes go back - the bytes first, so
            // BETA never holds a board file naming bytes it lacks.
            Assert.Equal(
                [DataTreeBuilder.Stored("PUBLISHED", BetaRollbackFlowTests.Sheet), DataTreeLayout.BoardFilePath(BetaRollbackFlowTests.BoardId)],
                outcome.Plan.Restored);

            Assert.Contains(DataTreeBuilder.Stored("ONLY IN BETA", BetaRollbackFlowTests.NewImage), outcome.Plan.Removed);
            Assert.Contains(DataTreeBuilder.Stored("SUBMITTED", BetaRollbackFlowTests.Sheet), outcome.Plan.Removed);
            Assert.Equal("hest@mailscan.dk", Assert.Single(outcome.Plan.Returning).ContactEmail);

            // A plan writes nothing.
            Assert.Equal("SUBMITTED", this.BetaHolds(BetaRollbackFlowTests.Sheet));
        }

        // ###########################################################################################
        // *** A SIGNED-IN CONTRIBUTOR IS NAMED BY THEIR ACCOUNT'S ADDRESS (code review, 2026-09-29).
        // *** Their submission carries the account and no contact address, so the confirmation named
        // "(no contact address)" and the push-back mail - sent to Returning's addresses - reached
        // nobody. The plan now resolves the account's address, as the Boards screen already did.
        // ###########################################################################################
        [Fact]
        public async Task A_signed_in_contributor_is_named_and_mailed_at_their_accounts_address()
        {
            FakeSubmissionStore store = this.StoreAheadOfProduction();
            var accounts = new FakeAccountStore();
            accounts.Accounts[42] = new AccountRecord(
                42, "anna@example.com", "anna@example.com", "hash", "Anna",
                IsVerified: true, IsAdministrator: false, IsLocked: false, BetaRollbackFlowTests.Now, null);

            long id = await store.CreateAsync(
                new NewSubmission(
                    BetaRollbackFlowTests.BoardId, "Commodore", "C64", "250407",
                    42, null, "192.0.2.1", "hash", "r0", "Corrected U8.", 1, [],
                    BetaRollbackFlowTests.Now.AddDays(-2), BetaRollbackFlowTests.Now.AddDays(-1)),
                CancellationToken.None);
            await store.SetStateAsync(id, SubmissionState.Merged, BetaRollbackFlowTests.Now.AddDays(-1), CancellationToken.None);

            this.BetaAheadOfProduction();

            BetaRollbackOutcome outcome = await this.Flow(store, accounts).PlanAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, this.Options());

            Assert.True(outcome.IsPlanned, outcome.Error);
            Assert.Equal("anna@example.com", Assert.Single(outcome.Plan!.Returning).ContactEmail);
        }

        // ###########################################################################################
        // *** A NEW BOARD LEAVING BETA LEAVES BETA'S DROP-DOWN LISTS TOO (2026-09-27). *** Its row
        // was added when it was published there; left in, every BETA user would be offered a board
        // that is gone. Its PLACEMENT is kept, so publishing it again puts it back in the same place.
        // ###########################################################################################
        [Fact]
        public async Task A_board_never_promoted_is_taken_out_of_BETAs_list_and_keeps_its_placement()
        {
            MasterListingRow other = ApprovePublishFlowTests.OtherListedBoard;
            MasterListingRow listed = new("Commodore 64", "250407", DataTreeBuilder.BoardFile, string.Empty);
            DataTreeBuilder.ListingMaster(this.thisBeta, other, listed);
            DataTreeBuilder.Board(this.thisBeta, other.BoardId);

            FakeSubmissionStore store = this.StoreAheadOfProduction(everPromoted: false);
            await this.MergedSubmissionAsync(store);
            await store.SetPlacementAsync(BetaRollbackFlowTests.BoardId, ApprovePublishFlowTests.Placement(), 1, BetaRollbackFlowTests.Now, CancellationToken.None);

            this.Beta((BetaRollbackFlowTests.Sheet, "THE NEW BOARD"));

            BetaRollbackOutcome outcome = await this.Flow(store).RollBackAsync(
                BetaRollbackFlowTests.Admin(), BetaRollbackFlowTests.BoardId, "Not ready.",
                this.Options(), BetaRollbackFlowTests.Now);

            Assert.True(outcome.IsDone, outcome.Error);
            Assert.Equal([other], DataTreeBuilder.ListedIn(this.thisBeta));
            Assert.True(store.Placements.ContainsKey(BetaRollbackFlowTests.BoardId));
        }
    }
}
