using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ApprovePublishFlow - approving a submission, end to end (Phase 5, tasks 5 and 6).
    //
    // *** THIS IS THE ONE IRREVERSIBLE OPERATION IN THE WHOLE SYSTEM. *** Task 7 was struck, so no
    // publish history is kept: approving overwrites the published board in place and the only way
    // back is to publish a correction. Everything here is written on that basis - the flow refuses
    // early and often, and each refusal leaves the tree exactly as it was.
    //
    // The ORDER of its own checks is the property worth testing, not just their presence. Deciding
    // authority, then state, then the plan, then writing, means the expensive and dangerous work
    // only ever happens after the cheap refusals have all passed.
    // ###########################################################################################
    // Shares the "BoardFiles" collection - see PublishExecutorTests for why these must not run
    // in parallel.
    [Collection("BoardFiles")]
    public sealed class ApprovePublishFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot;
        private readonly string thisDataTree;
        private readonly string thisBlobRoot;

        public ApprovePublishFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-approve", Guid.NewGuid().ToString("N"));
            this.thisDataTree = Path.Combine(this.thisRoot, "app-data-BETA");
            this.thisBlobRoot = Path.Combine(this.thisRoot, "blobs");

            Directory.CreateDirectory(this.thisDataTree);
            Directory.CreateDirectory(this.thisBlobRoot);

            // *** THE TREE'S MASTER WORKBOOKS, because they ARE the generations. *** A real tree
            // carries both: the unversioned original serving every pre-2.0.0 build, and the v2.0.0
            // one serving 2.0.0 and newer. A NEW system has no files of its own to read a
            // generation from, so it takes the newest of these - without them a publish is refused
            // rather than writing the frozen unversioned tree.
            File.WriteAllText(Path.Combine(this.thisDataTree, "Classic-Repair-Toolbox.xlsx"), "master");
            File.WriteAllText(Path.Combine(this.thisDataTree, "Classic-Repair-Toolbox.v2.0.0.xlsx"), "master v2");
        }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(this.thisRoot))
                    Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
                // A leftover temp folder is harmless; failing a test over cleanup is not.
            }
        }

        private BlobStore Blobs() => new(this.thisBlobRoot, NullLogger<BlobStore>.Instance);

        private ApprovePublishFlow Flow(ISubmissionStore store) =>
            new(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                NullLogger<ApprovePublishFlow>.Instance);

        private static Handlers.Accounts.AccountRecord Account(
            bool administrator = true, bool reviewer = false) =>
            new(
                Id: 7,
                Email: "admin@example.com",
                NormalisedEmail: "admin@example.com",
                PasswordHash: "hash",
                DisplayName: "Admin",
                IsVerified: true,
                IsAdministrator: administrator,
                IsReviewer: reviewer,
                IsLocked: false,
                CreatedUtc: ApprovePublishFlowTests.Now,
                LastLoginUtc: null);

        private static SubmissionManifest Manifest(params SubmissionFile[] files)
        {
            var manifest = new SubmissionManifest
            {
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files = [.. files]
            };

            manifest.Rows.RevisionDate = "2026-September-22";

            manifest.Rows.Components.Add(new ComponentEntry
            {
                BoardLabel = "U8",
                FriendlyName = "PLA",
                TechnicalNameOrValue = "906114-01",
                PartNumber = "251715-01"
            });

            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1",
                BoardLabel = "U8",
                X = "100",
                Y = "200",
                Width = "40",
                Height = "20"
            });

            return manifest;
        }

        // Creates a pending submission with its payload stored, as a real one would be.
        private static async Task<(FakeSubmissionStore Store, long Id)> PendingAsync(
            SubmissionManifest? manifest = null)
        {
            var store = new FakeSubmissionStore();
            manifest ??= ApprovePublishFlowTests.Manifest();

            // *** THE FILES GO TO CreateAsync, NOT JUST INTO THE PAYLOAD. *** The store rebuilds a
            // reloaded manifest's Files from the submission_files ROWS rather than from the stored
            // payload - the real SQL store joins that table, and the fake models it. A test that
            // only saved the payload would hand the publish a manifest with no files at all and
            // then wonder why nothing was copied.
            long id = await store.CreateAsync(
                new NewSubmission(
                    manifest.SystemId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                    null, "someone@example.com", "192.0.2.1", "hash", "2026-August-21",
                    "Corrected U8.", 1, [.. manifest.Files], ApprovePublishFlowTests.Now,
                    ApprovePublishFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SavePayloadAsync(id, manifest, CancellationToken.None);
            await store.SetStateAsync(id, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

            return (store, id);
        }

        private async Task<string> PutBlobAsync(byte[] bytes)
        {
            string hash = Convert.ToHexStringLower(SHA256.HashData(bytes));

            BlobStore blobs = this.Blobs();
            using var content = new MemoryStream(bytes);
            await blobs.AppendChunkAsync(1, hash, 0, content, CancellationToken.None);
            await blobs.TryCompleteAsync(1, hash, CancellationToken.None);

            return hash;
        }

        // -----------------------------------------------------------------------------------
        // The happy path, asserted on what actually reached the disk.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task An_approved_submission_PUBLISHES_a_board_that_reads_back()
        {
            // *** THE TEST EVERYTHING ELSE RESTS ON. *** Read through the ordinary reader, so a
            // publish that writes files no client can load fails here rather than in the field.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            string workbook = Path.Combine(
                this.thisDataTree, "Commodore", "C64", "250407", "Data C64 250407 v2.0.0.xlsx");

            Assert.True(File.Exists(workbook));

            BoardData? published = await BoardDataReader.LoadAsync(
                workbook, "approve-" + Guid.NewGuid().ToString("N"));

            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
            Assert.Equal("2026-September-22", published.RevisionDate);
        }

        [Fact]
        public async Task The_published_boards_HIGHLIGHTS_survive_the_approval()
        {
            // The sidecar, end to end through the real flow rather than through PublishExecutor
            // directly - this is the path an actual Approve button takes.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            string workbook = Path.Combine(
                this.thisDataTree, "Commodore", "C64", "250407", "Data C64 250407 v2.0.0.xlsx");

            ComponentHighlightEntry highlight = Assert.Single(
                BoardComponentHighlightStorage.LoadComponentHighlights(workbook));

            Assert.Equal("U8", highlight.BoardLabel);
            Assert.Equal("100", highlight.X);
        }

        [Fact]
        public async Task The_submission_is_recorded_as_MERGED_with_the_reviewer_who_approved_it()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);

            Assert.Equal(SubmissionState.Merged, record!.State);
            Assert.Equal(7, store.Decisions[id].DecidedByAccountId);
        }

        [Fact]
        public async Task The_REVISION_recorded_is_the_boards_own_revision_DATE()
        {
            // *** A CONTRACT WITH THE CLIENT. *** DraftBaseRevision stamps a draft's BaseRevision
            // from the official board's revision date, and submissions send it back to be diffed
            // against. Recording anything else - a counter, a timestamp - would make every
            // contributor re-base against a value their drafts never carry.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.Equal(
                "2026-September-22",
                store.PublishedSystems["Commodore/C64/250407"].Revision);
        }

        [Fact]
        public async Task A_submitted_FILE_lands_in_the_published_tree()
        {
            byte[] image = Encoding.UTF8.GetBytes("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                // Data-root-relative, which is the only shape a real board produces - see
                // PublishPlan's header for what the old "Images/sheet1.png" shape concealed.
                ApprovePublishFlowTests.Manifest(new SubmissionFile
                {
                    Path = "Commodore/C64/250407/Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            Assert.Equal(image, await File.ReadAllBytesAsync(Path.Combine(
                this.thisDataTree, "Commodore", "C64", "250407", "Images", "sheet1.png")));
        }

        // -----------------------------------------------------------------------------------
        // The refusals. Each one must leave the tree untouched.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_REVIEWER_cannot_approve_and_NOTHING_is_written()
        {
            // *** THE STRUCTURAL GUARANTEE, asserted on the DISK rather than on a bool. *** Phase
            // 6 gives Reviewer a blast radius of "none - no published data can change"; the only
            // honest way to test that is to check nothing changed.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id,
                ApprovePublishFlowTests.Account(administrator: false, reviewer: true),
                this.thisDataTree,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.NotEmpty(outcome.Error);

            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));

            // ...and the submission is untouched, so somebody entitled can still approve it.
            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);
            Assert.Equal(SubmissionState.Pending, record!.State);
        }

        [Fact]
        public async Task An_ALREADY_MERGED_submission_is_refused_and_nothing_is_rewritten()
        {
            // The double-publish guard on the path that matters. Two administrators with the queue
            // open both press Approve; the second must be refused on the state the first wrote.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome first = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(first.IsPublished, first.Error);

            ApproveOutcome second = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.False(second.IsPublished);
            Assert.Contains("already", second.Error, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task A_submission_that_does_not_EXIST_is_refused()
        {
            var store = new FakeSubmissionStore();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                999, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsNotFound);
        }

        [Fact]
        public async Task A_submission_whose_PAYLOAD_cannot_be_loaded_is_refused()
        {
            // *** NOTHING IS PUBLISHED FROM A SUBMISSION NOBODY CAN READ. *** Without the payload
            // there are no rows, and publishing would write an EMPTY board over a real one -
            // deleting the system's entire contents, irreversibly.
            var store = new FakeSubmissionStore();

            long id = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "No payload.",
                    1, [], ApprovePublishFlowTests.Now, ApprovePublishFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        [Fact]
        public async Task A_MISSING_BLOB_refuses_the_whole_approval_rather_than_publishing_part_of_it()
        {
            // The manifest references a file the blob store does not hold. PublishExecutor stops
            // before the workbook, so the board is never made to reference a file that is not
            // there - and the submission stays pending so it can be retried once the blob is back.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(new SubmissionFile
                {
                    Path = "Images/missing.png", Sha256 = new string('a', 64), SizeBytes = 10
                }));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);

            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);
            Assert.Equal(SubmissionState.Pending, record!.State);
        }

        [Fact]
        public async Task A_BLANK_data_tree_root_is_refused_rather_than_writing_beside_the_service()
        {
            // A missing DataTreeRoot setting must not make the service's own working directory the
            // data tree - which is where its configuration and binaries live.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), string.Empty,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
        }

        [Fact]
        public async Task A_NEW_SYSTEM_publishes_into_the_NEWEST_generation_and_never_the_frozen_one()
        {
            // *** THE MAINTAINER'S RULE, and the bug that writing these tests caught. *** A new
            // system's folder is EMPTY, so the generation cannot be read from it - and the first
            // version of this flow fell back to "no version", publishing
            // "Data C64 250407.xlsx". That is the frozen file serving every pre-2.0.0 build, the
            // one file the maintainer said must never be written.
            //
            // The generation now comes from the TREE's master workbooks instead. Asserted by the
            // ABSENCE of the unversioned file as well as the presence of the versioned one: only
            // checking the latter would pass against a flow that wrote both.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");

            Assert.True(File.Exists(Path.Combine(folder, "Data C64 250407 v2.0.0.xlsx")));
            Assert.False(File.Exists(Path.Combine(folder, "Data C64 250407.xlsx")));
        }

        [Fact]
        public async Task A_tree_with_NO_versioned_master_REFUSES_rather_than_writing_the_frozen_file()
        {
            // The other half of the rule. With no generation to be found anywhere, publishing
            // would have to write the unversioned file - so it is refused outright rather than
            // guessed at. This is the case a fresh or half-set-up tree presents.
            File.Delete(Path.Combine(this.thisDataTree, "Classic-Repair-Toolbox.v2.0.0.xlsx"));

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        [Fact]
        public async Task An_EXISTING_system_keeps_publishing_into_its_own_generation()
        {
            // The anti-vacuity half: the tree-level fallback must only apply when the system's own
            // folder says nothing. A system already on v2.0.0 reads its generation from its own
            // files, as it always did.
            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");
            Directory.CreateDirectory(folder);
            await File.WriteAllTextAsync(Path.Combine(folder, "Data C64 250407 v2.0.0.xlsx"), "placeholder");

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(
                Path.Combine(folder, "Data C64 250407 v2.0.0.xlsx"),
                "existing-" + Guid.NewGuid().ToString("N"));

            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
        }

        [Fact]
        public async Task A_submitted_CALIBRATION_reaches_the_published_sidecar()
        {
            // *** THE END OF A CHAIN THAT WAS BROKEN IN THREE PLACES. *** A contributor's KiCad
            // calibration had to cross four boundaries to get here, and until 2026-09-22 it fell
            // out at the first: SubmissionManifestBuilder takes a BoardData, which has no
            // calibration section, so the manifest's list was always empty. The publish path
            // handled it correctly the whole time and never received anything.
            //
            // This asserts the whole chain by reading the PUBLISHED sidecar back through the
            // shipped reader - the same file CRT itself loads.
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();

            manifest.Rows.KiCadCalibrations.Add(new KiCadCalibrationEntry
            {
                SchematicName = "Sheet 1",
                CadName = "board.kicad_pcb",
                OffsetX = 1.5,
                OffsetY = 2.5,
                ScaleX = 0.75,
                ScaleY = 0.8,
                MirrorX = true,
                MirrorY = false
            });

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(manifest);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            string workbook = Path.Combine(
                this.thisDataTree, "Commodore", "C64", "250407", "Data C64 250407 v2.0.0.xlsx");

            Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                workbook, "Sheet 1",
                out string cadName, out double offsetX, out _,
                out double scaleX, out _, out bool mirrorX, out _));

            Assert.Equal("board.kicad_pcb", cadName);
            Assert.Equal(1.5, offsetX);
            Assert.Equal(0.75, scaleX);
            Assert.True(mirrorX);
        }

        [Fact]
        public async Task A_NULL_account_is_refused()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, null, this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
        }

        // -----------------------------------------------------------------------------------
        // Republishing over an existing board.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Approving_a_SECOND_submission_replaces_the_first_ones_rows()
        {
            // *** THE MANIFEST IS THE COMPLETE INTENDED STATE. *** The second submission carries
            // only U9, so U8 must be GONE afterwards - the contributor deleted it. An overlay
            // would resurrect it and nothing would report that.
            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync();

            await this.Flow(store).ApproveAsync(
                first, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            var second = new SubmissionManifest
            {
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407"
            };

            second.Rows.RevisionDate = "2026-October-01";
            second.Rows.Components.Add(new ComponentEntry
            {
                BoardLabel = "U9", FriendlyName = "VIC", TechnicalNameOrValue = "6569"
            });

            long secondId = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "other@example.com", "192.0.2.2", "hash2", "2026-September-22",
                    "Replaced U8 with U9.", 1, [], ApprovePublishFlowTests.Now,
                    ApprovePublishFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SavePayloadAsync(secondId, second, CancellationToken.None);
            await store.SetStateAsync(secondId, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                secondId, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(
                Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "Data C64 250407 v2.0.0.xlsx"),
                "approve2-" + Guid.NewGuid().ToString("N"));

            Assert.Equal("U9", Assert.Single(published!.Components).BoardLabel);
        }
    }
}
