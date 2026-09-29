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
    // Covers PublishExecutor - the step that actually writes the data tree.
    //
    // THE BLOB STORE AND THE FILESYSTEM ARE REAL HERE, against a temp folder, for the same reason
    // SubmissionFlowTests uses a real one: the properties worth testing ARE the filesystem
    // behaviour. That a published board reads back through BoardDataReader, that a missing blob
    // stops the publish BEFORE the workbook is touched, that re-running is safe - none of these
    // mean anything against a fake filesystem.
    //
    // THE PROPERTIES THAT MATTER MOST, in order:
    //   1. A published board actually loads afterwards. If it does not, nothing else matters.
    //   2. A failure leaves the tree recoverable, never half-rewritten and broken.
    //   3. The content hash accounts for the generated workbook, or no client re-downloads.
    //   4. The revision reaches the systems row, or every later submission re-bases wrongly.
    //
    // *** SHARED COLLECTION: these classes must NOT run in parallel with each other. ***
    //
    // They all drive EPPlus and BoardDataReader's static cache against real files on disk. xunit
    // gives every class WITHOUT a [Collection] its own collection and runs those in parallel, so
    // `parallelizeTestCollections: false` in xunit.runner.json does not serialise them - it only
    // stops collections interleaving once they exist.
    //
    // *** THE NOTE THAT USED TO BE HERE WAS WRONG, AND IS KEPT AS A WARNING. *** It claimed this
    // collection fixed Re_running_the_same_publish_is_safe failing "roughly one run in five", on
    // the reasoning that a test passing alone but failing in a full run must be a cross-class
    // race. That inference is seductive and was false: the test kept failing intermittently
    // afterwards.
    //
    // The REAL cause, found 2026-09-22, was that BoardWorkbookWriter produced non-deterministic
    // bytes - an .xlsx is a ZIP and EPPlus stamped every entry with the current clock, so two
    // publishes either side of a second boundary hashed differently. The test was right and the
    // PRODUCTION CODE was wrong; a full run is simply slower, which is what made it cross that
    // boundary. See BoardWorkbookWriter.MakeDeterministic.
    //
    // The lesson worth keeping: "passes alone, fails in a full run" narrows the cause to something
    // timing- or order-dependent, and shared static state is only ONE candidate. Wall-clock
    // sensitivity is another, and it looks identical from the outside.
    //
    // CRT.Data.Tests already does this - BoardDataReaderTests and BoardWorkbookWriterTests share
    // a "BoardData" collection for the same reason, and CLAUDE.md states the rule outright: any
    // new test class touching one of these statics must join its collection.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class PublishExecutorTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot;
        private readonly string thisBlobRoot;
        private readonly string thisSystemFolder;

        public PublishExecutorTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-publish-tests", Guid.NewGuid().ToString("N"));
            this.thisBlobRoot = Path.Combine(this.thisRoot, "blobs");
            this.thisSystemFolder = Path.Combine(this.thisRoot, "beta", "Commodore", "C64", "250407");

            Directory.CreateDirectory(this.thisBlobRoot);

            // *** THE TREE'S MASTER WORKBOOKS, because they ARE the generations. *** A real tree
            // carries both - the unversioned original serving every pre-2.0.0 build, and the
            // v2.0.0 one serving 2.0.0 and newer. A system with no files of its OWN takes its
            // generation from these; without them PublishPlan refuses rather than writing the
            // frozen unversioned file, which is the rule added 2026-09-22.
            Directory.CreateDirectory(Path.Combine(this.thisRoot, "beta"));
            File.WriteAllText(Path.Combine(this.thisRoot, "beta", "Classic-Repair-Toolbox.xlsx"), "master");
            File.WriteAllText(Path.Combine(this.thisRoot, "beta", "Classic-Repair-Toolbox.v2.0.0.xlsx"), "master v2");
            Directory.CreateDirectory(this.thisSystemFolder);
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

        private PublishExecutor Executor(ISubmissionStore store) =>
            new(this.Blobs(), store, NullLogger<PublishExecutor>.Instance);

        private static string HashOf(byte[] bytes) => Convert.ToHexStringLower(SHA256.HashData(bytes));

        // Puts real bytes into the blob store the way an upload would, so TryCopyToAsync can find
        // them. Going through the store rather than writing the file by hand means these tests
        // break if the store's own layout changes, which is correct.
        private async Task<string> PutBlobAsync(byte[] bytes)
        {
            string hash = PublishExecutorTests.HashOf(bytes);

            BlobStore blobs = this.Blobs();
            using var content = new MemoryStream(bytes);
            await blobs.AppendChunkAsync(1, hash, 0, content, CancellationToken.None);
            await blobs.TryCompleteAsync(1, hash, CancellationToken.None);

            return hash;
        }

        private static BoardData Board() => new()
        {
            RevisionDate = "2026-August-21",
            Schematics =
            [
                new BoardSchematicEntry
                {
                    SchematicName = "Sheet 1",
                    SchematicImageFile = "Images/sheet1.png",
                    SchematicHighlightOpacity = "0.35"
                }
            ],
            Components =
            [
                new ComponentEntry
                {
                    BoardLabel = "U8",
                    FriendlyName = "PLA",
                    TechnicalNameOrValue = "906114-01",
                    PartNumber = "251715-01"
                }
            ]
        };

        // ###########################################################################################
        // Every file is CITED by a row (security review, 2026-09-25): the plan now refuses a file no
        // row uses, since nothing would show it and no maintainer could have seen it.
        // ###########################################################################################
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

            foreach (SubmissionFile file in files)
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

            return manifest;
        }

        // ###########################################################################################
        // Bytes that ARE a PNG as far as their opening goes (security review, 2026-09-25): a publish
        // now checks every file's bytes against its name before writing anything, so "PNGDATA"
        // alone - which used to stand in for an image - is correctly refused as not being one.
        // ###########################################################################################
        internal static byte[] Png(string marker) =>
            [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, .. Encoding.UTF8.GetBytes(marker)];


        private PublishPlanDetail Plan(SubmissionManifest manifest, IEnumerable<string>? existing = null)
        {
            PublishPlanResult result = PublishPlan.Build(
                Path.Combine(this.thisRoot, "beta"),
                this.thisSystemFolder,
                existing ?? [],
                "Data C64 250407",
                manifest,

                // ###########################################################################################
                // *** THE PLAN'S REVISION IS NOW THE ONE THE WORKBOOK GETS (2026-09-26). *** It used
                // to be a placeholder ("r2") because the executor stamped the workbook itself and
                // nothing compared the two - which was the defect: the board and
                // `systems.current_revision` disagreed. ApprovePublishFlow stamps it once now, so
                // the fixture passes what that would produce for `Now`.
                // ###########################################################################################
                BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now),
                PublishExecutorTests.Now,
                ["Someone"],
                SystemDescriptorRules.SystemOrigin.Contributed);

            Assert.True(result.IsPlanned, "the fixture's own plan must be valid");
            return result.Plan!;
        }

        private static FakeSubmissionStore StoreWithSystem()
        {
            var store = new FakeSubmissionStore();

            // The real store's UPDATE targets a row CreateAsync inserted, so a publish for an
            // unregistered system updates nothing. The fake refuses it outright; these tests
            // therefore have to register the system first, exactly as a real submission does.
            store.Systems["Commodore/C64/250407"] = new NewSubmission(
                "Commodore/C64/250407",
                "Commodore",
                "C64",
                "250407",
                null,
                "someone@example.com",
                "192.0.2.1",
                "hash",
                "r1",
                "Fixed a typo.",
                1,
                [],
                PublishExecutorTests.Now,
                PublishExecutorTests.Now.AddHours(24));

            return store;
        }

        // -----------------------------------------------------------------------------------
        // The board actually publishes and loads
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_published_board_reads_back_through_the_ordinary_reader()
        {
            // *** THE TEST EVERYTHING ELSE RESTS ON. *** A publish that writes files a client
            // cannot read is worse than no publish at all, because it replaces a board that
            // worked.
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Commodore/C64/250407/Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished);

            BoardData? published = await BoardDataReader.LoadAsync(plan.WorkbookPath, "publish-test-" + Guid.NewGuid().ToString("N"));

            Assert.NotNull(published);

            // ###########################################################################################
            // *** THE REVISION DATE IS NOW THE PUBLISH DATE, NOT THE SUBMITTED ONE (corrected
            // 2026-09-23, owner request). ***
            //
            // This used to assert the submitted "2026-August-21" survived untouched. It did, and
            // that was the defect: the line means "when this board was last published", so carrying
            // the contributor's own value through dated a freshly published board to whenever they
            // happened to start their draft - a week stale on anything that waited in the queue.
            //
            // `Now` is 2026-09-21, and BoardWorkbookStyle.FormatRevisionDate renders it in the
            // shape the shipped boards use.
            // ###########################################################################################
            Assert.Equal("2026-September-21", published!.RevisionDate);
            Assert.Equal("U8", Assert.Single(published.Components).BoardLabel);
            Assert.Equal("251715-01", published.Components[0].PartNumber);

            // The numeric-looking value survives the publish, not just the writer's own tests.
            Assert.Equal("0.35", Assert.Single(published.Schematics).SchematicHighlightOpacity);

            // ###########################################################################################
            // *** AND THE `systems` ROW SAYS THE SAME (owner confirmation, 2026-09-26: "the revision
            // date gets updated from server ... so server always wins"). ***
            //
            // The workbook is stamped with the publish date above, but the row was written from the
            // PLAN's revision, which is the SUBMITTED value - so the board said "2026-September-21"
            // while `systems.current_revision` said "2026-August-21". That row is what a
            // contributor's next draft re-bases against (DraftBaseRevision), so the two disagreeing
            // is the drift check comparing against a revision no board ever carried.
            // ###########################################################################################
            Assert.Equal(published.RevisionDate, store.PublishedSystems["Commodore/C64/250407"].Revision);
        }

        // -----------------------------------------------------------------------------------
        // The SIDECAR - the board's other half.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_published_boards_HIGHLIGHTS_read_back_through_the_ordinary_reader()
        {
            // *** THE DATA-LOSS TEST. *** A board is TWO files: the workbook and the `.json`
            // beside it holding every component highlight. BoardWorkbookWriter writes only the
            // workbook, so before the sidecar was wired in, publishing produced a board whose
            // highlights had silently vanished - with no retained revision to restore from.
            //
            // Read back through the SHIPPED reader rather than by parsing the JSON here, so a
            // writer that agreed with a test-only parser could not pass.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            BoardData board = PublishExecutorTests.Board();

            board.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1",
                BoardLabel = "U8",
                X = "100",
                Y = "200",
                Width = "40",
                Height = "20"
            });

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, board, [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished);

            List<ComponentHighlightEntry> highlights =
                BoardComponentHighlightStorage.LoadComponentHighlights(plan.WorkbookPath);

            ComponentHighlightEntry published = Assert.Single(highlights);

            Assert.Equal("U8", published.BoardLabel);
            Assert.Equal("100", published.X);
        }

        [Fact]
        public async Task A_KiCad_CALIBRATION_survives_the_publish()
        {
            // Calibrations are not part of BoardData at all - they live only in the sidecar - so
            // they are passed separately and would be the easiest thing to drop.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(store).ExecuteAsync(
                plan,
                PublishExecutorTests.Board(),
                [
                    new KiCadCalibrationEntry
                    {
                        SchematicName = "Sheet 1",
                        CadName = "board.kicad_pcb",
                        OffsetX = 1.5,
                        ScaleX = 0.75,
                        MirrorX = true
                    }
                ],
                submissionId: 1,
                PublishExecutorTests.Now);

            Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                plan.WorkbookPath, "Sheet 1",
                out string cadName, out double offsetX, out _,
                out double scaleX, out _, out bool mirrorX, out _));

            Assert.Equal("board.kicad_pcb", cadName);
            Assert.Equal(1.5, offsetX);
            Assert.Equal(0.75, scaleX);
            Assert.True(mirrorX);
        }

        [Fact]
        public async Task A_HIGHLIGHT_ONLY_change_MOVES_the_published_content_hash()
        {
            // *** THE SILENT FAILURE THE SIDECAR HASH EXISTS TO PREVENT, end to end. *** Moving a
            // highlight changes no workbook row - highlights are not in the workbook - and uploads
            // no file. So neither the uploaded-file list nor the workbook hash moves, and without
            // the sidecar folded into the descriptor the content hash would be identical to the
            // previous publish's. No client would re-download, and the moved highlight would never
            // reach anybody.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            BoardData first = PublishExecutorTests.Board();
            first.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "100", Y = "200", Width = "40", Height = "20"
            });

            await this.Executor(store).ExecuteAsync(
                this.Plan(PublishExecutorTests.Manifest()), first, [], submissionId: 1, PublishExecutorTests.Now);

            string before = store.PublishedSystems["Commodore/C64/250407"].ContentHash;

            // The same board with the highlight MOVED, and nothing else changed at all.
            BoardData second = PublishExecutorTests.Board();
            second.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "500", Y = "600", Width = "40", Height = "20"
            });

            await this.Executor(store).ExecuteAsync(
                this.Plan(PublishExecutorTests.Manifest()), second, [], submissionId: 2, PublishExecutorTests.Now);

            Assert.NotEqual(before, store.PublishedSystems["Commodore/C64/250407"].ContentHash);
        }

        [Fact]
        public async Task NULL_calibrations_are_REFUSED_rather_than_published_as_none()
        {
            // A caller passing null has almost certainly failed to read the submission's
            // calibrations rather than genuinely meaning "this board has none", and treating the
            // two alike publishes a board with a contributor's calibration work silently removed.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await Assert.ThrowsAsync<ArgumentNullException>(() =>
                this.Executor(store).ExecuteAsync(
                    plan, PublishExecutorTests.Board(), null!, submissionId: 1, PublishExecutorTests.Now));
        }

        [Fact]
        public async Task The_sidecar_sits_BESIDE_the_workbook_where_the_reader_looks_for_it()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(store).ExecuteAsync(
                plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(File.Exists(Path.ChangeExtension(plan.WorkbookPath, ".json")));
            Assert.Equal(Path.ChangeExtension(plan.WorkbookPath, ".json"), plan.SidecarPath);
        }

        [Fact]
        public async Task The_submitted_files_land_at_their_planned_paths_with_their_real_bytes()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            // *** A REAL SUBMITTED PATH IS DATA-ROOT-RELATIVE (corrected 2026-09-23). *** This test
            // used to pass "Images/sheet1.png" and expect it under the SYSTEM folder. That shape
            // does not occur: a board stores "Commodore/C64/250407/Images/sheet1.png", and the
            // unrealistic input is what let the publish resolve every file one level too deep -
            // writing the whole board into a copy of itself on the first real publish.
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile
                {
                    Path = "Commodore/C64/250407/Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            string landed = Path.Combine(this.thisSystemFolder, "Images", "sheet1.png");

            Assert.True(File.Exists(landed));
            Assert.Equal(image, await File.ReadAllBytesAsync(landed));
        }

        // ###########################################################################################
        // *** system.json IS RETIRED (owner decision, 2026-09-25). *** It was written beside
        // every published board and so downloaded by every user, who had no use for it. A publish
        // writes none, and removes one an earlier build left behind. What it recorded - revision
        // and content hash - still reaches the database, and the outcome still reports it.
        // ###########################################################################################
        [Fact]
        public async Task No_system_json_is_written_and_one_left_by_an_earlier_build_is_removed()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            string leftover = Path.Combine(this.thisSystemFolder, SystemDescriptorStore.FileName);
            Directory.CreateDirectory(this.thisSystemFolder);
            await File.WriteAllTextAsync(leftover, "{}");

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.False(File.Exists(leftover));

            Assert.Equal(
                BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now),
                outcome.Descriptor!.Revision);
            Assert.NotEmpty(outcome.Descriptor.ContentHash);
        }

        // -----------------------------------------------------------------------------------
        // The content hash accounts for the generated workbook
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_rows_only_publish_still_moves_the_content_hash()
        {
            // *** THE SILENT FAILURE. *** A typo fix uploads NO files, so the uploaded-file list
            // is identical between two publishes. Without the generated workbook folded into the
            // hash, system.json would advertise an unchanged system and no client would ever
            // re-download the board that just changed.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome first = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            var changed = new BoardData
            {
                RevisionDate = "2026-August-21",
                Schematics = PublishExecutorTests.Board().Schematics,
                Components =
                [
                    new ComponentEntry
                    {
                        BoardLabel = "U8",
                        FriendlyName = "PLA",
                        TechnicalNameOrValue = "906114-01",
                        PartNumber = "251715-02"
                    }
                ]
            };

            PublishOutcome second = await this.Executor(store)
                .ExecuteAsync(plan, changed, [], submissionId: 1, PublishExecutorTests.Now);

            Assert.NotEqual(first.Descriptor!.ContentHash, second.Descriptor!.ContentHash);

            // *** AND THIS IS WHAT MAKES THE FOLD-IN LOAD-BEARING. *** Both publishes used the
            // SAME plan, whose descriptor covers uploaded files only - so had the executor
            // written that planned descriptor instead, both system.json files would have carried
            // this identical hash and no client would have re-downloaded anything. The assertion
            // above only means something because of this one.
            Assert.NotEqual(plan.Descriptor.ContentHash, first.Descriptor.ContentHash);
            Assert.NotEqual(plan.Descriptor.ContentHash, second.Descriptor.ContentHash);
        }

        [Fact]
        public async Task The_reported_workbook_hash_matches_the_file_on_disk()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            byte[] written = await File.ReadAllBytesAsync(plan.WorkbookPath);

            Assert.Equal(PublishExecutorTests.HashOf(written), outcome.WorkbookSha256);
        }

        // -----------------------------------------------------------------------------------
        // The database row
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_revision_and_content_hash_reach_the_systems_row()
        {
            // current_revision is the base a contributor's NEXT submission is diffed against. A
            // publish that fails to record it leaves every later submission re-basing against a
            // revision that no longer describes the tree.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            PublishedSystemRow row = store.PublishedSystems["Commodore/C64/250407"];

            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now), row.Revision);
            Assert.Equal(outcome.Descriptor!.ContentHash, row.ContentHash);
        }

        [Fact]
        public async Task The_submission_is_marked_merged()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            long id = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "Fixed a typo.",
                    1, [], PublishExecutorTests.Now, PublishExecutorTests.Now.AddHours(24)),
                CancellationToken.None);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], id, PublishExecutorTests.Now);

            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_publish_with_no_submission_behind_it_is_allowed()
        {
            // The project owner correcting their own data publishes without a submission. The
            // system's revision still has to be recorded.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: null, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished);
            Assert.Equal(
                BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now),
                store.PublishedSystems["Commodore/C64/250407"].Revision);
        }

        // -----------------------------------------------------------------------------------
        // Failure leaves the tree recoverable
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_missing_blob_stops_the_publish_BEFORE_the_workbook_is_touched()
        {
            // *** THE ORDERING PROPERTY. *** Files are copied first because they are inert until
            // the workbook references them. Writing the workbook first and failing on a file
            // would leave a board referencing images that never arrived - a board that fails to
            // load, replacing one that worked.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile
                {
                    Path = "Commodore/C64/250407/Images/sheet1.png",
                    Sha256 = new string('a', 64),   // never uploaded
                    SizeBytes = 10
                }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("Commodore/C64/250407/Images/sheet1.png", outcome.Failure);

            Assert.False(File.Exists(plan.WorkbookPath));
            Assert.False(File.Exists(Path.Combine(this.thisSystemFolder, SystemDescriptorStore.FileName)));
            Assert.Empty(store.PublishedSystems);
        }

        // ###########################################################################################
        // *** A DESTINATION THAT CANNOT BE WRITTEN IS A REPORTED FAILURE, NOT A 500 (owner
        // report, 2026-09-23). ***
        //
        // TryCopyToAsync returns false for a MISSING blob - covered above - but THROWS when the
        // destination refuses the write. Nothing caught that, so it escaped through
        // ApprovePublishFlow and the endpoint, and "Approve and publish" answered a bare 500.
        //
        // The real cause was a deployment one: board files replaced over a network share lost
        // their group ownership, so the service could create new files but not overwrite existing
        // ones. That is precisely the kind of fault that needs to NAME the file - which is what a
        // 500 cannot do.
        //
        // Provoked by making the destination a DIRECTORY. Opening it as a file throws
        // UnauthorizedAccessException, the same type the server hit, without needing to
        // manipulate real permissions (which behave differently on the Linux CI runner and on a
        // Windows developer machine).
        // ###########################################################################################
        [Fact]
        public async Task A_destination_that_cannot_be_written_STOPS_the_publish_and_names_the_file()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile
                {
                    Path = "Commodore/C64/250407/Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            // A directory exactly where the file must go.
            Directory.CreateDirectory(plan.Files[0].AbsolutePath);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            // Reported, not thrown - the assertion that fails against the version that shipped,
            // where this test would end in an unhandled UnauthorizedAccessException.
            Assert.False(outcome.IsPublished);
            Assert.Contains("Commodore/C64/250407/Images/sheet1.png", outcome.Failure);

            // And it stopped BEFORE touching the board, so nothing downstream is half-written.
            Assert.False(File.Exists(plan.WorkbookPath));
            Assert.False(File.Exists(Path.Combine(this.thisSystemFolder, SystemDescriptorStore.FileName)));
            Assert.Empty(store.PublishedSystems);
        }

        // ###########################################################################################
        // *** A FOLDER THE SERVICE MAY NOT WRITE STOPS THE PUBLISH BEFORE ANYTHING IS WRITTEN
        // (owner report, 2026-09-28). *** Production's files had been copied into BETA by hand as
        // root; the approval wrote the new workbook and was then refused at the highlight file
        // beside it - a 500 over a half-replaced board. Now every folder the publish writes into is
        // asked first, and the maintainer is told which, with nothing changed. (A folder that may
        // not be written cannot be made on every OS, so the executor's probe is told the answer.)
        // ###########################################################################################
        [Fact]
        public async Task A_board_folder_the_service_may_not_write_STOPS_the_publish_before_anything_is_written()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile
                {
                    Path = "Commodore/C64/250407/Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            var executor = new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance)
            {
                CanWriteFolderForTests = folder => folder != this.thisSystemFolder
            };

            PublishOutcome outcome = await executor
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("[Commodore/C64/250407]", outcome.Failure);
            Assert.Contains("nothing was changed", outcome.Failure);

            // Not the image (its "Images" folder is not there yet, so the board folder is what must
            // allow it), not the workbook, not the database.
            Assert.False(File.Exists(plan.Files[0].AbsolutePath));
            Assert.False(File.Exists(plan.WorkbookPath));
            Assert.False(File.Exists(plan.SidecarPath));
            Assert.Empty(store.PublishedSystems);
        }

        // ###########################################################################################
        // *** A HIGHLIGHT FILE THAT CANNOT BE WRITTEN IS REPORTED, NOT THROWN (owner report,
        // 2026-09-28). *** The sidecar step caught nothing, so the refusal escaped as a bare 500
        // "The server answered 500." with the workbook already written. It now says the board is
        // part-published and that approving again repairs it, and records nothing in the database.
        // Provoked with a DIRECTORY where the sidecar goes, as the file test above does.
        // ###########################################################################################
        [Fact]
        public async Task A_highlight_file_that_cannot_be_written_is_a_REPORTED_part_publish_not_an_exception()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            Directory.CreateDirectory(plan.SidecarPath);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("highlight file could not be written", outcome.Failure);
            Assert.Contains("part-published", outcome.Failure);
            Assert.Empty(store.PublishedSystems);
        }

        // ###########################################################################################
        // *** FILES COPIED IN BY HAND ARE REPLACED, NOT REFUSED (owner report, 2026-09-28). *** A
        // file copied into BETA as root may be replaced by the service - that is its folder's
        // permission - but not opened for writing, and the highlight file was written by opening
        // it. Both board files are now written beside themselves and renamed into place, so the
        // publish goes through. Made here as files with no write permission for their owner, which
        // refuses an open the same way; not on Windows (no such mode) and not as root (which
        // ignores it). Fails against the in-place sidecar write.
        // ###########################################################################################
        [Fact]
        public async Task A_board_whose_files_may_not_be_opened_for_writing_is_still_published()
        {
            // The return is for the platform analyzer, which cannot see that Skip throws.
            if (OperatingSystem.IsWindows())
            {
                Assert.Skip("Unix permissions only.");
                return;
            }
            Assert.SkipWhen(Environment.UserName == "root", "root may write anything.");

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            File.WriteAllText(plan.WorkbookPath, "copied in by hand");
            File.WriteAllText(plan.SidecarPath, "{}");
            File.SetUnixFileMode(plan.WorkbookPath, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
            File.SetUnixFileMode(plan.SidecarPath, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);

            // The board carries a highlight the hand-copied "{}" does not, so the sidecar really
            // has to be REPLACED. Without it WriteIfChanged rightly finds "{}" the same as a board
            // with no highlights, leaves it alone, and this test proves nothing about the replace
            // (it went red on CI exactly that way once WriteIfChanged arrived).
            BoardData board = PublishExecutorTests.Board();
            board.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "100", Y = "200", Width = "40", Height = "20"
            });

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, board, [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);

            BoardData? published = await BoardDataReader.LoadAsync(plan.WorkbookPath, "hand-copied-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
            Assert.Equal("U8", Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(plan.WorkbookPath)).BoardLabel);
        }

        [Fact]
        public async Task An_interrupted_publish_leaves_the_PREVIOUS_board_loadable()
        {
            // The recovery property stated in PublishExecutor's header: a failure leaves files
            // nothing points at - wasted space - rather than a broken board.
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string good = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail first = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Commodore/C64/250407/Images/sheet1.png", Sha256 = good, SizeBytes = image.LongLength }));

            await this.Executor(store)
                .ExecuteAsync(first, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            // A second publish whose blob is missing.
            PublishPlanDetail broken = this.Plan(
                PublishExecutorTests.Manifest(
                    new SubmissionFile { Path = "Commodore/C64/250407/Images/new.png", Sha256 = new string('b', 64), SizeBytes = 10 }),
                existing: [first.WorkbookFileName]);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(broken, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);

            // The board published by the FIRST run still loads.
            BoardData? still = await BoardDataReader.LoadAsync(first.WorkbookPath, "recover-" + Guid.NewGuid().ToString("N"));

            Assert.NotNull(still);
            Assert.Equal("U8", Assert.Single(still!.Components).BoardLabel);
        }

        // *** FORMERLY INTERMITTENT. TWO REAL BUGS CAME OUT OF IT, AND BOTH ARE FIXED. ***
        //
        // 1. PublishedBoardReader cached boards by FILE PATH, so the second publish was served the
        //    first one's cached board. Fixed 2026-09-21; PublishedBoardReaderTests pins it.
        //
        // 2. The remaining failures - roughly one run in five, and ONLY in a full-assembly run -
        //    were a CROSS-CLASS RACE. xunit gives every class without a [Collection] its own
        //    collection and runs those in parallel, so `parallelizeTestCollections: false` did not
        //    serialise them; three classes here drive EPPlus and BoardDataReader's static cache
        //    against real files at once. Fixed 2026-09-22 by the "BoardFiles" collection on this
        //    class, PublishedBoardReaderTests and ReviewAssetLocatorTests.
        //
        // The Windows-file-handle theory previously recorded here was WRONG, and the tell was in
        // the evidence all along: this test never failed when run in ISOLATION, however many
        // times, which a handle race within one test would not explain. Six consecutive
        // full-assembly runs are green since the collection was added.
        [Fact]
        public async Task Re_running_the_same_publish_is_safe()
        {
            // Stated in the header as the recovery for an interrupted publish: every step
            // overwrites, and content-addressed blobs re-copy byte-identically.
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Commodore/C64/250407/Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome first = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            PublishOutcome second = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(second.IsPublished);
            Assert.Equal(first.Descriptor!.ContentHash, second.Descriptor!.ContentHash);

            BoardData? published = await BoardDataReader.LoadAsync(plan.WorkbookPath, "rerun-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: everything is checked BEFORE anything is written.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** A REPLACED IMAGE IS NOT INERT. *** The old workbook already points at it, so a refusal
        // discovered at file 600 of 1,200 used to leave half a board replaced. Every blob is now
        // proved intact and of the type its name claims before the FIRST byte lands - so the good
        // file ahead of the bad one here must NOT have been written.
        // ###########################################################################################
        [Fact]
        public async Task A_file_whose_bytes_are_not_its_type_stops_the_publish_before_ANY_file_is_written()
        {
            byte[] good = PublishExecutorTests.Png("GOOD");
            byte[] notAnImage = Encoding.UTF8.GetBytes("MZ not a picture at all");

            string goodHash = await this.PutBlobAsync(good);
            string badHash = await this.PutBlobAsync(notAnImage);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Commodore/C64/250407/Images/a.png", Sha256 = goodHash, SizeBytes = good.LongLength },
                new SubmissionFile { Path = "Commodore/C64/250407/Images/z.png", Sha256 = badHash, SizeBytes = notAnImage.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("Images/z.png", outcome.Failure);

            Assert.False(File.Exists(Path.Combine(this.thisSystemFolder, "Images", "a.png")));
            Assert.False(File.Exists(plan.WorkbookPath));
            Assert.Empty(store.PublishedSystems);
        }

        // A blob that no longer matches its hash - changed on disk after it was accepted - is found
        // before anything is written, rather than copied into every user's data on trust.
        [Fact]
        public async Task A_blob_changed_in_the_store_stops_the_publish_before_anything_is_written()
        {
            byte[] image = PublishExecutorTests.Png("ORIGINAL");
            string hash = await this.PutBlobAsync(image);

            string stored = Path.Combine(this.thisBlobRoot, BlobStorePaths.BlobFolderName, hash[..2], hash[2..4], hash);
            await File.WriteAllBytesAsync(stored, PublishExecutorTests.Png("TAMPERED"));

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Commodore/C64/250407/Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.False(File.Exists(Path.Combine(this.thisSystemFolder, "Images", "sheet1.png")));
            Assert.False(File.Exists(plan.WorkbookPath));
        }

        // Another board's file cited UNCHANGED is part of the system's content but is never
        // written - it is already there, and writing it would only be a chance to get it wrong.
        [Fact]
        public async Task Another_boards_file_cited_unchanged_is_never_written()
        {
            string beta = Path.Combine(this.thisRoot, "beta");
            string foreignFull = Path.Combine(beta, "Commodore", "C128", "310378", "notes.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(foreignFull)!);
            await File.WriteAllTextAsync(foreignFull, "published notes");
            DateTime written = File.GetLastWriteTimeUtc(foreignFull);

            string hash = PublishExecutorTests.HashOf(Encoding.UTF8.GetBytes("published notes"));

            PublishPlanResult result = PublishPlan.Build(
                beta,
                this.thisSystemFolder,
                [],
                "Data C64 250407",
                PublishExecutorTests.Manifest(new SubmissionFile { Path = "Commodore/C128/310378/notes.txt", Sha256 = hash, SizeBytes = 15 }),
                "r2",
                PublishExecutorTests.Now,
                ["Someone"],
                SystemDescriptorRules.SystemOrigin.Contributed,
                PublishedTreeProbe.For(beta));

            Assert.True(result.IsPlanned, string.Join(" ", result.Problems.Select(problem => problem.Message)));

            // The blob is deliberately NOT in the store: were the executor to try to write this
            // file, it would stop on the missing blob.
            PublishOutcome outcome = await this.Executor(PublishExecutorTests.StoreWithSystem())
                .ExecuteAsync(result.Plan!, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.Equal("published notes", await File.ReadAllTextAsync(foreignFull));
            Assert.Equal(written, File.GetLastWriteTimeUtc(foreignFull));
        }

        // -----------------------------------------------------------------------------------
        // The generation guard, end to end
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Publishing_writes_the_newest_generation_and_leaves_an_older_one_untouched()
        {
            // *** THE PROJECT OWNER'S RULE, PROVED ON DISK. *** An older generation is a frozen
            // compatibility target still serving older application builds. Writing it is silent
            // damage, so this asserts the older file is byte-identical afterwards rather than
            // merely that the newer one was written.
            string legacy = Path.Combine(this.thisSystemFolder, "Data C64 250407.xlsx");
            await File.WriteAllTextAsync(legacy, "the frozen generation");
            byte[] before = await File.ReadAllBytesAsync(legacy);

            string current = Path.Combine(this.thisSystemFolder, "Data C64 250407 v2.0.0.xlsx");
            await File.WriteAllTextAsync(current, "placeholder");

            FakeSubmissionStore store = PublishExecutorTests.StoreWithSystem();
            PublishPlanDetail plan = this.Plan(
                PublishExecutorTests.Manifest(),
                existing: ["Data C64 250407.xlsx", "Data C64 250407 v2.0.0.xlsx"]);

            Assert.Equal("Data C64 250407 v2.0.0.xlsx", plan.WorkbookFileName);

            await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.Equal(before, await File.ReadAllBytesAsync(legacy));

            // Anti-vacuity: the newest generation really was rewritten, so the assertion above is
            // about restraint rather than about nothing having happened at all.
            BoardData? published = await BoardDataReader.LoadAsync(current, "gen-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
        }
    }
}
