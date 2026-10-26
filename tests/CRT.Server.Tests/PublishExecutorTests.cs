using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
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
    // stops the publish BEFORE the board file is touched, that re-running is safe - none of these
    // mean anything against a fake filesystem.
    //
    // In the CRT 4.0 tree (2026-10-25) a publish stores the bytes the tree does not hold yet under
    // their checksum ("Files/<ab>/<sha256>.<ext>") and writes the board file "<board id>/board.json".
    //
    // THE PROPERTIES THAT MATTER MOST, in order:
    //   1. A published board actually loads afterwards. If it does not, nothing else matters.
    //   2. A failure leaves the tree recoverable, never half-rewritten and broken.
    //   3. The content hash accounts for the generated board file, or no client re-downloads.
    //   4. The revision reaches the boards row, or every later submission re-bases wrongly.
    //
    // *** SHARED COLLECTION: these classes must NOT run in parallel with each other. *** They drive
    // BoardDataReader's static cache against real files on disk.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class PublishExecutorTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot;
        private readonly string thisBeta;
        private readonly string thisBlobRoot;
        private readonly string thisBoardFolder;

        public PublishExecutorTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-publish-tests", Guid.NewGuid().ToString("N"));
            this.thisBlobRoot = Path.Combine(this.thisRoot, "blobs");
            this.thisBeta = Path.Combine(this.thisRoot, "beta");
            this.thisBoardFolder = Path.Combine(this.thisBeta, "Commodore", "C64", "250407");

            Directory.CreateDirectory(this.thisBlobRoot);
            Directory.CreateDirectory(this.thisBoardFolder);
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

        // Every file is CITED by a row (security review, 2026-09-25): the plan refuses a file no row
        // uses, since nothing would show it and no maintainer could have seen it.
        private static SubmissionManifest Manifest(params SubmissionFile[] files)
        {
            var manifest = new SubmissionManifest
            {
                BoardId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files = [.. files]
            };

            foreach (SubmissionFile file in files)
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

            return manifest;
        }

        // Bytes that ARE a PNG as far as their opening goes (security review, 2026-09-25): a publish
        // checks every file's bytes against its name before writing anything.
        internal static byte[] Png(string marker) =>
            [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, .. Encoding.UTF8.GetBytes(marker)];

        // The plan, with the revision ApprovePublishFlow stamps for `Now` - the one the board file
        // and the boards row both get (2026-09-26).
        private PublishPlanDetail Plan(SubmissionManifest manifest)
        {
            PublishPlanResult result = PublishPlan.Build(
                this.thisBeta,
                manifest,
                BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now),
                PublishExecutorTests.Now,
                ["Someone"],
                BoardDescriptorRules.BoardOrigin.Contributed,
                board: null,
                stored => PublishedBoardFiles.IsStored(this.thisBeta, stored));

            Assert.True(result.IsPlanned, "the fixture's own plan must be valid: " + string.Join(" ", result.Problems.Select(problem => problem.Message)));
            return result.Plan!;
        }

        private static FakeSubmissionStore StoreWithBoard()
        {
            var store = new FakeSubmissionStore();

            // The real store's UPDATE targets a row CreateAsync inserted, so a publish for an
            // unregistered board updates nothing. The fake refuses it outright; these tests
            // therefore have to register the board first, exactly as a real submission does.
            store.Boards["Commodore/C64/250407"] = new NewSubmission(
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

        private async Task<(PublishPlanDetail Plan, PublishOutcome Outcome, byte[] Image)> PublishSheetAsync(FakeSubmissionStore store, string marker = "PNGDATA")
        {
            byte[] image = PublishExecutorTests.Png(marker);
            string hash = await this.PutBlobAsync(image);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            return (plan, outcome, image);
        }

        // -----------------------------------------------------------------------------------
        // The board actually publishes and loads
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_published_board_reads_back_through_the_ordinary_reader()
        {
            // *** THE TEST EVERYTHING ELSE RESTS ON. *** A publish that writes a board a client
            // cannot read is worse than no publish at all, because it replaces a board that worked.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            (PublishPlanDetail plan, PublishOutcome outcome, _) = await this.PublishSheetAsync(store);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.Equal(Path.Combine(this.thisBoardFolder, "board.json"), plan.BoardFilePath);

            BoardData? published = await BoardDataReader.LoadAsync(plan.BoardFilePath, "publish-test-" + Guid.NewGuid().ToString("N"));

            Assert.NotNull(published);

            // The revision date is the PUBLISH date, not the submitted one (owner request,
            // 2026-09-23): it means "when this board was last published".
            Assert.Equal("2026-September-21", published!.RevisionDate);
            Assert.Equal("U8", Assert.Single(published.Components).BoardLabel);
            Assert.Equal("251715-01", published.Components[0].PartNumber);
            Assert.Equal("0.35", Assert.Single(published.Schematics).SchematicHighlightOpacity);
            Assert.Equal("Images/sheet1.png", published.Schematics[0].SchematicImageFile);

            // And the `boards` row says the same (owner confirmation, 2026-09-26: "server always
            // wins") - the row is what a contributor's next draft re-bases against.
            Assert.Equal(published.RevisionDate, store.PublishedBoards["Commodore/C64/250407"].Revision);
        }

        [Fact]
        public async Task The_board_file_lists_every_file_by_name_with_its_bytes()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            (PublishPlanDetail plan, _, byte[] image) = await this.PublishSheetAsync(store);

            BoardDocument? document = BoardFiles.ReadDocument(plan.BoardFilePath);

            BoardDocument.StoredFileRow file = Assert.Single(document!.Files);
            Assert.Equal("Images/sheet1.png", file.Name);
            Assert.Equal(PublishExecutorTests.HashOf(image), file.Sha256);
            Assert.Equal(image.LongLength, file.Size);
        }

        [Fact]
        public async Task A_file_is_stored_under_its_checksum_with_its_real_bytes()
        {
            // The submission named a file; the server decided where its bytes go.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            (_, _, byte[] image) = await this.PublishSheetAsync(store);

            string stored = Path.Combine(this.thisBeta, DataTreeLayout.StoredFilePath(PublishExecutorTests.HashOf(image), "sheet1.png").Replace('/', Path.DirectorySeparatorChar));

            Assert.Equal(image, await File.ReadAllBytesAsync(stored));
            Assert.False(File.Exists(Path.Combine(this.thisBoardFolder, "Images", "sheet1.png")), "Nothing is written under a name.");
        }

        [Fact]
        public async Task Two_names_of_the_same_bytes_are_stored_once()
        {
            byte[] image = PublishExecutorTests.Png("SAME");
            string hash = await this.PutBlobAsync(image);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "a.png", Sha256 = hash, SizeBytes = image.LongLength },
                new SubmissionFile { Path = "b.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome outcome = await this.Executor(PublishExecutorTests.StoreWithBoard())
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.Equal(1, outcome.FileCount);
            Assert.Equal(2, BoardFiles.ReadDocument(plan.BoardFilePath)!.Files.Count);
        }

        // ###########################################################################################
        // *** BYTES THE TREE ALREADY HOLDS ARE NEVER WRITTEN AGAIN. *** A one-cell typo fix on a
        // board of 1,200 files stores nothing - and the stored file is not even touched, so the
        // manifest's cached checksums and every client's copy stay as they are.
        // ###########################################################################################
        [Fact]
        public async Task Bytes_the_tree_already_holds_are_left_exactly_as_they_are()
        {
            byte[] image = PublishExecutorTests.Png("PUBLISHED");
            string hash = PublishExecutorTests.HashOf(image);

            string stored = Path.Combine(this.thisBeta, DataTreeLayout.StoredFilePath(hash, "u8.png").Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(stored)!);
            await File.WriteAllBytesAsync(stored, image);
            DateTime written = File.GetLastWriteTimeUtc(stored);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(new SubmissionFile { Path = "u8.png", Sha256 = hash, SizeBytes = image.LongLength }));

            // The blob is deliberately NOT in the store: were the executor to try to write this
            // file, it would stop on the missing blob.
            PublishOutcome outcome = await this.Executor(PublishExecutorTests.StoreWithBoard())
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.Equal(0, outcome.FileCount);
            Assert.Equal(written, File.GetLastWriteTimeUtc(stored));
            Assert.Equal("u8.png", Assert.Single(BoardFiles.ReadDocument(plan.BoardFilePath)!.Files).Name);
        }

        // -----------------------------------------------------------------------------------
        // The highlights, calibrations and renames - in the board file now
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_published_boards_HIGHLIGHTS_read_back_through_the_ordinary_reader()
        {
            // *** THE DATA-LOSS TEST. *** Highlights are the rectangles the whole app is built around.
            // Read back through the SHIPPED reader, so a writer agreeing with a test-only parser
            // could not pass.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            BoardData board = PublishExecutorTests.Board();
            board.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "100", Y = "200", Width = "40", Height = "20"
            });

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, board, [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished);

            ComponentHighlightEntry published = Assert.Single(BoardComponentHighlightStorage.LoadComponentHighlights(plan.BoardFilePath));

            Assert.Equal("U8", published.BoardLabel);
            Assert.Equal("100", published.X);
        }

        [Fact]
        public async Task A_KiCad_CALIBRATION_survives_the_publish()
        {
            // Calibrations are not part of BoardData at all, so they are passed separately and
            // would be the easiest thing to drop.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(store).ExecuteAsync(
                plan,
                PublishExecutorTests.Board(),
                [new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board.kicad_pcb", OffsetX = 1.5, ScaleX = 0.75, MirrorX = true }],
                submissionId: 1,
                PublishExecutorTests.Now);

            Assert.True(BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                plan.BoardFilePath, "Sheet 1",
                out string cadName, out double offsetX, out _,
                out double scaleX, out _, out bool mirrorX, out _));

            Assert.Equal("board.kicad_pcb", cadName);
            Assert.Equal(1.5, offsetX);
            Assert.Equal(0.75, scaleX);
            Assert.True(mirrorX);
        }

        // The renames the approval hands over are the board's whole list (owner decision, 2026-10-10).
        [Fact]
        public async Task The_schematic_renames_handed_over_are_the_board_files_list()
        {
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(PublishExecutorTests.StoreWithBoard()).ExecuteAsync(
                plan,
                PublishExecutorTests.Board(),
                [],
                [new SchematicRenameRecord("Top", "Sheet 1", "Images/sheet1.png", "2026-September-21")],
                submissionId: 1,
                PublishExecutorTests.Now,
                listing: null);

            SchematicRenameRecord rename = Assert.Single(BoardComponentHighlightStorage.LoadSchematicRenames(plan.BoardFilePath));
            Assert.Equal(("Top", "Sheet 1"), (rename.From, rename.To));
        }

        [Fact]
        public async Task A_HIGHLIGHT_ONLY_change_MOVES_the_published_content_hash()
        {
            // Moving a highlight changes no row and stores no file - only the board file's bytes
            // move, so without them folded in no client would re-download the board.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();

            BoardData first = PublishExecutorTests.Board();
            first.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "100", Y = "200", Width = "40", Height = "20" });

            await this.Executor(store).ExecuteAsync(this.Plan(PublishExecutorTests.Manifest()), first, [], submissionId: 1, PublishExecutorTests.Now);

            string before = store.PublishedBoards["Commodore/C64/250407"].ContentHash;

            BoardData second = PublishExecutorTests.Board();
            second.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "500", Y = "600", Width = "40", Height = "20" });

            await this.Executor(store).ExecuteAsync(this.Plan(PublishExecutorTests.Manifest()), second, [], submissionId: 2, PublishExecutorTests.Now);

            Assert.NotEqual(before, store.PublishedBoards["Commodore/C64/250407"].ContentHash);
        }

        [Fact]
        public async Task NULL_calibrations_are_REFUSED_rather_than_published_as_none()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await Assert.ThrowsAsync<ArgumentNullException>(() =>
                this.Executor(store).ExecuteAsync(plan, PublishExecutorTests.Board(), null!, submissionId: 1, PublishExecutorTests.Now));
        }

        // A field a later CRT added at the top of the board file survives a publish through this
        // server (DataTreeJson's rule).
        [Fact]
        public async Task A_field_this_server_does_not_know_is_kept_in_the_board_file()
        {
            var earlier = new BoardDocument
            {
                Unknown = new Dictionary<string, JsonElement> { ["laterField"] = JsonDocument.Parse("\"kept\"").RootElement.Clone() }
            };

            BoardFiles.Write(Path.Combine(this.thisBoardFolder, "board.json"), earlier);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(PublishExecutorTests.StoreWithBoard())
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.Equal("kept", BoardFiles.ReadDocument(plan.BoardFilePath)!.Unknown!["laterField"].GetString());
        }

        // ###########################################################################################
        // The approval reads the published board file once (code review, 2026-10-26: about five
        // parses under the publish lock) and hands it in - what is kept of it is that read, not
        // another. Shown here by a document that differs from the file on disk.
        // ###########################################################################################
        [Fact]
        public async Task The_board_file_handed_in_is_the_one_kept_from_not_another_read()
        {
            string boardFile = Path.Combine(this.thisBoardFolder, "board.json");
            BoardFiles.Write(boardFile, new BoardDocument
            {
                Unknown = new Dictionary<string, JsonElement> { ["laterField"] = JsonDocument.Parse("\"on disk\"").RootElement.Clone() }
            });

            var readOnce = new BoardDocument
            {
                Unknown = new Dictionary<string, JsonElement> { ["laterField"] = JsonDocument.Parse("\"read once\"").RootElement.Clone() }
            };

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(PublishExecutorTests.StoreWithBoard()).ExecuteAsync(
                plan,
                PublishExecutorTests.Board(),
                [],
                renames: [],
                submissionId: 1,
                PublishExecutorTests.Now,
                listing: null,
                publishedBoardFile: new PublishedBoardRead(boardFile, exists: true, readOnce));

            Assert.Equal("read once", BoardFiles.ReadDocument(plan.BoardFilePath)!.Unknown!["laterField"].GetString());
        }

        // And on a ROW (code review, 2026-10-25): the board file is built again from the rows, which
        // dropped a later CRT's field from every row - the edited one too, since it is the same
        // component (BoardDocumentUnknowns).
        [Fact]
        public async Task A_field_this_server_does_not_know_on_a_row_is_kept_in_the_board_file()
        {
            var earlier = new BoardDocument
            {
                Components =
                [
                    new BoardDocument.ComponentRow
                    {
                        BoardLabel = "U8",
                        FriendlyName = "PLA before the edit",
                        Unknown = new Dictionary<string, JsonElement> { ["pinCount"] = JsonDocument.Parse("28").RootElement.Clone() }
                    }
                ]
            };

            BoardFiles.Write(Path.Combine(this.thisBoardFolder, "board.json"), earlier);

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            await this.Executor(PublishExecutorTests.StoreWithBoard())
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            BoardDocument.ComponentRow written = Assert.Single(BoardFiles.ReadDocument(plan.BoardFilePath)!.Components);
            Assert.Equal("PLA", written.FriendlyName);
            Assert.Equal(28, written.Unknown!["pinCount"].GetInt32());
        }

        // ###########################################################################################
        // *** A BOARD FILE THAT CANNOT BE READ IN FULL IS NOT WRITTEN OVER (code review, 2026-10-25). ***
        // One of a newer format, or a hand edit gone wrong, read as "nothing there" and was replaced
        // by what this server knows - its unknown fields gone. The publish now stops, says why, and
        // records nothing.
        // ###########################################################################################
        [Fact]
        public async Task A_board_file_that_cannot_be_read_in_full_STOPS_the_publish_and_is_left_as_it_was()
        {
            string boardFile = Path.Combine(this.thisBoardFolder, "board.json");
            const string newer = "{\"formatVersion\": 99, \"hardwareName\": \"C64\"}\n";
            File.WriteAllText(boardFile, newer);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("could not be read in full", outcome.Failure);
            Assert.Equal(newer, File.ReadAllText(boardFile));
            Assert.Empty(store.PublishedBoards);
        }

        // ###########################################################################################
        // *** system.json IS RETIRED (owner decision, 2026-09-25). *** A publish writes none, and
        // removes one an earlier build left behind. What it recorded - revision and content hash -
        // still reaches the database, and the outcome still reports it.
        // ###########################################################################################
        [Fact]
        public async Task No_system_json_is_written_and_one_left_by_an_earlier_build_is_removed()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            string leftover = Path.Combine(this.thisBoardFolder, BoardDescriptorStore.FileName);
            await File.WriteAllTextAsync(leftover, "{}");

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);
            Assert.False(File.Exists(leftover));
            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now), outcome.Descriptor!.Revision);
            Assert.NotEmpty(outcome.Descriptor.ContentHash);
        }

        // -----------------------------------------------------------------------------------
        // The content hash accounts for the generated board file
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_rows_only_publish_still_moves_the_content_hash()
        {
            // *** THE SILENT FAILURE. *** A typo fix stores NO files, so the file list is identical
            // between two publishes. Without the board file folded into the hash, the promotion
            // would never offer the board that just changed.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome first = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            BoardData changed = PublishExecutorTests.Board();
            changed.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-02" };

            PublishOutcome second = await this.Executor(store)
                .ExecuteAsync(plan, changed, [], submissionId: 1, PublishExecutorTests.Now);

            Assert.NotEqual(first.Descriptor!.ContentHash, second.Descriptor!.ContentHash);

            // What makes the fold-in load-bearing: both publishes used the SAME plan, whose
            // descriptor covers the files only.
            Assert.NotEqual(plan.Descriptor.ContentHash, first.Descriptor.ContentHash);
            Assert.NotEqual(plan.Descriptor.ContentHash, second.Descriptor.ContentHash);
        }

        [Fact]
        public async Task The_reported_board_file_hash_matches_the_file_on_disk()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.Equal(PublishExecutorTests.HashOf(await File.ReadAllBytesAsync(plan.BoardFilePath)), outcome.WorkbookSha256);
        }

        // -----------------------------------------------------------------------------------
        // The database row
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_revision_and_content_hash_reach_the_boards_row()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            PublishedBoardRow row = store.PublishedBoards["Commodore/C64/250407"];

            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now), row.Revision);
            Assert.Equal(outcome.Descriptor!.ContentHash, row.ContentHash);
        }

        [Fact]
        public async Task The_submission_is_marked_merged()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();

            long id = await store.CreateAsync(
                new NewSubmission(
                    "Commodore/C64/250407", "Commodore", "C64", "250407",
                    null, "someone@example.com", "192.0.2.1", "hash", "r1", "Fixed a typo.",
                    1, [], PublishExecutorTests.Now, PublishExecutorTests.Now.AddHours(24)),
                CancellationToken.None);

            await this.Executor(store)
                .ExecuteAsync(this.Plan(PublishExecutorTests.Manifest()), PublishExecutorTests.Board(), [], id, PublishExecutorTests.Now);

            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);
        }

        [Fact]
        public async Task A_publish_with_no_submission_behind_it_is_allowed()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(this.Plan(PublishExecutorTests.Manifest()), PublishExecutorTests.Board(), [], submissionId: null, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished);
            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(PublishExecutorTests.Now), store.PublishedBoards["Commodore/C64/250407"].Revision);
        }

        // -----------------------------------------------------------------------------------
        // Failure leaves the tree recoverable
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_missing_blob_stops_the_publish_BEFORE_the_board_file_is_touched()
        {
            // *** THE ORDERING PROPERTY. *** Stored files go first because nothing points at them
            // until the board file lists them.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();

            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/sheet1.png", Sha256 = new string('a', 64), SizeBytes = 10 }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("Images/sheet1.png", outcome.Failure);
            Assert.False(File.Exists(plan.BoardFilePath));
            Assert.Empty(store.PublishedBoards);
        }

        // ###########################################################################################
        // *** A KiCad SCHEMATIC NAMING A SHEET OUTSIDE ITS "KiCad data" FOLDER IS NEVER WRITTEN
        // (security review, 2026-10-11). *** Finalise refuses one; the publish checks again, so a
        // submission that reached the queue before the rule existed cannot carry one into BETA.
        // ###########################################################################################
        [Fact]
        public async Task A_KiCad_sheet_naming_a_file_outside_its_KiCad_folder_stops_the_publish()
        {
            byte[] schematic = Encoding.UTF8.GetBytes(
                "(kicad_sch (version 20231120)\n" +
                "  (sheet (property \"Sheetname\" \"x\") (property \"Sheetfile\" \"//attacker.example/share/x.kicad_sch\"))\n" +
                ")\n");

            string hash = await this.PutBlobAsync(schematic);

            SubmissionManifest manifest = PublishExecutorTests.Manifest();
            manifest.Files.Add(new SubmissionFile { Path = "KiCad data/root.kicad_sch", Sha256 = hash, SizeBytes = schematic.LongLength });

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(manifest);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("KiCad data/root.kicad_sch", outcome.Failure);
            Assert.False(File.Exists(plan.Files[0].AbsolutePath));
            Assert.False(File.Exists(plan.BoardFilePath));
            Assert.Empty(store.PublishedBoards);
        }

        // ###########################################################################################
        // *** A DESTINATION THAT CANNOT BE WRITTEN IS A REPORTED FAILURE, NOT A 500 (owner
        // report, 2026-09-23). *** Provoked by making the destination a DIRECTORY: opening it as a
        // file throws UnauthorizedAccessException, the type the server hit, without manipulating
        // real permissions.
        // ###########################################################################################
        [Fact]
        public async Task A_destination_that_cannot_be_written_STOPS_the_publish_and_names_the_file()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            Directory.CreateDirectory(plan.Files[0].AbsolutePath);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("Images/sheet1.png", outcome.Failure);
            Assert.False(File.Exists(plan.BoardFilePath));
            Assert.Empty(store.PublishedBoards);
        }

        // ###########################################################################################
        // *** A FOLDER THE SERVICE MAY NOT WRITE STOPS THE PUBLISH BEFORE ANYTHING IS WRITTEN
        // (owner report, 2026-09-28). *** Files copied in by hand as root once refused the third
        // write of a publish. Every folder it writes into is asked first. (A folder that may not be
        // written cannot be made on every OS, so the executor's probe is told the answer.)
        // ###########################################################################################
        [Fact]
        public async Task A_board_folder_the_service_may_not_write_STOPS_the_publish_before_anything_is_written()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            var executor = new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance)
            {
                CanWriteFolderForTests = folder => folder != this.thisBoardFolder
            };

            PublishOutcome outcome = await executor
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("[Commodore/C64/250407]", outcome.Failure);
            Assert.Contains("nothing was changed", outcome.Failure);
            Assert.False(File.Exists(plan.Files[0].AbsolutePath));
            Assert.False(File.Exists(plan.BoardFilePath));
            Assert.Empty(store.PublishedBoards);
        }

        // A board file that cannot be written is reported and records nothing; the board as it was
        // stays, since the file is replaced whole.
        [Fact]
        public async Task A_board_file_that_cannot_be_written_is_REPORTED_not_thrown()
        {
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            Directory.CreateDirectory(plan.BoardFilePath);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("board file could not be written", outcome.Failure);
            Assert.Empty(store.PublishedBoards);
        }

        // ###########################################################################################
        // *** FILES COPIED IN BY HAND ARE REPLACED, NOT REFUSED (owner report, 2026-09-28). *** A
        // file copied into BETA as root may be replaced by the service - that is its folder's
        // permission - but not opened for writing. Made here as a file with no write permission for
        // its owner; not on Windows (no such mode) and not as root (which ignores it).
        // ###########################################################################################
        [Fact]
        public async Task A_board_file_that_may_not_be_opened_for_writing_is_still_replaced()
        {
            // The return is for the platform analyzer, which cannot see that Skip throws.
            if (OperatingSystem.IsWindows())
            {
                Assert.Skip("Unix permissions only.");
                return;
            }
            Assert.SkipWhen(Environment.UserName == "root", "root may write anything.");

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest());

            BoardFiles.Write(plan.BoardFilePath, new BoardDocument());
            File.SetUnixFileMode(plan.BoardFilePath, UnixFileMode.UserRead | UnixFileMode.GroupRead | UnixFileMode.OtherRead);

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(outcome.IsPublished, outcome.Failure);

            BoardData? published = await BoardDataReader.LoadAsync(plan.BoardFilePath, "hand-copied-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
        }

        [Fact]
        public async Task An_interrupted_publish_leaves_the_PREVIOUS_board_loadable()
        {
            // The recovery property in PublishExecutor's header: a failure leaves stored files
            // nothing points at - wasted space - rather than a broken board.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            (PublishPlanDetail first, _, _) = await this.PublishSheetAsync(store);

            PublishPlanDetail broken = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/new.png", Sha256 = new string('b', 64), SizeBytes = 10 }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(broken, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);

            BoardData? still = await BoardDataReader.LoadAsync(first.BoardFilePath, "recover-" + Guid.NewGuid().ToString("N"));

            Assert.NotNull(still);
            Assert.Equal("U8", Assert.Single(still!.Components).BoardLabel);
            Assert.Equal("Images/sheet1.png", Assert.Single(BoardFiles.ReadDocument(first.BoardFilePath)!.Files).Name);
        }

        [Fact]
        public async Task Re_running_the_same_publish_is_safe()
        {
            // The recovery for an interrupted publish: stored files are content-addressed, and the
            // board file is replaced again.
            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            (PublishPlanDetail plan, PublishOutcome first, _) = await this.PublishSheetAsync(store);

            PublishOutcome second = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.True(second.IsPublished, second.Failure);
            Assert.Equal(first.Descriptor!.ContentHash, second.Descriptor!.ContentHash);

            BoardData? published = await BoardDataReader.LoadAsync(plan.BoardFilePath, "rerun-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: everything is checked BEFORE anything is written.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_file_whose_bytes_are_not_its_type_stops_the_publish_before_ANY_file_is_written()
        {
            byte[] good = PublishExecutorTests.Png("GOOD");
            byte[] notAnImage = Encoding.UTF8.GetBytes("MZ not a picture at all");

            string goodHash = await this.PutBlobAsync(good);
            string badHash = await this.PutBlobAsync(notAnImage);

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/a.png", Sha256 = goodHash, SizeBytes = good.LongLength },
                new SubmissionFile { Path = "Images/z.png", Sha256 = badHash, SizeBytes = notAnImage.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.Contains("Images/z.png", outcome.Failure);
            Assert.All(plan.Files, file => Assert.False(File.Exists(file.AbsolutePath)));
            Assert.False(File.Exists(plan.BoardFilePath));
            Assert.Empty(store.PublishedBoards);
        }

        [Fact]
        public async Task A_blob_changed_in_the_store_stops_the_publish_before_anything_is_written()
        {
            byte[] image = PublishExecutorTests.Png("ORIGINAL");
            string hash = await this.PutBlobAsync(image);

            string stored = Path.Combine(this.thisBlobRoot, BlobStorePaths.BlobFolderName, hash[..2], hash[2..4], hash);
            await File.WriteAllBytesAsync(stored, PublishExecutorTests.Png("TAMPERED"));

            FakeSubmissionStore store = PublishExecutorTests.StoreWithBoard();
            PublishPlanDetail plan = this.Plan(PublishExecutorTests.Manifest(
                new SubmissionFile { Path = "Images/sheet1.png", Sha256 = hash, SizeBytes = image.LongLength }));

            PublishOutcome outcome = await this.Executor(store)
                .ExecuteAsync(plan, PublishExecutorTests.Board(), [], submissionId: 1, PublishExecutorTests.Now);

            Assert.False(outcome.IsPublished);
            Assert.False(File.Exists(plan.Files[0].AbsolutePath));
            Assert.False(File.Exists(plan.BoardFilePath));
        }
    }
}
