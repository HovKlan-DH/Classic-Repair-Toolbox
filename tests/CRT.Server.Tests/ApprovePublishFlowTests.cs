using System.Security.Cryptography;
using System.Text;
using CRT.Server.Configuration;
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
    // *** THIS IS THE ONE IRREVERSIBLE OPERATION IN THE WHOLE BOARD. *** Task 7 was struck, so no
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

            // *** THE TREE'S MAIN FILE (CRT 4.0, 2026-10-25), listing one other board. *** A NEW
            // board's publish adds its row to it (MasterListing, 2026-09-27); without it a new board
            // is refused, since it could never be listed. The other board's own file is not there,
            // so a removal preview reads this tree as incomplete - the tests of removals write a
            // complete one (UseRealTree).
            DataTreeBuilder.ListingMaster(this.thisDataTree, ApprovePublishFlowTests.OtherListedBoard);
        }

        // The board the fixture master lists - somewhere for a new board to be placed after.
        internal static readonly MasterListingRow OtherListedBoard =
            new("Amstrad CPC 664", "MC0005A", "Amstrad/CPC 664/MC0005A/board.json", string.Empty);

        // Where the fixture places its new board - after the one other board.
        internal static BoardPlacement Placement(string hardware = "Commodore 64", string board = "250407") =>
            new(hardware, board, "Has 6581 SID.", ApprovePublishFlowTests.OtherListedBoard.ExcelDataFile);

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

        private ApprovePublishFlow Flow(ISubmissionStore store, FakeAccountStore? accounts = null) =>
            new(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                accounts ?? new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance);

        // An account store in which the board HAS a maintainer - which made a shared-file change
        // need two approvals in the Excel tree; in the 4.0 tree one approval still publishes.
        private static FakeAccountStore AccountsWithAMaintainer()
        {
            // The pool row AND its account: the store drops a row whose account is missing, as
            // the real join does.
            var accounts = new FakeAccountStore();
            accounts.Accounts[7] = ApprovePublishFlowTests.AccountRecord(administrator: false);
            accounts.Maintainers.Add((ApprovePublishFlowTests.BoardId, 7));
            return accounts;
        }

        private const string BoardId = "Commodore/C64/250407";

        // The administrator, in every pool by definition.
        private static ReviewAccess Account() =>
            ReviewAccess.For(ApprovePublishFlowTests.AccountRecord(administrator: true));

        // A maintainer of exactly the boards named - and of nothing else.
        private static ReviewAccess MaintainerOf(params string[] boards) =>
            ReviewAccess.For(ApprovePublishFlowTests.AccountRecord(administrator: false), boards);

        private static Handlers.Accounts.AccountRecord AccountRecord(bool administrator) =>
            new(
                Id: 7,
                Email: administrator ? "admin@example.com" : "maintainer@example.com",
                NormalisedEmail: administrator ? "admin@example.com" : "maintainer@example.com",
                PasswordHash: "hash",
                DisplayName: administrator ? "Admin" : "Maintainer",
                IsVerified: true,
                IsAdministrator: administrator,
                IsLocked: false,
                CreatedUtc: ApprovePublishFlowTests.Now,
                LastLoginUtc: null);

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

            // Each file cited by a row - an uncited file is now refused (security review, 2026-09-25).
            foreach (SubmissionFile file in files)
                manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = file.Path });

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
            SubmissionManifest? manifest = null,
            bool placed = true,
            IReadOnlyList<SubmissionFile>? leftAsPublished = null)
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
                    manifest.BoardId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                    null, "someone@example.com", "192.0.2.1", "hash", "2026-August-21",
                    "Corrected U8.", 1, [.. manifest.Files], ApprovePublishFlowTests.Now,
                    ApprovePublishFlowTests.Now.AddHours(24),
                    LeftAsPublished: leftAsPublished),
                CancellationToken.None);

            await store.SavePayloadAsync(id, manifest, CancellationToken.None);
            await store.SetStateAsync(id, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

            // The fixture's board is NEW to the empty tree, and a new board must be placed in the
            // drop-down lists before it can be published (owner request, 2026-09-27).
            if (placed)
                await store.SetPlacementAsync(manifest.BoardId, ApprovePublishFlowTests.Placement(), 1, ApprovePublishFlowTests.Now, CancellationToken.None);

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
                this.thisDataTree, "Commodore", "C64", "250407", "board.json");

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
                this.thisDataTree, "Commodore", "C64", "250407", "board.json");

            ComponentHighlightEntry highlight = Assert.Single(
                BoardComponentHighlightStorage.LoadComponentHighlights(workbook));

            Assert.Equal("U8", highlight.BoardLabel);
            Assert.Equal("100", highlight.X);
        }

        [Fact]
        public async Task The_submission_is_recorded_as_MERGED_with_the_maintainer_who_approved_it()
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
                store.PublishedBoards["Commodore/C64/250407"].Revision);
        }

        [Fact]
        public async Task A_submitted_FILE_lands_in_the_published_tree()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                // The board's NAME for the file (CRT 4.0 tree, 2026-10-25) - the tree decides
                // where its bytes lie.
                ApprovePublishFlowTests.Manifest(new SubmissionFile
                {
                    Path = "Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            // Stored once under its checksum, and listed by the board under its name.
            Assert.Equal(image, await File.ReadAllBytesAsync(DataTreeBuilder.Full(
                this.thisDataTree, DataTreeLayout.StoredFilePath(hash, "Images/sheet1.png"))));
            Assert.Equal("Images/sheet1.png", this.PublishedListing().Single().Name);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "Images")));
        }

        // What the published board file lists.
        private IReadOnlyList<BoardDocument.StoredFileRow> PublishedListing() =>
            BoardFiles.ReadDocument(DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.BoardFile))!.Files;

        // ###########################################################################################
        // *** A RENAME THAT IS PUBLISHED IS LISTED IN THE BOARD'S .json (owner decision, 2026-10-10:
        // it "must also change for everyone updating to this data", in the board JSON file and not
        // in Excel). *** The second publish names the same image "PCB top side": the server sees
        // the image it published as "Top", lists the rename with the publish's revision date - the
        // list every CRT follows for its user's own worklogs and settings - and the next publish
        // keeps it. Nothing extra is sent with the submission.
        // ###########################################################################################
        [Fact]
        public async Task A_published_rename_is_listed_in_the_boards_json_and_kept_by_the_next_publish()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);
            var file = new SubmissionFile { Path = "Images/top.png", Sha256 = hash, SizeBytes = image.LongLength };

            SubmissionManifest Named(string name)
            {
                SubmissionManifest manifest = ApprovePublishFlowTests.Manifest(file);
                manifest.Rows.Schematics.Add(new BoardSchematicEntry { SchematicName = name, SchematicImageFile = file.Path });
                return manifest;
            }

            async Task<long> PendingWithAsync(FakeSubmissionStore store, SubmissionManifest manifest, string email)
            {
                long id = await store.CreateAsync(
                    new NewSubmission(
                        manifest.BoardId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                        null, email, "192.0.2.2", "hash-" + email, "2026-August-21",
                        "Renamed an image.", 1, [.. manifest.Files], ApprovePublishFlowTests.Now,
                        ApprovePublishFlowTests.Now.AddHours(24)),
                    CancellationToken.None);

                await store.SavePayloadAsync(id, manifest, CancellationToken.None);
                await store.SetStateAsync(id, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

                return id;
            }

            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync(Named("Top"));
            ApprovePublishFlow flow = this.Flow(store);

            ApproveOutcome published = await flow.ApproveAsync(first, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);
            Assert.True(published.IsPublished, published.Error);

            string workbook = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "board.json");
            Assert.Empty(BoardComponentHighlightStorage.LoadSchematicRenames(workbook));

            long second = await PendingWithAsync(store, Named("PCB top side"), "renamer@example.com");
            ApproveOutcome renamed = await flow.ApproveAsync(second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddDays(1), CancellationToken.None);
            Assert.True(renamed.IsPublished, renamed.Error);

            SchematicRenameRecord rename = Assert.Single(BoardComponentHighlightStorage.LoadSchematicRenames(workbook));
            Assert.Equal("Top", rename.From);
            Assert.Equal("PCB top side", rename.To);

            // The image it was renamed on - so a CRT can tell it from another image given the old
            // name later (owner request, 2026-10-10: "to be unambiguous").
            Assert.Equal(file.Path, rename.File);

            BoardData? board = BoardDataReader.ReadWorkbookUncached(workbook);
            Assert.Equal(board!.RevisionDate, rename.Date);

            long third = await PendingWithAsync(store, Named("PCB top side"), "later@example.com");
            ApproveOutcome later = await flow.ApproveAsync(third, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddDays(2), CancellationToken.None);
            Assert.True(later.IsPublished, later.Error);

            Assert.Equal([rename], BoardComponentHighlightStorage.LoadSchematicRenames(workbook));
        }

        // -----------------------------------------------------------------------------------
        // What it changed, for the board's History view (owner request, 2026-10-04: "summarize
        // each submission change in a textual form ... to get an idea, besides the sometimes vague
        // description from the contributor").
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** RECORDED AT THE PUBLISH - THE ONE MOMENT THE BOARD BEFORE STILL EXISTS. *** A new
        // board's first publish is the new board: its rows counted, its file new at its path.
        // ###########################################################################################
        [Fact]
        public async Task A_publish_records_what_it_changed_and_a_new_board_says_so()
        {
            byte[] image = PublishExecutorTests.Png("PNGDATA");
            string hash = await this.PutBlobAsync(image);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(new SubmissionFile
                {
                    Path = "Images/sheet1.png",
                    Sha256 = hash,
                    SizeBytes = image.LongLength
                }));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            SubmissionChanges changes = store.Changes[id];

            Assert.True(changes.IsNewBoard);
            Assert.Equal(1, changes.Sections.Single(section => section.Section == BoardWorkbookSchema.SheetComponents).AddedCount);
            Assert.Equal(["Images/sheet1.png"], changes.Files.Added);
            Assert.Empty(changes.Files.Replaced);
        }

        // ###########################################################################################
        // *** THE PUBLISH THAT CREATES A BOARD SAYS SO (owner request, 2026-10-13) *** - it is what
        // makes the board's creator its maintainer (NewBoardMaintainerFlow, from ReviewEndpoints).
        // The board's second publish does not.
        // ###########################################################################################
        [Fact]
        public async Task Only_the_publish_that_creates_the_board_says_it_created_it()
        {
            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync();
            ApprovePublishFlow flow = this.Flow(store);

            ApproveOutcome created = await flow.ApproveAsync(
                first, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            long second = await ApprovePublishFlowTests.AnotherPendingAsync(store);

            ApproveOutcome updated = await flow.ApproveAsync(
                second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddDays(1), CancellationToken.None);

            Assert.True(created.IsPublished, created.Error);
            Assert.True(created.CreatedBoard);
            Assert.True(updated.IsPublished, updated.Error);
            Assert.False(updated.CreatedBoard);
        }

        // ###########################################################################################
        // *** A BOARD THE STABLE TREE HOLDS IS NOT NEW (code review, 2026-10-14). *** "Created" was
        // "BETA holds no such board" - so a board placed in stable by hand, or whose BETA copy alone
        // was deleted, handed whoever sent its next change the maintainer role (or mailed an
        // invitation to a typed address). It is created only when NEITHER tree holds it.
        // ###########################################################################################
        [Fact]
        public async Task A_first_publish_to_BETA_of_a_board_the_stable_tree_holds_does_not_say_it_created_it()
        {
            string stable = Path.Combine(this.thisRoot, "app-data", "Data");
            string stableBoard = Path.Combine(stable, "Commodore", "C64", "250407");
            Directory.CreateDirectory(stableBoard);
            File.WriteAllText(Path.Combine(stableBoard, "board.json"), "the stable copy");

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            var flow = new ApprovePublishFlow(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance,
                options: new ServerOptions { ProductionDataTreeRoot = stable });

            ApproveOutcome outcome = await flow.ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.False(outcome.CreatedBoard);
        }

        // The stable tree named only by the setting that keeps BETA away from it (ProductionTreeRoot)
        // is read the same way.
        [Fact]
        public void The_stable_tree_is_looked_in_under_either_of_its_settings()
        {
            string stable = Path.Combine(this.thisRoot, "app-data", "Data");
            string stableBoard = Path.Combine(stable, "Commodore", "C64", "250407");
            Directory.CreateDirectory(stableBoard);
            File.WriteAllText(Path.Combine(stableBoard, "board.json"), "the stable copy");

            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();

            Assert.True(ApprovePublishFlow.StableHolds(new ServerOptions { ProductionDataTreeRoot = stable }, manifest));
            Assert.True(ApprovePublishFlow.StableHolds(new ServerOptions { ProductionTreeRoot = stable }, manifest));
            Assert.False(ApprovePublishFlow.StableHolds(new ServerOptions(), manifest));
            Assert.False(ApprovePublishFlow.StableHolds(null, manifest));
        }

        // ###########################################################################################
        // The next publish is compared with BETA as the first one left it: only what it changed, with
        // the field named - which is what lets the history say more than the contributor's words.
        // ###########################################################################################
        [Fact]
        public async Task A_later_publish_records_only_the_rows_and_fields_it_changed()
        {
            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync();
            ApprovePublishFlow flow = this.Flow(store);

            Assert.True((await flow.ApproveAsync(first, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None)).IsPublished);

            long second = await ApprovePublishFlowTests.AnotherPendingAsync(store);
            SubmissionManifest changed = (await store.LoadPayloadAsync(second, CancellationToken.None))!;
            // The fixture's U8 (Manifest), with another part number.
            changed.Rows.Components[0] = new ComponentEntry
            {
                BoardLabel = "U8",
                FriendlyName = "PLA",
                TechnicalNameOrValue = "906114-01",
                PartNumber = "906114-02"
            };
            await store.SavePayloadAsync(second, changed, CancellationToken.None);

            ApproveOutcome outcome = await flow.ApproveAsync(
                second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddDays(1), CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            SubmissionChanges changes = store.Changes[second];

            Assert.False(changes.IsNewBoard);

            SectionChanges components = Assert.Single(changes.Sections);
            ChangedRowFact row = Assert.Single(components.Changed);
            Assert.Equal("U8", row.Row);
            Assert.Equal([BoardWorkbookSchema.ColPartNumber], row.Fields);
        }

        // A publish that is refused records nothing - there is no change to describe.
        [Fact]
        public async Task A_refused_publish_records_no_changes()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.Empty(store.Changes);
        }

        // -----------------------------------------------------------------------------------
        // The refusals. Each one must leave the tree untouched.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_maintainer_of_ANOTHER_board_cannot_approve_and_NOTHING_is_written()
        {
            // *** THE STRUCTURAL GUARANTEE, asserted on the DISK rather than on a bool. *** A
            // maintainer's authority stops at their own boards; the only honest way to test that
            // is to check nothing changed.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id,
                ApprovePublishFlowTests.MaintainerOf("Commodore/C128/310378"),
                this.thisDataTree,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsForbidden);
            Assert.Contains(ApprovePublishFlowTests.BoardId, outcome.Error, StringComparison.Ordinal);

            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));

            // ...and the submission is untouched, so somebody entitled can still approve it.
            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);
            Assert.Equal(SubmissionState.Pending, record!.State);
        }

        [Fact]
        public async Task A_maintainer_of_THIS_board_publishes_it()
        {
            // The project owner's model: a maintainer assigned to a board publishes to it, with no
            // administrator in the loop.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id,
                ApprovePublishFlowTests.MaintainerOf(ApprovePublishFlowTests.BoardId),
                this.thisDataTree,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // -----------------------------------------------------------------------------------
        // The code review of 2026-10-13: what the maintainer saw, what they wrote, and what happens
        // after the tree is written.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** AN APPROVAL ON A VERSION THE APPROVER NEVER SAW IS REFUSED. *** Another maintainer's
        // amendment changes nothing on the queue row, so an Approve used to publish an amendment the
        // approver had not seen. The version the screen showed travels with the approval now.
        // ###########################################################################################
        [Fact]
        public async Task An_approval_made_on_an_older_amendment_version_is_refused_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            // Another maintainer saved a change after this one opened the submission at version 0.
            store.Amendments.Add((id, 1, ApprovePublishFlowTests.Manifest(), "Another Maintainer", ApprovePublishFlowTests.Now));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                new ApproveChecks(null, ExpectedVersion: 0), CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Equal(ReviewDecisionRules.AmendedSinceMessage, outcome.Error);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore", "C64")));
            Assert.Equal(SubmissionState.Pending, store.Submissions[id].State);
        }

        [Fact]
        public async Task An_approval_made_on_the_version_on_screen_publishes()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();
            store.Amendments.Add((id, 1, ApprovePublishFlowTests.Manifest(), "Another Maintainer", ApprovePublishFlowTests.Now));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                new ApproveChecks(null, ExpectedVersion: 1), CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        // ###########################################################################################
        // *** THE APPROVER'S COMMENT REACHES THE CONTRIBUTOR. *** The comment box kept its text across
        // Approve, and the merged mail has a paragraph for it - but the approval recorded no comment,
        // so it never printed. It is on the row now, which is what the mail reads.
        // ###########################################################################################
        [Fact]
        public async Task An_approvals_comment_is_recorded_on_the_submission_for_the_merged_mail()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                new ApproveChecks(null, Comment: "  Thanks - the PLA part number was overdue.  "), CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal("Thanks - the PLA part number was overdue.", store.Decisions[id].Comment);
            Assert.Equal("Thanks - the PLA part number was overdue.", store.Submissions[id].DecisionComment);
        }

        [Fact]
        public async Task An_over_long_approval_comment_is_refused_before_anything_is_written()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                new ApproveChecks(null, Comment: new string('x', ReviewDecisionRules.MaximumReasonLength + 1)), CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore", "C64")));
        }

        // ###########################################################################################
        // *** A FAILURE AFTER THE TREE IS WRITTEN IS STILL A PUBLISH. *** Recording the approval threw
        // out of the flow as a 500, so the endpoint's follow-up - the checksum manifest and the
        // contributor's mail - never ran for a board that was already in BETA.
        // ###########################################################################################
        [Fact]
        public async Task A_database_failure_after_the_tree_is_written_still_answers_published()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();
            store.FailAddApproval = true;

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);

            // The steps after the failing one still ran.
            Assert.Equal(SubmissionState.Merged, store.Decisions[id].State);
        }

        // ###########################################################################################
        // *** A CLIENT GIVING UP AFTER THE FIRST WRITE DOES NOT STOP THE PUBLISH. *** The request's
        // token was honoured to the end, so a closed laptop or a proxy timeout left the board written
        // and the submission still pending, or merged with nobody recorded as approving it.
        // ###########################################################################################
        [Fact]
        public async Task A_client_giving_up_after_the_tree_is_written_does_not_stop_the_publish()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();
            using var request = new CancellationTokenSource();

            // The workbook and its files are written by now; the client goes away.
            store.BeforeBoardPublished = request.Cancel;

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, request.Token);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal(SubmissionState.Merged, store.Submissions[id].State);
            Assert.Equal(7, store.Decisions[id].DecidedByAccountId);
            Assert.True(store.PublishedBoards.ContainsKey(ApprovePublishFlowTests.BoardId));
        }

        // ###########################################################################################
        // *** A FILE THE CONTRIBUTOR LEFT ALONE IS NOT REVERTED. *** S1 lists a file exactly as
        // published; S2 replaces it and is published; S1 is approved later. S1's copy of the OLD
        // bytes used to go back over S2's, because a manifest carries every file the board lists.
        // What S1 left as published at create is recorded, and the approval leaves it as it is
        // published now. In the CRT 4.0 tree (2026-10-25) a file is the board's NAME for it.
        // ###########################################################################################
        [Fact]
        public async Task A_file_a_submission_left_unchanged_keeps_the_bytes_a_later_publish_gave_it()
        {
            const string name = "pinout.txt";
            this.UseRealTree();

            // S1 is created while the file is as published, and lists it unchanged.
            byte[] asPublished = Encoding.UTF8.GetBytes("the pinout as published when S1 was made");
            DataTreeBuilder.BoardWithBytes(this.thisDataTree, ApprovePublishFlowTests.BoardId, (name, asPublished));
            string oldHash = await this.PutBlobAsync(asPublished);

            var cited = new SubmissionFile { Path = name, Sha256 = oldHash, SizeBytes = asPublished.Length };

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(cited),
                leftAsPublished: [cited]);

            // S2 replaces the file and is published.
            DataTreeBuilder.Board(this.thisDataTree, ApprovePublishFlowTests.BoardId, (name, "the corrected pinout S2 published"));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal("the corrected pinout S2 published", this.PublishedContentOf(name));
        }

        // A file the contributor CHANGED is no such file: its bytes are published as before.
        [Fact]
        public async Task A_file_the_submission_changed_is_still_published_whatever_the_record_says_of_others()
        {
            const string name = "pinout.txt";
            this.UseRealTree();
            DataTreeBuilder.Board(this.thisDataTree, ApprovePublishFlowTests.BoardId, (name, "the pinout as published"));

            byte[] better = Encoding.UTF8.GetBytes("the contributor's better pinout");
            string hash = await this.PutBlobAsync(better);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(new SubmissionFile { Path = name, Sha256 = hash, SizeBytes = better.Length }),
                leftAsPublished: []);

            // The bytes it replaces are listed by nothing afterwards, so they go - shown first.
            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                [DataTreeBuilder.Stored("the pinout as published", name)], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal("the contributor's better pinout", this.PublishedContentOf(name));
        }

        // What the published board file lists under `name`, read through its stored bytes.
        private string? PublishedContentOf(string name)
        {
            BoardDocument? document = BoardFiles.ReadDocument(DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.BoardFile));
            BoardDocument.StoredFileRow? row = document?.Files.FirstOrDefault(file => string.Equals(file.Name, name, StringComparison.Ordinal));

            return row is null
                ? null
                : File.ReadAllText(DataTreeBuilder.Full(this.thisDataTree, DataTreeLayout.StoredFilePath(row.Sha256, row.Name)));
        }

        // ###########################################################################################
        // *** AN APPROVAL IS RECORDED ONLY FOR SOMETHING THAT CAN BE PUBLISHED (code review,
        // 2026-09-25). *** The first of two approvals used to be recorded - and the administrator
        // mailed that their approval was needed - before the payload was read or the plan built,
        // so a submission the plan refuses was only refused at the SECOND approval, in the hands of
        // the person who had just been asked to give it.
        //
        // In the CRT 4.0 tree (2026-10-25) nothing is shared, so no submission needs two approvals
        // any more; the shared-file approval cases this class had went with the shared files
        // (ApprovalRules keeps its own tests). This one stays: a submission flagged under the old
        // rule is still refused by its plan before anything of it is recorded.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_the_plan_refuses_is_refused_BEFORE_its_first_approval_is_recorded()
        {
            // A file no row cites: the publish plan refuses it.
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest(new SubmissionFile
            {
                Path = "notes.txt", Sha256 = new string('b', 64), SizeBytes = 3
            });
            manifest.Rows.BoardLocalFiles.Clear();

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(manifest);
            store.Submissions[id] = store.Submissions[id] with { TouchesSharedFiles = true };

            ApproveOutcome outcome = await this.Flow(store, ApprovePublishFlowTests.AccountsWithAMaintainer()).ApproveAsync(
                id, ApprovePublishFlowTests.MaintainerOf(ApprovePublishFlowTests.BoardId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(outcome.IsAwaitingApproval);
            Assert.Contains("cannot be published", outcome.Error, StringComparison.Ordinal);
            Assert.Empty(await store.GetApprovalsAsync(id));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
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
            // deleting the board's entire contents, irreversibly.
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

        // ###########################################################################################
        // *** A BOARD IS PUBLISHED AS ITS BOARD FILE, AND NOTHING EXCEL (CRT 4.0, 2026-10-25). *** The
        // Excel generations are frozen for older CRTs in trees of their own; this tree holds the
        // board file and the bytes it lists, nothing else in the board's folder.
        // ###########################################################################################
        [Fact]
        public async Task A_NEW_board_is_published_as_its_board_file_alone()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            string folder = Path.Combine(this.thisDataTree, "Commodore", "C64", "250407");

            Assert.Equal(["board.json"], Directory.GetFiles(folder, "*", SearchOption.AllDirectories).Select(Path.GetFileName));
        }

        // ###########################################################################################
        // A tree with no main file - a fresh or half-set-up one - cannot list a new board, so the
        // publish is refused outright rather than writing a board nobody could choose.
        // ###########################################################################################
        [Fact]
        public async Task A_tree_with_NO_main_file_refuses_a_new_board_and_writes_nothing()
        {
            File.Delete(Path.Combine(this.thisDataTree, DataTreeLayout.MainFileName));

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, DataTreeLayout.FilesFolder)));
        }

        // ###########################################################################################
        // An EXISTING board's file is replaced with the submission's rows - and a field a newer CRT
        // wrote into it that this server does not know is KEPT (DataTreeJson: a field is never
        // dropped by a reader that does not know it).
        // ###########################################################################################
        [Fact]
        public async Task An_EXISTING_boards_file_is_replaced_and_keeps_the_fields_this_server_does_not_know()
        {
            DataTreeBuilder.Board(
                this.thisDataTree,
                ApprovePublishFlowTests.BoardId,
                new BoardData { RevisionDate = "2026-August-21", Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" }] },
                []);

            string boardFile = DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.BoardFile);
            BoardDocument existing = BoardFiles.ReadDocument(boardFile)!;
            existing.Unknown = new Dictionary<string, System.Text.Json.JsonElement>
            {
                ["aFieldFromANewerCrt"] = System.Text.Json.JsonDocument.Parse("7").RootElement.Clone()
            };
            BoardFiles.Write(boardFile, existing);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(boardFile, "existing-" + Guid.NewGuid().ToString("N"));
            Assert.Equal("U8", Assert.Single(published!.Components).BoardLabel);

            BoardDocument written = BoardFiles.ReadDocument(boardFile)!;
            Assert.Equal(7, written.Unknown!["aFieldFromANewerCrt"].GetInt32());
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
                this.thisDataTree, "Commodore", "C64", "250407", "board.json");

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
                BoardId = "Commodore/C64/250407",
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
                Path.Combine(this.thisDataTree, "Commodore", "C64", "250407", "board.json"),
                "approve2-" + Guid.NewGuid().ToString("N"));

            Assert.Equal("U9", Assert.Single(published!.Components).BoardLabel);
        }

        // ###########################################################################################
        // FILES THE BOARD NO LONGER USES ARE REMOVED - AND ONLY THOSE THE MAINTAINER WAS SHOWN
        // (owner decisions, 2026-09-25: "there must be no orphan files", and the list "must be
        // visible BEFORE the maintainer/admin approves it").
        //
        // In the CRT 4.0 tree (2026-10-25) what goes is a STORED file ("Files/<ab>/<sha256>.<ext>")
        // no board lists any more; the list matched is of those, and the maintainer reads each by
        // the name the board knew it by (FileRemovalPreview.Names, ApproveOutcome.RemovedFiles).
        //
        // These need a COMPLETE tree: the constructor's main file lists a board whose file is not
        // there, and a tree that cannot be read completely removes nothing.
        // ###########################################################################################
        private void UseRealTree(params string[] alsoListed) =>
            DataTreeBuilder.Master(this.thisDataTree, [ApprovePublishFlowTests.BoardId, .. alsoListed]);

        private const string OldManualName = "old-manual.pdf";
        private const string OldManualContent = "the old manual";

        private static readonly string OldManual = DataTreeBuilder.Stored(ApprovePublishFlowTests.OldManualContent, ApprovePublishFlowTests.OldManualName);

        // The published board lists a manual; the submission (rows only) no longer does.
        private async Task<(FakeSubmissionStore Store, long Id, SubmissionManifest Manifest)> DroppingTheManualAsync(params (string Name, string Content)[] alsoListed)
        {
            this.UseRealTree();
            DataTreeBuilder.Board(
                this.thisDataTree,
                ApprovePublishFlowTests.BoardId,
                [(ApprovePublishFlowTests.OldManualName, ApprovePublishFlowTests.OldManualContent), .. alsoListed]);

            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(manifest);
            return (store, id, manifest);
        }

        [Fact]
        public async Task The_maintainer_is_shown_what_a_publish_would_remove_before_approving()
        {
            (_, _, SubmissionManifest manifest) = await this.DroppingTheManualAsync();

            BoardData? published = await new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance)
                .TryReadAsync(this.thisDataTree, manifest);

            FileRemovalPreview preview = ApprovePublishFlow.PreviewRemovals(this.thisDataTree, manifest, published, ApprovePublishFlowTests.Now);

            Assert.False(preview.IsBlocked, preview.BlockedBecause);
            Assert.Equal([ApprovePublishFlowTests.OldManual], preview.Files);
            Assert.Equal(ApprovePublishFlowTests.OldManualName, preview.NameOf(ApprovePublishFlowTests.OldManual));
        }

        [Fact]
        public async Task Approving_with_the_list_shown_publishes_and_removes_exactly_that_list()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                [ApprovePublishFlowTests.OldManual], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ApprovePublishFlowTests.OldManualName], outcome.RemovedFiles);
            Assert.False(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.OldManual)));
            // The board's history says it by name too.
            Assert.Equal([ApprovePublishFlowTests.OldManualName], store.Changes[id].Files.Removed);
        }

        // What is removed is what was on screen. A maintainer shown a different list - here an empty
        // one, because the list changed since they opened it - is refused, and NOTHING is written.
        [Fact]
        public async Task Approving_with_a_different_list_is_refused_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();
            string boardFile = DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.BoardFile);
            byte[] before = File.ReadAllBytes(boardFile);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: [], CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Equal(ApprovePublishFlow.RemovalsChangedMessage, outcome.Error);
            Assert.Equal(before, File.ReadAllBytes(boardFile));
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.OldManual)));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // ###########################################################################################
        // *** NO LIST AT ALL IS AN OLDER MAINTAINER APPLICATION, NOT A CHANGED ONE (code review,
        // 2026-09-25). *** The current one always sends its list, empty or not; one built before
        // the list existed sends none. It used to be told the list had "changed since you opened
        // it" - false, and reopening changed nothing, so it could never publish and never learn
        // why. It is told to update instead - still refused, still nothing written.
        // ###########################################################################################
        [Fact]
        public async Task A_review_application_that_sends_no_list_is_told_to_update_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();
            string boardFile = DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.BoardFile);
            byte[] before = File.ReadAllBytes(boardFile);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: null, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(outcome.IsConflict);
            Assert.Equal(ApprovePublishFlow.RemovalsNotSentMessage, outcome.Error);

            // It tells the maintainer to update CRT - the separate CRT Maintainer application it
            // used to name was folded into CRT on 2026-09-29.
            Assert.Contains("Update CRT,", outcome.Error, StringComparison.Ordinal);
            Assert.DoesNotContain("CRT Maintainer", outcome.Error, StringComparison.Ordinal);
            Assert.Equal(before, File.ReadAllBytes(boardFile));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // A submission that removes NOTHING still publishes from an older application: there was no
        // list to show, so none was missed.
        [Fact]
        public async Task A_review_application_that_sends_no_list_still_publishes_what_removes_nothing()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: null, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        // ###########################################################################################
        // *** BYTES ANOTHER BOARD LISTS ARE NEITHER LISTED NOR REMOVED (CRT 4.0 tree, 2026-10-25). ***
        // A file's bytes are stored once for every board naming them: the manual this board stops
        // listing that another board lists too is still that board's manual.
        // ###########################################################################################
        [Fact]
        public async Task Bytes_another_board_still_lists_are_neither_listed_nor_removed()
        {
            const string other = "Commodore/C128/310378";

            (FakeSubmissionStore store, long id, SubmissionManifest manifest) = await this.DroppingTheManualAsync(("manual.pdf", "the shared manual"));
            this.UseRealTree(other);
            DataTreeBuilder.Board(this.thisDataTree, other, ("Manuals/the same manual.pdf", "the shared manual"));

            BoardData? published = await new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance)
                .TryReadAsync(this.thisDataTree, manifest);

            Assert.Equal(
                [ApprovePublishFlowTests.OldManual],
                ApprovePublishFlow.PreviewRemovals(this.thisDataTree, manifest, published, ApprovePublishFlowTests.Now).Files);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                [ApprovePublishFlowTests.OldManual], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ApprovePublishFlowTests.OldManualName], outcome.RemovedFiles);
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.Stored("the shared manual", "manual.pdf"))));
        }

        // ###########################################################################################
        // A submission queued under an older rule as touching shared files (TouchesSharedFiles). The
        // approval asks what it would do NOW - and in the CRT 4.0 tree nothing is shared - so it
        // needs one approval, and the flag comes down, so the queue stops saying it needs the
        // administrator.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_flagged_as_touching_shared_files_needs_one_approval_and_the_flag_comes_down()
        {
            byte[] bytes = Encoding.UTF8.GetBytes("a new manual");
            string hash = await this.PutBlobAsync(bytes);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(new SubmissionFile { Path = "new-manual.txt", Sha256 = hash, SizeBytes = bytes.Length }));
            store.Submissions[id] = store.Submissions[id] with { TouchesSharedFiles = true };

            ApproveOutcome outcome = await this.Flow(store, ApprovePublishFlowTests.AccountsWithAMaintainer()).ApproveAsync(
                id, ApprovePublishFlowTests.MaintainerOf(ApprovePublishFlowTests.BoardId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.False(store.Submissions[id].TouchesSharedFiles);
            Assert.Equal("a new manual", File.ReadAllText(DataTreeBuilder.Full(this.thisDataTree, DataTreeLayout.StoredFilePath(hash, "new-manual.txt"))));
        }

        // A tree that cannot be read completely removes nothing - and says so on the list, which
        // is then empty, so the approval goes through with nothing removed.
        [Fact]
        public async Task An_unreadable_tree_removes_nothing_and_the_publish_still_goes_through()
        {
            (FakeSubmissionStore store, long id, SubmissionManifest manifest) = await this.DroppingTheManualAsync();

            // The main file lists a board whose board file is not there.
            this.UseRealTree("Amstrad/CPC/464");

            BoardData? published = await new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance)
                .TryReadAsync(this.thisDataTree, manifest);
            FileRemovalPreview preview = ApprovePublishFlow.PreviewRemovals(this.thisDataTree, manifest, published, ApprovePublishFlowTests.Now);

            Assert.True(preview.IsBlocked);
            Assert.Empty(preview.Files);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: [], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Empty(outcome.RemovedFiles);
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.OldManual)));
        }

        // -----------------------------------------------------------------------------------
        // The revision date, which the SERVER decides
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** A NEW BOARD CAN BE PUBLISHED AT ALL (reported by the project owner, 2026-09-26:
        // "This submission cannot be published: A publish must carry a revision"). ***
        //
        // A new board's workbook is seeded with no revision date (DraftSeeder.CreateNewBoard - it
        // has no published board to inherit one from, and CRT never asks for one), PublishMerge's
        // fallback to the published board finds none either, and PublishPlan then refused
        // `publish.no-revision` - at the approval, the one irreversible step. It failed closed, so
        // nothing was corrupted; it simply could not be published.
        //
        // The server stamps the revision now, so a submission carrying none is fine by
        // construction. This test would fail against the version that read the board's own date.
        // ###########################################################################################
        [Fact]
        public void A_NEW_board_that_carries_no_revision_date_can_still_be_planned()
        {
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();

            // Exactly what a seeded new board sends: no revision date at all.
            manifest.Rows.RevisionDate = string.Empty;

            PublishPlanResult result = ApprovePublishFlow.BuildPlan(
                this.thisDataTree,
                manifest,
                ApprovePublishFlowTests.Now);

            Assert.True(
                result.IsPlanned,
                string.Join(" ", result.Problems.Select(problem => problem.Message)));

            Assert.Equal(
                BoardWorkbookStyle.FormatRevisionDate(ApprovePublishFlowTests.Now),
                result.Plan!.Descriptor.Revision);
        }

        // ###########################################################################################
        // *** THE SERVER'S DATE WINS OVER WHATEVER THE SUBMISSION CARRIES (owner confirmation,
        // 2026-09-26: "server always wins, and what is typed by user is not important"). ***
        //
        // The workbook was already stamped with the publish date, but the PLAN was built from the
        // submitted one - and the plan's descriptor is what reaches `boards.current_revision`. So
        // the board and the database row disagreed, and that row is the base a contributor's next
        // draft is diffed against.
        // ###########################################################################################
        [Fact]
        public void The_submissions_own_revision_date_is_ignored_in_favour_of_the_publish_date()
        {
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();

            // A contributor who started their draft long before it was reviewed.
            manifest.Rows.RevisionDate = "2026-May-12";

            PublishPlanResult result = ApprovePublishFlow.BuildPlan(
                this.thisDataTree,
                manifest,
                ApprovePublishFlowTests.Now);

            Assert.True(result.IsPlanned);
            Assert.Equal(
                BoardWorkbookStyle.FormatRevisionDate(ApprovePublishFlowTests.Now),
                result.Plan!.Descriptor.Revision);
        }

        // -----------------------------------------------------------------------------------
        // A NEW board's place in the drop-down lists (owner request, 2026-09-27): "The maintainer
        // should order the new system, so it becomes visible in the right location for the
        // drop-down lists. This must be done before it can be pushed to BETA."
        // -----------------------------------------------------------------------------------

        private const string PublishedWorkbook = "Commodore/C64/250407/board.json";

        [Fact]
        public async Task A_NEW_board_nobody_has_PLACED_is_refused_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);
            byte[] masterBefore = File.ReadAllBytes(Path.Combine(this.thisDataTree, DataTreeLayout.MainFileName));

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Equal(BoardListingRules.NotPlacedMessage, outcome.Error);

            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
            Assert.Equal(masterBefore, File.ReadAllBytes(Path.Combine(this.thisDataTree, DataTreeLayout.MainFileName)));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);

            // And no approval was recorded for something that could not be published.
            Assert.Empty(await store.GetApprovalsAsync(id, CancellationToken.None));
        }

        [Fact]
        public async Task A_PLACED_new_board_is_added_to_the_drop_down_lists_where_it_was_placed()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            Assert.Equal(
                [
                    ApprovePublishFlowTests.OtherListedBoard,
                    new MasterListingRow("Commodore 64", "250407", ApprovePublishFlowTests.PublishedWorkbook, "Has 6581 SID."),
                ],
                DataTreeBuilder.ListedIn(this.thisDataTree));

            // The row names the workbook the publish actually wrote.
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.PublishedWorkbook)));
        }

        // Every board the list already carries publishes exactly as before - no placement asked,
        // and the file left as it was.
        [Fact]
        public async Task A_board_the_list_already_carries_needs_no_placement_and_the_list_is_untouched()
        {
            MasterListingRow listed = new("Commodore 64", "250407 (long board)", ApprovePublishFlowTests.PublishedWorkbook, string.Empty);
            DataTreeBuilder.ListingMaster(this.thisDataTree, ApprovePublishFlowTests.OtherListedBoard, listed);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ApprovePublishFlowTests.OtherListedBoard, listed], DataTreeBuilder.ListedIn(this.thisDataTree));
        }

        // ###########################################################################################
        // *** THE "# Hardware:" / "# Board:" CAPTION SURVIVES AN APPROVAL (owner report, 2026-09-28).
        // *** The rows never carried it and the publish left it out, so every sheet of the C128
        // workbook approved into BETA had lost the two lines production's copy has. The board being
        // replaced gives it - not the drop-down's names, which word it differently.
        // ###########################################################################################
        [Fact]
        public async Task An_existing_board_keeps_its_own_caption()
        {
            string workbook = DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.PublishedWorkbook);
            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            BoardWorkbookWriter.Write(workbook, new BoardData { RevisionDate = "2026-August-21", HardwareName = "Commodore 64 and 64C", BoardName = "250407" });

            DataTreeBuilder.ListingMaster(
                this.thisDataTree,
                ApprovePublishFlowTests.OtherListedBoard,
                new MasterListingRow("Commodore 64", "250407 (long board)", ApprovePublishFlowTests.PublishedWorkbook, string.Empty));

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(workbook, "caption-" + Guid.NewGuid().ToString("N"));

            Assert.Equal(("Commodore 64 and 64C", "250407"), (published!.HardwareName, published.BoardName));
        }

        // A new board has no board to take it from: it is captioned with the names it was placed
        // under - the ones its contributor typed, so what the draft already showed.
        [Fact]
        public async Task A_new_board_is_captioned_with_the_names_it_was_placed_under()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(
                DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.PublishedWorkbook), "caption-" + Guid.NewGuid().ToString("N"));

            Assert.Equal(("Commodore 64", "250407"), (published!.HardwareName, published.BoardName));
        }

        // A board published before this was fixed has no caption left: it gets the drop-down's
        // names rather than staying blank for good.
        [Fact]
        public async Task A_board_that_lost_its_caption_gets_the_names_it_is_listed_under()
        {
            string workbook = DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.PublishedWorkbook);
            Directory.CreateDirectory(Path.GetDirectoryName(workbook)!);
            BoardWorkbookWriter.Write(workbook, new BoardData { RevisionDate = "2026-August-21" });

            DataTreeBuilder.ListingMaster(
                this.thisDataTree,
                ApprovePublishFlowTests.OtherListedBoard,
                new MasterListingRow("Commodore 64", "250407 (long board)", ApprovePublishFlowTests.PublishedWorkbook, string.Empty));

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);

            BoardData? published = await BoardDataReader.LoadAsync(workbook, "caption-" + Guid.NewGuid().ToString("N"));

            Assert.Equal(("Commodore 64", "250407 (long board)"), (published!.HardwareName, published.BoardName));
        }

        // ###########################################################################################
        // A board ALREADY IN THE TREE is published as it always was when the list cannot be read -
        // the file was never part of publishing one, and must not start blocking it. Only a board
        // new to the tree needs the file, since it could never be listed without it.
        // ###########################################################################################
        [Fact]
        public async Task With_no_readable_list_a_board_already_in_the_tree_still_publishes()
        {
            File.WriteAllText(Path.Combine(this.thisDataTree, DataTreeLayout.MainFileName), "not a workbook");
            DataTreeBuilder.Board(this.thisDataTree, ApprovePublishFlowTests.BoardId);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        [Fact]
        public async Task With_no_readable_list_a_NEW_board_is_refused()
        {
            File.WriteAllText(Path.Combine(this.thisDataTree, DataTreeLayout.MainFileName), "not a workbook");

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.Contains("could not be read", outcome.Error);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        // Another board listed under the same names since this one was placed (the lists can change
        // between the placement and the approval): refused before a byte is written.
        [Fact]
        public async Task A_board_whose_names_were_taken_since_it_was_placed_is_refused_before_anything_is_written()
        {
            MasterListingRow sameNames = new("Commodore 64", "250407", "Commodore/C64/326298/board.json", string.Empty);
            DataTreeBuilder.ListingMaster(this.thisDataTree, ApprovePublishFlowTests.OtherListedBoard, sameNames);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.Contains("[Commodore 64] / [250407] is already in the drop-down lists", outcome.Error);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore", "C64", "250407")));
            Assert.Equal([ApprovePublishFlowTests.OtherListedBoard, sameNames], DataTreeBuilder.ListedIn(this.thisDataTree));
        }

        // Placed after a row that has since left the list: refused before a byte is written, rather
        // than put somewhere nobody chose.
        [Fact]
        public async Task A_board_placed_after_a_row_that_has_LEFT_the_list_is_refused_before_anything_is_written()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(placed: false);

            await store.SetPlacementAsync(
                ApprovePublishFlowTests.BoardId,
                ApprovePublishFlowTests.Placement() with { AfterExcelDataFile = "Commodore/VIC-20/999999/board.json" },
                1,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.Contains("no longer in the drop-down lists", outcome.Error);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        // -----------------------------------------------------------------------------------
        // One submission in BETA per board (owner decision, 2026-09-27)
        // -----------------------------------------------------------------------------------

        private static readonly ServerOptions WithProduction = new()
        {
            ProductionDataTreeRoot = "production",
            ProductionManifestPath = "production/dataChecksums.json",
            ProductionPublicDataBaseUrl = "https://example.com/app-data/Data"
        };

        private ApprovePublishFlow FlowWithProduction(ISubmissionStore store) =>
            new(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance,
                publishLock: null,
                options: ApprovePublishFlowTests.WithProduction);

        // A second pending submission of the fixture's board, in the same store.
        private static async Task<long> AnotherPendingAsync(FakeSubmissionStore store)
        {
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();

            long id = await store.CreateAsync(
                new NewSubmission(
                    manifest.BoardId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                    null, "other@example.com", "192.0.2.2", "hash2", "2026-August-21",
                    "Another fix.", 1, [.. manifest.Files], ApprovePublishFlowTests.Now,
                    ApprovePublishFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SavePayloadAsync(id, manifest, CancellationToken.None);
            await store.SetStateAsync(id, SubmissionState.Pending, ApprovePublishFlowTests.Now, CancellationToken.None);

            return id;
        }

        // ###########################################################################################
        // *** WHILE A BOARD WAITS IN BETA, NO OTHER SUBMISSION OF IT IS APPROVED. *** A push-back
        // returns everything merged since the last promotion - it cannot pick one contributor's work
        // out - so two in BETA at once could only be pushed back together. Refused with the shared
        // sentence, nothing written; once production has caught up, the same approval goes through.
        // ###########################################################################################
        [Fact]
        public async Task A_second_submission_is_not_approved_while_the_first_waits_in_BETA()
        {
            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync();
            ApprovePublishFlow flow = this.FlowWithProduction(store);

            ApproveOutcome published = await flow.ApproveAsync(
                first, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);
            Assert.True(published.IsPublished, published.Error);

            long second = await ApprovePublishFlowTests.AnotherPendingAsync(store);

            ApproveOutcome refused = await flow.ApproveAsync(
                second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.False(refused.IsPublished);
            Assert.True(refused.IsConflict);
            Assert.Equal(OneSubmissionInBeta.BusyMessage(ApprovePublishFlowTests.BoardId), refused.Error);
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(second, CancellationToken.None))!.State);
            Assert.Empty(await store.GetApprovalsAsync(second, CancellationToken.None));

            // Production catches up (Beta > Prod): the board no longer waits, and the approval goes through.
            store.ProductionBoards[ApprovePublishFlowTests.BoardId] = store.PublishedBoards[ApprovePublishFlowTests.BoardId];

            ApproveOutcome afterPromotion = await flow.ApproveAsync(
                second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddMinutes(2), CancellationToken.None);

            Assert.True(afterPromotion.IsPublished, afterPromotion.Error);
        }

        // With no production to publish to, nothing ever leaves BETA - the rule would close every
        // board after its first approval, so it is off.
        [Fact]
        public async Task Without_production_publishing_a_second_submission_is_approved_as_before()
        {
            (FakeSubmissionStore store, long first) = await ApprovePublishFlowTests.PendingAsync();
            ApprovePublishFlow flow = this.Flow(store);

            Assert.True((await flow.ApproveAsync(first, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None)).IsPublished);

            long second = await ApprovePublishFlowTests.AnotherPendingAsync(store);

            ApproveOutcome outcome = await flow.ApproveAsync(
                second, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        // ###########################################################################################
        // The review screen's shortcut (code review, 2026-10-25 - looked up by name now, not every
        // published file against every submitted one): nothing can go only while EVERY published file
        // is still sent with the same bytes. Names and checksums ignore capitals, as they always did;
        // a name sent twice counts with either checksum.
        // ###########################################################################################
        [Fact]
        public void Nothing_can_go_only_while_every_published_file_is_still_sent_with_the_same_bytes()
        {
            string a = new('a', 64);
            string b = new('b', 64);

            var published = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase)
            {
                ["Images/top.png"] = new() { Name = "Images/top.png", Sha256 = a },
                ["KiCad data/board.kicad_pcb"] = new() { Name = "KiCad data/board.kicad_pcb", Sha256 = b },
            };

            static SubmissionFile Sent(string path, string sha) => new() { Path = path, Sha256 = sha };

            Assert.True(ApprovePublishFlow.EveryPublishedFileStillSent(published,
                [Sent("images/TOP.png", a.ToUpperInvariant()), Sent("KiCad data/board.kicad_pcb", b), Sent("Images/new.png", a)]));

            // One changed, one missing.
            Assert.False(ApprovePublishFlow.EveryPublishedFileStillSent(published,
                [Sent("Images/top.png", b), Sent("KiCad data/board.kicad_pcb", b)]));
            Assert.False(ApprovePublishFlow.EveryPublishedFileStillSent(published, [Sent("Images/top.png", a)]));

            // A name sent twice, once with the published bytes.
            Assert.True(ApprovePublishFlow.EveryPublishedFileStillSent(published,
                [Sent("Images/top.png", b), Sent("Images/top.png", a), Sent("KiCad data/board.kicad_pcb", b)]));

            Assert.True(ApprovePublishFlow.EveryPublishedFileStillSent(new Dictionary<string, BoardDocument.StoredFileRow>(), []));
        }
    }
}
