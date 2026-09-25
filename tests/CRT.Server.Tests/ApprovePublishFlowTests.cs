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

        private ApprovePublishFlow Flow(ISubmissionStore store, FakeAccountStore? accounts = null) =>
            new(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance),
                store,
                accounts ?? new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance);

        // An account store in which the board HAS a reviewer - which is what makes a shared-file
        // change need two approvals rather than the administrator's alone.
        private static FakeAccountStore AccountsWithAReviewer()
        {
            // The pool row AND its account: the store drops a row whose account is missing, as
            // the real join does.
            var accounts = new FakeAccountStore();
            accounts.Accounts[7] = ApprovePublishFlowTests.AccountRecord(administrator: false);
            accounts.Reviewers.Add((ApprovePublishFlowTests.SystemId, 7));
            return accounts;
        }

        private const string SystemId = "Commodore/C64/250407";

        // The administrator, in every pool by definition.
        private static ReviewAccess Account() =>
            ReviewAccess.For(ApprovePublishFlowTests.AccountRecord(administrator: true));

        // A reviewer of exactly the systems named - and of nothing else.
        private static ReviewAccess ReviewerOf(params string[] systems) =>
            ReviewAccess.For(ApprovePublishFlowTests.AccountRecord(administrator: false), systems);

        private static Handlers.Accounts.AccountRecord AccountRecord(bool administrator) =>
            new(
                Id: 7,
                Email: administrator ? "admin@example.com" : "reviewer@example.com",
                NormalisedEmail: administrator ? "admin@example.com" : "reviewer@example.com",
                PasswordHash: "hash",
                DisplayName: administrator ? "Admin" : "Reviewer",
                IsVerified: true,
                IsAdministrator: administrator,
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
            byte[] image = PublishExecutorTests.Png("PNGDATA");
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
        public async Task A_reviewer_of_ANOTHER_system_cannot_approve_and_NOTHING_is_written()
        {
            // *** THE STRUCTURAL GUARANTEE, asserted on the DISK rather than on a bool. *** A
            // reviewer's authority stops at their own systems; the only honest way to test that
            // is to check nothing changed.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id,
                ApprovePublishFlowTests.ReviewerOf("Commodore/C128/310378"),
                this.thisDataTree,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsForbidden);
            Assert.Contains(ApprovePublishFlowTests.SystemId, outcome.Error, StringComparison.Ordinal);

            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));

            // ...and the submission is untouched, so somebody entitled can still approve it.
            SubmissionRecord? record = await store.FindAsync(id, CancellationToken.None);
            Assert.Equal(SubmissionState.Pending, record!.State);
        }

        [Fact]
        public async Task A_reviewer_of_THIS_system_publishes_it()
        {
            // The maintainer's model: a reviewer assigned to a board publishes to it, with no
            // administrator in the loop.
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id,
                ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId),
                this.thisDataTree,
                ApprovePublishFlowTests.Now,
                CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // -----------------------------------------------------------------------------------
        // A shared-file change needs the board's reviewer AND the administrator (2026-09-25).
        // -----------------------------------------------------------------------------------

        private async Task<(FakeSubmissionStore Store, long Id)> SharedPendingAsync()
        {
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync();
            store.Submissions[id] = store.Submissions[id] with { TouchesSharedFiles = true };
            return (store, id);
        }

        [Fact]
        public async Task The_REVIEWERS_approval_of_a_shared_change_alone_publishes_NOTHING()
        {
            // *** THE MAINTAINER'S RULE, asserted on the DISK. *** The reviewer's approval is
            // recorded and the submission waits - in the queue, as 'approved' - for the
            // administrator's.
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();

            ApproveOutcome outcome = await this.Flow(store, ApprovePublishFlowTests.AccountsWithAReviewer()).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Administrator], outcome.WaitingFor);

            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
            Assert.Equal(SubmissionState.Approved, (await store.FindAsync(id, CancellationToken.None))!.State);
            Assert.Equal(ApproverRole.Reviewer, Assert.Single(await store.GetApprovalsAsync(id)).Role);
            Assert.Contains(await store.GetQueueAsync(100), record => record.Id == id);
        }

        [Fact]
        public async Task The_ADMINISTRATORS_approval_after_the_reviewers_publishes_it()
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();
            FakeAccountStore accounts = ApprovePublishFlowTests.AccountsWithAReviewer();

            await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            ApproveOutcome second = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now.AddMinutes(5), CancellationToken.None);

            Assert.True(second.IsPublished, second.Error);
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
            Assert.Equal(
                [ApproverRole.Reviewer, ApproverRole.Administrator],
                (await store.GetApprovalsAsync(id)).Select(approval => approval.Role));
        }

        [Fact]
        public async Task Either_may_go_FIRST_the_administrator_then_the_reviewer_publishes_too()
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();
            FakeAccountStore accounts = ApprovePublishFlowTests.AccountsWithAReviewer();

            ApproveOutcome first = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(first.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Reviewer], first.WaitingFor);

            ApproveOutcome second = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now.AddMinutes(5), CancellationToken.None);

            Assert.True(second.IsPublished, second.Error);
        }

        [Fact]
        public async Task The_same_role_approving_TWICE_is_refused_and_still_publishes_nothing()
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();
            FakeAccountStore accounts = ApprovePublishFlowTests.AccountsWithAReviewer();

            await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            ApproveOutcome again = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.True(again.IsConflict);
            Assert.Contains("already approved", again.Error, StringComparison.Ordinal);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        // ###########################################################################################
        // *** A BOARD'S SECOND REVIEWER IS TOLD WHO APPROVED - NOT "YOU HAVE" (code review,
        // 2026-09-25). *** The reviewer half is done either way, so the refusal stands; what was
        // wrong was the sentence, which named the wrong person.
        // ###########################################################################################
        [Fact]
        public async Task A_SECOND_reviewer_of_the_board_is_told_a_colleague_approved_rather_than_that_they_did()
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();
            FakeAccountStore accounts = ApprovePublishFlowTests.AccountsWithAReviewer();

            await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            ReviewAccess colleague = ReviewAccess.For(
                ApprovePublishFlowTests.AccountRecord(administrator: false) with { Id = 8, Email = "bo@example.com", DisplayName = "Bo" },
                [ApprovePublishFlowTests.SystemId]);

            ApproveOutcome refused = await this.Flow(store, accounts).ApproveAsync(
                id, colleague, this.thisDataTree, ApprovePublishFlowTests.Now.AddMinutes(1), CancellationToken.None);

            Assert.True(refused.IsConflict);
            Assert.DoesNotContain("You have", refused.Error, StringComparison.Ordinal);
            Assert.Contains("Reviewer (reviewer@example.com) has already approved", refused.Error, StringComparison.Ordinal);
            Assert.Contains("waiting for the administrator", refused.Error, StringComparison.Ordinal);
            Assert.False(Directory.Exists(Path.Combine(this.thisDataTree, "Commodore")));
        }

        // ###########################################################################################
        // *** A POOL ROW THAT CANNOT GIVE THE REVIEWER HALF DOES NOT COUNT (code review,
        // 2026-09-25). *** An account granted a pool and later made administrator by hand keeps its
        // row but approves as the administrator; a locked or unverified one cannot approve at all.
        // Counted as "the board has a reviewer", each made a shared-file change wait for ever for a
        // reviewer approval nobody could give.
        // ###########################################################################################
        [Theory]
        [InlineData("administrator")]
        [InlineData("locked")]
        [InlineData("unverified")]
        public async Task A_pool_row_that_cannot_approve_as_a_reviewer_leaves_the_administrator_to_decide_alone(string kind)
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();

            Handlers.Accounts.AccountRecord member = ApprovePublishFlowTests.AccountRecord(administrator: false) with { Id = 9 };

            member = kind switch
            {
                "administrator" => member with { IsAdministrator = true },
                "locked" => member with { IsLocked = true },
                _ => member with { IsVerified = false }
            };

            var accounts = new FakeAccountStore();
            accounts.Accounts[9] = member;
            accounts.Reviewers.Add((ApprovePublishFlowTests.SystemId, 9));

            ApproveOutcome outcome = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
        }

        // ###########################################################################################
        // *** WHETHER A SHARED FILE CHANGES IS DECIDED AGAINST THE TREE AS IT IS NOW (code review,
        // 2026-09-25). *** The flag was frozen at create: a submission citing a shared file
        // UNCHANGED was an ordinary one-approval item for ever. Once another publish changed that
        // file, this submission's old bytes would REVERT it for every board - on one reviewer's
        // approval, with the administrator never asked.
        // ###########################################################################################
        [Fact]
        public async Task A_shared_file_whose_published_copy_CHANGED_since_the_submission_arrived_needs_both_approvals()
        {
            const string shared = "Commodore/Shared files/Board local files/notes.txt";
            string onDisk = DataTreeBuilder.Full(this.thisDataTree, shared);

            // The submission arrived citing the shared file as it was then - unchanged, so not a
            // shared change at the time...
            byte[] oldBytes = Encoding.UTF8.GetBytes("the notes as they were");
            string oldHash = await this.PutBlobAsync(oldBytes);

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(
                ApprovePublishFlowTests.Manifest(new SubmissionFile { Path = shared, Sha256 = oldHash, SizeBytes = oldBytes.Length }));

            Assert.False(store.Submissions[id].TouchesSharedFiles);

            // ...and another publish has changed the shared file since.
            File.WriteAllText(onDisk, "the notes as another publish left them");

            ApproveOutcome outcome = await this.Flow(store, ApprovePublishFlowTests.AccountsWithAReviewer()).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsAwaitingApproval);
            Assert.Equal([ApproverRole.Administrator], outcome.WaitingFor);
            Assert.Equal("the notes as another publish left them", File.ReadAllText(onDisk));

            // Recorded, so the queue shows it as the administrator's too.
            Assert.True(store.Submissions[id].TouchesSharedFiles);
        }

        // ###########################################################################################
        // *** AN APPROVAL IS RECORDED ONLY FOR SOMETHING THAT CAN BE PUBLISHED (code review,
        // 2026-09-25). *** The first of two approvals used to be recorded - and the administrator
        // mailed that their approval was needed - before the payload was read or the plan built,
        // so a submission the plan refuses was only refused at the SECOND approval, in the hands of
        // the person who had just been asked to give it.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_the_plan_refuses_is_refused_BEFORE_its_first_approval_is_recorded()
        {
            // A file no row cites: the publish plan refuses it.
            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest(new SubmissionFile
            {
                Path = "Commodore/C64/250407/notes.txt", Sha256 = new string('b', 64), SizeBytes = 3
            });
            manifest.Rows.BoardLocalFiles.Clear();

            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(manifest);
            store.Submissions[id] = store.Submissions[id] with { TouchesSharedFiles = true };

            ApproveOutcome outcome = await this.Flow(store, ApprovePublishFlowTests.AccountsWithAReviewer()).ApproveAsync(
                id, ApprovePublishFlowTests.ReviewerOf(ApprovePublishFlowTests.SystemId), this.thisDataTree,
                ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(outcome.IsAwaitingApproval);
            Assert.Contains("cannot be published", outcome.Error, StringComparison.Ordinal);
            Assert.Empty(await store.GetApprovalsAsync(id));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // ###########################################################################################
        // *** THE REQUIREMENT SHRINKING AFTER THE FIRST APPROVAL DOES NOT STRAND THE ITEM (code
        // review, 2026-09-25). *** The administrator approves first; the board's only reviewer then
        // leaves its pool, so the change needs the administrator alone - who was refused as
        // "already approved" while nobody else could approve. See ApprovalRules.Status.
        // ###########################################################################################
        [Fact]
        public async Task A_shared_change_whose_reviewer_LEAVES_after_the_administrator_approved_is_then_published_by_the_administrator()
        {
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();
            FakeAccountStore accounts = ApprovePublishFlowTests.AccountsWithAReviewer();

            ApproveOutcome first = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(first.IsAwaitingApproval);

            accounts.Reviewers.Clear();

            ApproveOutcome second = await this.Flow(store, accounts).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree,
                ApprovePublishFlowTests.Now.AddMinutes(5), CancellationToken.None);

            Assert.True(second.IsPublished, second.Error);
            Assert.Equal(SubmissionState.Merged, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        [Fact]
        public async Task A_shared_change_on_a_board_with_NO_reviewers_is_published_by_the_administrator_alone()
        {
            // Nobody to ask for the other half.
            (FakeSubmissionStore store, long id) = await this.SharedPendingAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now, CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
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
                    Path = "Commodore/C64/250407/Images/missing.png", Sha256 = new string('a', 64), SizeBytes = 10
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
        // ###########################################################################################
        // FILES THE BOARD NO LONGER USES ARE REMOVED - AND ONLY THOSE THE REVIEWER WAS SHOWN
        // (maintainer decisions, 2026-09-25: "there must be no orphan files", and the list "must be
        // visible BEFORE the reviewer/admin approves it").
        //
        // These need a REAL master: the placeholder the constructor writes makes the tree
        // unreadable, and an unreadable tree removes nothing.
        // ###########################################################################################
        private void UseRealTree(params string[] alsoListed)
        {
            File.Delete(Path.Combine(this.thisDataTree, "Classic-Repair-Toolbox.xlsx"));
            DataTreeBuilder.Master(this.thisDataTree, [DataTreeBuilder.Workbook, .. alsoListed]);
        }

        private const string OldManual = "Commodore/C64/250407/old-manual.pdf";

        // The published board cites a manual; the submission (rows only) no longer does.
        private async Task<(FakeSubmissionStore Store, long Id, SubmissionManifest Manifest)> DroppingTheManualAsync(params string[] alsoCited)
        {
            this.UseRealTree();
            DataTreeBuilder.Board(this.thisDataTree, DataTreeBuilder.Workbook, [ApprovePublishFlowTests.OldManual, .. alsoCited]);
            DataTreeBuilder.Files(this.thisDataTree, ApprovePublishFlowTests.OldManual);
            DataTreeBuilder.Files(this.thisDataTree, alsoCited);

            SubmissionManifest manifest = ApprovePublishFlowTests.Manifest();
            (FakeSubmissionStore store, long id) = await ApprovePublishFlowTests.PendingAsync(manifest);
            return (store, id, manifest);
        }

        [Fact]
        public async Task The_reviewer_is_shown_what_a_publish_would_remove_before_approving()
        {
            (_, _, SubmissionManifest manifest) = await this.DroppingTheManualAsync();

            BoardData? published = await new PublishedBoardReader(NullLogger<PublishedBoardReader>.Instance)
                .TryReadAsync(this.thisDataTree, manifest);

            FileRemovalPreview preview = ApprovePublishFlow.PreviewRemovals(this.thisDataTree, manifest, published, ApprovePublishFlowTests.Now);

            Assert.False(preview.IsBlocked, preview.BlockedBecause);
            Assert.Equal([ApprovePublishFlowTests.OldManual], preview.Files);
        }

        [Fact]
        public async Task Approving_with_the_list_shown_publishes_and_removes_exactly_that_list()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                [ApprovePublishFlowTests.OldManual], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ApprovePublishFlowTests.OldManual], outcome.RemovedFiles);
            Assert.False(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.OldManual)));
        }

        // What is removed is what was on screen. A reviewer shown a different list - here an empty
        // one, because the list changed since they opened it - is refused, and NOTHING is written.
        [Fact]
        public async Task Approving_with_a_different_list_is_refused_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();
            string workbook = DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.Workbook);
            byte[] before = File.ReadAllBytes(workbook);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: [], CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.True(outcome.IsConflict);
            Assert.Equal(ApprovePublishFlow.RemovalsChangedMessage, outcome.Error);
            Assert.Equal(before, File.ReadAllBytes(workbook));
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, ApprovePublishFlowTests.OldManual)));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        // ###########################################################################################
        // *** NO LIST AT ALL IS AN OLDER REVIEW APPLICATION, NOT A CHANGED ONE (code review,
        // 2026-09-25). *** The current one always sends its list, empty or not; one built before
        // the list existed sends none. It used to be told the list had "changed since you opened
        // it" - false, and reopening changed nothing, so it could never publish and never learn
        // why. It is told to update instead - still refused, still nothing written.
        // ###########################################################################################
        [Fact]
        public async Task A_review_application_that_sends_no_list_is_told_to_update_and_nothing_is_written()
        {
            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync();
            string workbook = DataTreeBuilder.Full(this.thisDataTree, DataTreeBuilder.Workbook);
            byte[] before = File.ReadAllBytes(workbook);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                shownRemovals: null, CancellationToken.None);

            Assert.False(outcome.IsPublished);
            Assert.False(outcome.IsConflict);
            Assert.Equal(ApprovePublishFlow.RemovalsNotSentMessage, outcome.Error);
            Assert.Equal(before, File.ReadAllBytes(workbook));
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

        // A shared file one board stops citing stays while another board still cites it - it is
        // not even on the list.
        [Fact]
        public async Task A_shared_file_another_board_still_uses_is_neither_listed_nor_removed()
        {
            const string shared = "Commodore/Shared files/Board local files/manual.pdf";
            const string other = "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx";

            (FakeSubmissionStore store, long id, _) = await this.DroppingTheManualAsync(shared);
            DataTreeBuilder.Master(this.thisDataTree, DataTreeBuilder.Workbook, other);
            DataTreeBuilder.Board(this.thisDataTree, other, shared);

            ApproveOutcome outcome = await this.Flow(store).ApproveAsync(
                id, ApprovePublishFlowTests.Account(), this.thisDataTree, ApprovePublishFlowTests.Now,
                [ApprovePublishFlowTests.OldManual], CancellationToken.None);

            Assert.True(outcome.IsPublished, outcome.Error);
            Assert.Equal([ApprovePublishFlowTests.OldManual], outcome.RemovedFiles);
            Assert.True(File.Exists(DataTreeBuilder.Full(this.thisDataTree, shared)));
        }

        // A tree that cannot be read completely removes nothing - and says so on the list, which
        // is then empty, so the approval goes through with nothing removed.
        [Fact]
        public async Task An_unreadable_tree_removes_nothing_and_the_publish_still_goes_through()
        {
            (FakeSubmissionStore store, long id, SubmissionManifest manifest) = await this.DroppingTheManualAsync();
            DataTreeBuilder.Master(this.thisDataTree, DataTreeBuilder.Workbook, "Amstrad/CPC/464/Data CPC 464 v2.0.0.xlsx");

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
    }
}
