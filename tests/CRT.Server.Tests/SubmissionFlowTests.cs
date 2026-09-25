using System.Security.Cryptography;
using System.Text;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the three-step submission transport: create and negotiate, upload, finalise.
    //
    // THE BLOB STORE IS REAL HERE, against a temp folder. That is a deliberate exception to the
    // usual "fake the I/O" rule, and the reason is that the properties worth testing ARE the
    // filesystem behaviour: that a resumed upload continues rather than truncating, that a
    // mismatched hash is caught and the partial discarded, that a completed blob is content-
    // addressed and shared. A fake blob store would assert that the fake works.
    //
    // No network, no display, no spawned process - the temp folder is the same seam TempWorkspace
    // provides in CRT.App.Tests, so this stays inside CLAUDE.md test rule 6.
    //
    // THE PROPERTIES THAT MATTER MOST, in order:
    //   1. A submission is only ever touched by whoever holds its capability token.
    //   2. Uploaded bytes are verified against the declared hash, always.
    //   3. Nothing here publishes; a finalised submission is queued.
    //   4. A typo fix uploads nothing.
    // ###########################################################################################
    public sealed class SubmissionFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private readonly string thisRoot;

        public SubmissionFlowTests()
        {
            this.thisRoot = Path.Combine(Path.GetTempPath(), "crt-submission-tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(this.thisRoot);
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

        private BlobStore Blobs() => new(this.thisRoot, NullLogger<BlobStore>.Instance);

        private string SystemFolder() => Path.Combine(this.thisRoot, "system");

        // The published tree under SystemFolder(), read the way the running service reads it.
        private PublishedTreeView? Tree() => PublishedTreeProbe.For(this.SystemFolder());

        // Puts a file into the published tree at a data-root-relative path.
        private void Publish(string relativePath, string content)
        {
            string path = Path.Combine(this.SystemFolder(), relativePath.Replace('/', Path.DirectorySeparatorChar));

            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, SubmissionFlowTests.Bytes(content));
        }

        // Contributing requires no account - the ordinary submitter is anonymous with a contact
        // address. See NewContributeStrategy.md, "CONTRIBUTING NEEDS NO ACCOUNT".
        private static Submitter Contributor(string email = "dennis@example.com") =>
            Submitter.Anonymous(email, "192.0.2.1");

        // ###########################################################################################
        // Every test file is a REAL PNG as far as its opening bytes go (security review, 2026-09-25).
        // Finalise now checks a file's bytes against its name, so "PNGDATA" alone - which used to
        // stand in for an image - is correctly refused as not being one. The content after the
        // signature is still the readable marker each test chooses.
        // ###########################################################################################
        private static readonly byte[] PngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

        private static byte[] Bytes(string content) =>
            [.. SubmissionFlowTests.PngSignature, .. Encoding.UTF8.GetBytes(content)];

        // Where the fixture's schematic image lives - data-root-relative, inside the board's own
        // folder, the shape every real submitted path has. A bare "main.png" is now refused as
        // belonging to no board this submission may change.
        private const string MainPath = "Commodore/C64/250407/main.png";

        private static string HashOf(string content) =>
            Convert.ToHexStringLower(SHA256.HashData(SubmissionFlowTests.Bytes(content)));

        // A manifest with one schematic whose image is the given content.
        private static SubmissionManifest Manifest(string imageContent = "PNGDATA")
        {
            return new SubmissionManifest
            {
                // "Manufacturer/Hardware/Board" - the `systems` table's own primary key, and what
                // a real client sends. It has to agree with the Hardware/Board below, which the
                // validator now checks.
                SystemId = "Commodore/C64/250407",
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Summary = "Corrected R12.",
                CreatedUtc = SubmissionFlowTests.Now,
                Files =
                {
                    new SubmissionFile
                    {
                        Path = SubmissionFlowTests.MainPath,
                        Sha256 = SubmissionFlowTests.HashOf(imageContent),
                        SizeBytes = SubmissionFlowTests.Bytes(imageContent).Length
                    }
                },
                Rows = new SubmissionRows
                {
                    Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = SubmissionFlowTests.MainPath } },
                    Components = { new ComponentEntry { BoardLabel = "R12" } }
                }
            };
        }

        // -----------------------------------------------------------------------------------
        // Create and negotiate.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Creating_a_submission_asks_for_the_blobs_the_server_lacks()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.True(outcome.IsAccepted);
            Assert.Single(outcome.Negotiation!.MissingHashes);
            Assert.Equal(0, outcome.Negotiation.AlreadyHeldCount);
        }

        [Fact]
        public async Task A_file_the_server_ALREADY_HOLDS_is_not_requested_again()
        {
            // "A typo fix uploads a manifest and nothing else" - the definition of done. This is
            // also what stops a shared image being re-sent for every board that references it.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest first = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome initial = await SubmissionFlows.CreateAsync(
                first, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            await this.UploadAsync(initial.Negotiation!, first.Files[0], "PNGDATA", store, blobs);

            // A second submission of the same content.
            SubmissionCreationOutcome second = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            Assert.Empty(second.Negotiation!.MissingHashes);
            Assert.Equal(1, second.Negotiation.AlreadyHeldCount);
            Assert.Equal(0, second.Negotiation.TotalBytesToUpload);
        }

        // ###########################################################################################
        // *** A FILE THE PUBLISHED TREE ALREADY HOLDS UNCHANGED IS NOT UPLOADED (2026-09-25). ***
        //
        // The test above only ever proved the second submission of content the BLOB STORE had
        // seen. The published tree was never in it, so the FIRST submission to any board - the
        // commonest case there is - uploaded the whole board: changing one component's text on a
        // shipped board sent 1,212 files and 121 MB (reported). The server holds those bytes on
        // its own disk, and now takes them from there.
        //
        // SystemFolder() is the data-tree root here, exactly as in production, where the endpoint
        // passes DataTreeRoot as both the containment root and the tree the probe reads.
        // ###########################################################################################
        [Fact]
        public async Task A_file_already_PUBLISHED_unchanged_is_not_uploaded_and_the_submission_still_finalises()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            this.Publish(SubmissionFlowTests.MainPath, "PNGDATA");

            SubmissionManifest manifest = SubmissionFlowTests.Manifest("PNGDATA");

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now, CancellationToken.None, this.Tree());

            Assert.True(created.IsAccepted);
            Assert.Empty(created.Negotiation!.MissingHashes);
            Assert.Equal(1, created.Negotiation.AlreadyHeldCount);
            Assert.Equal(0, created.Negotiation.TotalBytesToUpload);

            // Nothing was uploaded, so nothing counts against the sender's upload budget.
            Assert.Equal(0, store.Created[created.Negotiation.SubmissionId].BytesToUpload);

            // The bytes ARE in the store - the reviewer reads them from there and the publish copies
            // them from there, so a submission told "already held" must be able to finish.
            Assert.True(blobs.Contains(manifest.Files[0].Sha256));

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation.SubmissionId, created.Negotiation.UploadToken,
                store, blobs, SubmissionFlowTests.Now);

            Assert.True(
                result.IsAccepted,
                "findings: " + string.Join("; ", result.Findings.Select(finding => finding.Code)));
        }

        // The file the contributor actually changed is still asked for - "unchanged" is decided by
        // the bytes at that path, never by the path being published at all.
        [Fact]
        public async Task A_published_file_the_submission_CHANGED_is_still_asked_for()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            this.Publish(SubmissionFlowTests.MainPath, "THE OLD SCAN");

            SubmissionManifest manifest = SubmissionFlowTests.Manifest("THE NEW SCAN");

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now, CancellationToken.None, this.Tree());

            Assert.Equal(manifest.Files[0].Sha256, Assert.Single(created.Negotiation!.MissingHashes));
            Assert.Equal(manifest.Files[0].SizeBytes, created.Negotiation.TotalBytesToUpload);
            Assert.False(blobs.Contains(manifest.Files[0].Sha256));
        }

        // ###########################################################################################
        // *** A STALE ANSWER FROM THE TREE FALLS BACK TO AN UPLOAD, NEVER TO A BAD BLOB. *** The
        // tree's hashes are cached on length and write time, so the tree can say "unchanged" about
        // a file that no longer is. The import hashes what it copies; on a mismatch nothing enters
        // the store and the file is simply asked for - and then it counts as uploaded bytes.
        // ###########################################################################################
        [Fact]
        public async Task When_the_published_copy_does_not_match_what_the_tree_reported_the_file_is_asked_for()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            this.Publish(SubmissionFlowTests.MainPath, "SOMETHING ELSE ENTIRELY");

            SubmissionManifest manifest = SubmissionFlowTests.Manifest("PNGDATA");

            // A view that claims the published file is exactly the submitted one.
            var staleTree = new PublishedTreeView(
                path => path == SubmissionFlowTests.MainPath ? manifest.Files[0].Sha256 : null,
                _ => null);

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now, CancellationToken.None, staleTree);

            Assert.True(created.IsAccepted);
            Assert.Equal(manifest.Files[0].Sha256, Assert.Single(created.Negotiation!.MissingHashes));
            Assert.Equal(0, created.Negotiation.AlreadyHeldCount);
            Assert.Equal(manifest.Files[0].SizeBytes, store.Created[created.Negotiation.SubmissionId].BytesToUpload);
            Assert.False(blobs.Contains(manifest.Files[0].Sha256));
        }

        // Without a view of the tree nothing is taken from it: "could not be consulted" is not
        // "unchanged", and an upload is the safe answer.
        [Fact]
        public async Task Without_a_view_of_the_published_tree_nothing_is_taken_from_it()
        {
            var store = new FakeSubmissionStore();

            this.Publish(SubmissionFlowTests.MainPath, "PNGDATA");

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest("PNGDATA"), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.Single(created.Negotiation!.MissingHashes);
        }

        // ###########################################################################################
        // WHETHER A SUBMISSION CHANGES SHARED FILES IS DECIDED AT CREATE AND STORED ON THE ROW
        // (Phase 6 roles, 2026-09-25) - it is what routes it to the administrator rather than to
        // the board's reviewers, and the queue reads it without loading the payload.
        // ###########################################################################################
        [Fact]
        public async Task A_submission_adding_a_SHARED_file_is_recorded_as_touching_shared_files()
        {
            var store = new FakeSubmissionStore();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Files.Add(new SubmissionFile
            {
                Path = "Commodore/Shared files/74LS08.png",
                Sha256 = SubmissionFlowTests.HashOf("shared"),
                SizeBytes = SubmissionFlowTests.Bytes("shared").Length
            });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { File = "Commodore/Shared files/74LS08.png" });

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now, CancellationToken.None, this.Tree());

            Assert.True(created.IsAccepted, string.Join("; ", created.Findings.Select(finding => finding.Message)));

            SubmissionRecord? record = await store.FindAsync(created.Negotiation!.SubmissionId);
            Assert.True(record!.TouchesSharedFiles);
        }

        [Fact]
        public async Task An_ordinary_submission_to_a_boards_OWN_folder_does_not_touch_shared_files()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now, CancellationToken.None, this.Tree());

            Assert.False((await store.FindAsync(created.Negotiation!.SubmissionId))!.TouchesSharedFiles);
        }

        // ###########################################################################################
        // A SUBMISSION CREATES THE `systems` ROW IT NEEDS.
        //
        // submissions.system_id is NOT NULL with a foreign key to systems(system_id). Nothing ever
        // wrote to `systems`, so against the real database EVERY submission would have failed on
        // that foreign key - and no test noticed, because the fake store accepted a submission for
        // a system that did not exist. See migration 0004's header.
        //
        // This is also the FIRST submission for a system the contributor invented locally, which is
        // the case the whole "new system" path exists for.
        // ###########################################################################################
        [Fact]
        public async Task A_first_submission_creates_the_system_row_it_refers_to()
        {
            var store = new FakeSubmissionStore();

            await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            NewSubmission system = Assert.Single(store.Systems).Value;

            Assert.Equal("Commodore/C64/250407", system.SystemId);

            // The three parts are stored separately so the review app can list by manufacturer
            // without parsing the id - 0001_initial.sql says so explicitly.
            Assert.Equal("Commodore", system.Manufacturer);
            Assert.Equal("C64", system.Hardware);
            Assert.Equal("250407", system.Board);
        }

        // A system that already exists is left alone. An ON DUPLICATE KEY UPDATE would rewrite its
        // origin on every submission, turning a system that SHIPPED with CRT into a "contributed"
        // one the first time anybody corrected a typo in it.
        [Fact]
        public async Task A_second_submission_for_the_same_system_does_not_create_it_again()
        {
            var store = new FakeSubmissionStore();

            for (int i = 0; i < 2; i++)
            {
                await SubmissionFlows.CreateAsync(
                    SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, this.Blobs(), SubmissionFlowTests.Now);
            }

            Assert.Single(store.Systems);
            Assert.Equal(2, store.Submissions.Count);
        }

        // ###########################################################################################
        // A COMPLETE SUBMISSION IS ACCEPTED - the end-to-end path, and the test this file lacked.
        //
        // Validation runs TWICE: once at create against the manifest as sent, and again at finalise
        // against the manifest RELOADED from the store. Nothing here ever asserted that the second
        // pass passes, so a reload that silently dropped fields went unnoticed - and it did:
        // MySqlSubmissionStore.LoadPayloadAsync rebuilt everything except Manufacturer, Hardware
        // and Board, so the finalise pass saw three empty strings and rejected EVERY submission
        // with "does not name the hardware" / "does not name the board" plus a consequent id
        // mismatch. Three live submissions were spent chasing that in the client, because the
        // findings say "a fault in the submitting application".
        //
        // This is the regression test. It fails against a store whose reload loses the name parts.
        // ###########################################################################################
        [Fact]
        public async Task A_complete_submission_is_ACCEPTED_and_survives_the_reload_at_finalise()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            Assert.True(created.IsAccepted);

            await this.UploadAsync(created.Negotiation!, manifest.Files[0], "PNGDATA", store, blobs);

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation!.SubmissionId, created.Negotiation.UploadToken,
                store, blobs, SubmissionFlowTests.Now);

            // The whole point: no findings, and QUEUED rather than rejected.
            Assert.True(
                result.IsAccepted,
                "a valid submission must survive finalise - findings: " +
                string.Join("; ", result.Findings.Select(finding => finding.Code)));

            Assert.Empty(result.Findings);
            Assert.Equal(SubmissionState.Pending, result.State);
        }

        // The identity the client sent must come back intact from the store, because the finalise
        // pass validates against the RELOADED manifest rather than the one in flight.
        [Fact]
        public async Task The_reloaded_manifest_still_names_its_manufacturer_hardware_and_board()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            SubmissionManifest? reloaded = await store.LoadPayloadAsync(created.Negotiation!.SubmissionId);

            Assert.NotNull(reloaded);
            Assert.Equal("Commodore", reloaded!.Manufacturer);
            Assert.Equal("C64", reloaded.Hardware);
            Assert.Equal("250407", reloaded.Board);
            Assert.Equal("Commodore/C64/250407", reloaded.SystemId);
        }

        [Fact]
        public async Task Contributing_needs_NO_account()
        {
            // The maintainer's decision (NewContributeStrategy.md, "CONTRIBUTING NEEDS NO
            // ACCOUNT"): a sign-up wall before a hobbyist can fix a typo is how a contribution
            // does not happen. An anonymous submitter with a contact address is the ordinary case,
            // not an exception.
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.Anonymous("someone@example.com", "192.0.2.1"),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.True(outcome.IsAccepted);
            Assert.Single(store.Submissions);
            Assert.Null(store.Submissions[1].AccountId);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public async Task A_submission_with_no_contact_address_is_refused(string? email)
        {
            // The ONE thing a contributor must give. Without it a reviewer cannot say "accepted"
            // or "this needs changing", so the contribution can only be taken or dropped.
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.Anonymous(email, "192.0.2.1"),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "contact.missing");
            Assert.Empty(store.Submissions);
        }

        [Fact]
        public async Task The_contact_refusal_says_no_account_is_being_created()
        {
            // Somebody asked for an email address expects a sign-up to follow. Saying plainly that
            // one is not coming is the difference between a contribution and an abandoned one.
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.Anonymous(null, null),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            string message = Assert.Single(outcome.Findings).Message;

            Assert.Contains("no account", message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("no password", message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task A_malformed_contact_address_is_refused()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.Anonymous("not-an-address", "192.0.2.1"),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "contact.malformed");
        }

        [Fact]
        public async Task A_signed_in_maintainer_needs_no_separate_contact_address()
        {
            // Their account already carries one.
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.SignedIn(accountId: 7, "192.0.2.1"),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.True(outcome.IsAccepted);
            Assert.Equal(7, store.Submissions[1].AccountId);
        }

        [Fact]
        public async Task Creating_a_submission_returns_a_capability_token()
        {
            // With no account, this token is the ONLY thing standing between a submission and
            // anyone who can guess a small integer.
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(),
                this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(string.IsNullOrWhiteSpace(outcome.Negotiation!.UploadToken));

            // Only the HASH is stored - a database read must not hand somebody the ability to act
            // on submissions in flight.
            Assert.NotEqual(outcome.Negotiation.UploadToken, store.Submissions[1].UploadTokenHash);
        }

        [Fact]
        public async Task A_manifest_with_a_TRAVERSING_path_is_refused_and_creates_no_submission()
        {
            // Refused before a row exists, so a broken client cannot fill the table with rows that
            // will never be finalised - and the contributor learns before uploading anything.
            var store = new FakeSubmissionStore();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Files[0] = new SubmissionFile
            {
                Path = "../../evil.png",
                Sha256 = SubmissionFlowTests.HashOf("x"),
                SizeBytes = 1
            };

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "path.rejected");
            Assert.Empty(store.Submissions);
        }

        [Fact]
        public async Task A_manifest_whose_ROWS_are_invalid_is_refused_before_any_upload()
        {
            var store = new FakeSubmissionStore();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "R12" });

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "component.duplicate_label");
            Assert.Empty(store.Submissions);
        }

        // -----------------------------------------------------------------------------------
        // Upload. Ownership and hash verification are the two that matter.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("a-token-that-was-never-issued")]
        public async Task Without_the_RIGHT_TOKEN_nobody_can_upload_into_a_submission(string? token)
        {
            // THE authorisation boundary of the whole phase. A submission id is a small
            // consecutive integer, so without this check anyone could replace a file in a
            // submission somebody else is about to have reviewed and merged.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            using var content = new MemoryStream(SubmissionFlowTests.Bytes("PNGDATA"));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, content,
                token, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.NotFound, outcome.Status);
        }

        [Fact]
        public async Task A_wrong_token_and_a_MISSING_submission_answer_identically()
        {
            // Distinguishing them would confirm which ids exist, letting a caller enumerate
            // submissions by walking integers.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            using var first = new MemoryStream(SubmissionFlowTests.Bytes("PNGDATA"));
            using var second = new MemoryStream(SubmissionFlowTests.Bytes("PNGDATA"));

            BlobUploadOutcome wrongToken = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, first,
                "wrong", store, blobs, SubmissionFlowTests.Now);

            BlobUploadOutcome noSuchSubmission = await SubmissionFlows.UploadChunkAsync(
                999_999, manifest.Files[0].Sha256, 0, second,
                created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(noSuchSubmission.Status, wrongToken.Status);
        }

        [Fact]
        public async Task Uploaded_bytes_that_do_NOT_match_the_declared_hash_are_refused()
        {
            // THE property that makes "content-addressed" true rather than a lie the client can
            // tell. Without it, uploading arbitrary bytes under the hash of an innocent file would
            // let a submission reference something the server believes it has vetted.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            // Right length, wrong content.
            using var content = new MemoryStream(SubmissionFlowTests.Bytes("EVILDAT"));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, content,
                created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.HashMismatch, outcome.Status);

            // And the blob is NOT in the store.
            Assert.False(blobs.Contains(manifest.Files[0].Sha256));
        }

        [Fact]
        public async Task A_hash_the_submission_never_asked_for_is_refused()
        {
            // Otherwise the upload endpoint becomes a way to put arbitrary content into the shared
            // blob store under a hash of the caller's choosing - which a LATER submission could
            // then reference without uploading it.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            using var content = new MemoryStream(SubmissionFlowTests.Bytes("SOMETHING ELSE"));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, SubmissionFlowTests.HashOf("SOMETHING ELSE"), 0,
                content, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.UnexpectedHash, outcome.Status);
        }

        [Fact]
        public async Task An_upload_RESUMES_rather_than_starting_again()
        {
            // "A contributor whose 300 MB upload dies at 90% and must start again will not start
            // again." Resumption works because the partial file's LENGTH is the offset - there is
            // no separate bookkeeping to get out of step with the bytes on disk.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            const string whole = "PNGDATA-LONGER-CONTENT";
            SubmissionManifest manifest = SubmissionFlowTests.Manifest(whole);

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            long id = created.Negotiation!.SubmissionId;
            string hash = manifest.Files[0].Sha256;

            // First half.
            byte[] all = SubmissionFlowTests.Bytes(whole);

            using (var first = new MemoryStream(all[..10]))
            {
                BlobUploadOutcome partial = await SubmissionFlows.UploadChunkAsync(
                    id, hash, 0, first, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

                Assert.Equal(BlobUploadStatus.Partial, partial.Status);
                Assert.Equal(10, partial.ResumeFrom);
            }

            // Resume from where it stopped.
            using (var second = new MemoryStream(all[10..]))
            {
                BlobUploadOutcome complete = await SubmissionFlows.UploadChunkAsync(
                    id, hash, 10, second, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

                Assert.Equal(BlobUploadStatus.Completed, complete.Status);
            }

            Assert.True(blobs.Contains(hash));
        }

        [Fact]
        public async Task Resuming_from_the_WRONG_offset_is_refused_and_says_where_to_resume()
        {
            // Appending anyway would silently corrupt the blob - it would fail the hash check, but
            // only after the whole upload completed, which is exactly what resumption exists to
            // avoid.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            const string whole = "PNGDATA-LONGER-CONTENT";
            SubmissionManifest manifest = SubmissionFlowTests.Manifest(whole);

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            long id = created.Negotiation!.SubmissionId;
            string hash = manifest.Files[0].Sha256;
            byte[] all = SubmissionFlowTests.Bytes(whole);

            using (var first = new MemoryStream(all[..10]))
            {
                await SubmissionFlows.UploadChunkAsync(
                    id, hash, 0, first, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);
            }

            using var wrong = new MemoryStream(all[15..]);

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                id, hash, 15, wrong, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.ChunkRejected, outcome.Status);
            Assert.Equal(10, outcome.ResumeFrom);
        }

        [Fact]
        public async Task Uploading_after_the_window_has_passed_is_refused()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            using var content = new MemoryStream(SubmissionFlowTests.Bytes("PNGDATA"));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, content,
                created.Negotiation.UploadToken, store, blobs,
                SubmissionFlowTests.Now + SubmissionFlows.UploadWindow + TimeSpan.FromMinutes(1));

            Assert.Equal(BlobUploadStatus.Expired, outcome.Status);
        }

        // -----------------------------------------------------------------------------------
        // Finalise.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_complete_submission_is_QUEUED_not_published()
        {
            // Nothing in this pipeline publishes. Promotion to the production tree stays a manual
            // act by the maintainer, and Phase 3's filesystem permissions make it impossible for
            // this service to do otherwise.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            await this.UploadAsync(created.Negotiation!, manifest.Files[0], "PNGDATA", store, blobs);

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation!.SubmissionId, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.True(result.IsAccepted);
            Assert.Equal(SubmissionState.Pending, result.State);
        }

        [Fact]
        public async Task Finalising_with_files_still_missing_is_refused_and_names_them()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation!.SubmissionId, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.False(result.IsAccepted);

            ValidationFinding finding = Assert.Single(result.Findings);
            Assert.Equal("file.not_uploaded", finding.Code);
            Assert.Equal(SubmissionFlowTests.MainPath, finding.Subject);
        }

        [Fact]
        public async Task Without_the_right_token_nobody_can_FINALISE_a_submission()
        {
            // Finalising is what puts work in front of a reviewer. A stranger finalising somebody
            // else's half-uploaded submission would queue an incomplete contribution in their name.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation!.SubmissionId, "not-the-right-token",
                store, blobs, SubmissionFlowTests.Now);

            Assert.False(result.IsAccepted);
            Assert.Contains(result.Findings, finding => finding.Code == "submission.not_found");
        }

        [Fact]
        public async Task Finalising_TWICE_answers_with_the_state_rather_than_failing()
        {
            // A client that retried after a dropped connection should get a sensible answer, not a
            // failure for something that actually worked.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            long id = created.Negotiation!.SubmissionId;

            await this.UploadAsync(created.Negotiation!, manifest.Files[0], "PNGDATA", store, blobs);

            await SubmissionFlows.FinaliseAsync(
                id, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            SubmissionResult again = await SubmissionFlows.FinaliseAsync(
                id, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.True(again.IsAccepted);
            Assert.Equal(SubmissionState.Pending, again.State);
        }

        // -----------------------------------------------------------------------------------
        // Garbage collection.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task An_abandoned_submission_is_collected_after_its_window()
        {
            // Blobs from abandoned submissions must be collected or the disk fills quietly - named
            // as a trap in NewContributeStrategy.md.
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            int collected = await SubmissionFlows.CollectAbandonedAsync(
                store, blobs, SubmissionFlowTests.Now + SubmissionFlows.UploadWindow + TimeSpan.FromMinutes(1));

            Assert.Equal(1, collected);
            Assert.Equal(
                SubmissionState.Abandoned,
                store.Submissions[created.Negotiation!.SubmissionId].State);
        }

        [Fact]
        public async Task A_submission_still_inside_its_window_is_NOT_collected()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(0, await SubmissionFlows.CollectAbandonedAsync(
                store, blobs, SubmissionFlowTests.Now + TimeSpan.FromHours(1)));
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: per-address limits.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** CREATING A SUBMISSION USED TO HAVE NO LIMIT OF ANY KIND. *** No account is needed, so
        // the address is all there is to count against. These pin that the count and the byte
        // budget both bite, that the refusal is a sentence the client can show, and that a refused
        // request leaves nothing behind.
        // ###########################################################################################
        [Fact]
        public async Task Too_many_submissions_from_one_address_are_refused_with_a_reason()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            for (int i = 0; i < SubmissionRateLimitPolicy.MaxSubmissionsPerAddress; i++)
            {
                SubmissionCreationOutcome allowed = await SubmissionFlows.CreateAsync(
                    SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, blobs, SubmissionFlowTests.Now);

                Assert.True(allowed.IsAccepted);
            }

            SubmissionCreationOutcome refused = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            Assert.True(refused.IsRateLimited);
            Assert.True(refused.RetryAfter > TimeSpan.Zero);
            Assert.Equal("submission.rate_limited", Assert.Single(refused.Findings).Code);
            Assert.Equal(SubmissionRateLimitPolicy.MaxSubmissionsPerAddress, store.Submissions.Count);
        }

        // Another address is another bucket - a limit that caught everybody would be a way to shut
        // contributions off for all.
        [Fact]
        public async Task The_limit_is_per_address()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            for (int i = 0; i < SubmissionRateLimitPolicy.MaxSubmissionsPerAddress; i++)
            {
                await SubmissionFlows.CreateAsync(
                    SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, blobs, SubmissionFlowTests.Now);
            }

            SubmissionCreationOutcome other = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.Anonymous("other@example.com", "198.51.100.7"),
                this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            Assert.True(other.IsAccepted);
        }

        // A reviewer or administrator is trusted by the database row, and checking things often is
        // their job. An ordinary signed-in account is NOT exempt - anybody may register one.
        [Fact]
        public async Task A_trusted_reviewer_is_not_limited_but_an_ordinary_account_is()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            for (int i = 0; i < SubmissionRateLimitPolicy.MaxSubmissionsPerAddress; i++)
            {
                await SubmissionFlows.CreateAsync(
                    SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, blobs, SubmissionFlowTests.Now);
            }

            SubmissionCreationOutcome reviewer = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.SignedIn(7, "192.0.2.1", isTrusted: true),
                this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            SubmissionCreationOutcome ordinary = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), Submitter.SignedIn(8, "192.0.2.1"),
                this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            Assert.True(reviewer.IsAccepted);
            Assert.True(ordinary.IsRateLimited);
        }

        // The bytes the server was asked to store are what the budget counts, so the create records
        // them - and a file it already held costs nothing.
        [Fact]
        public async Task A_submission_records_the_bytes_it_asked_to_upload()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.Equal(
                created.Negotiation!.TotalBytesToUpload,
                store.Created[created.Negotiation.SubmissionId].BytesToUpload);

            Assert.True(store.Created[created.Negotiation.SubmissionId].BytesToUpload > 0);
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: storage, closed boards, other boards, content.
        // -----------------------------------------------------------------------------------

        // Below the disk reserve, nothing new is accepted - "contributions paused" rather than "the
        // site's disk is full" - and nothing is written.
        [Fact]
        public async Task Below_the_disk_reserve_a_submission_is_refused_and_nothing_is_written()
        {
            var store = new FakeSubmissionStore();
            var blobs = new BlobStore(this.thisRoot, NullLogger<BlobStore>.Instance, () => 100, minimumFreeBytes: 1000);

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            Assert.True(outcome.IsNoRoom);
            Assert.Equal("server.storage_full", Assert.Single(outcome.Findings).Code);
            Assert.Empty(store.Submissions);
        }

        [Fact]
        public async Task Below_the_disk_reserve_an_upload_chunk_is_refused()
        {
            var store = new FakeSubmissionStore();
            var room = new long[] { 10_000 };
            var blobs = new BlobStore(this.thisRoot, NullLogger<BlobStore>.Instance, () => room[0], minimumFreeBytes: 1000);

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            room[0] = 10;

            using var content = new MemoryStream(SubmissionFlowTests.Bytes("PNGDATA"));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, content,
                created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.NoRoom, outcome.Status);
            Assert.Equal(0, blobs.GetUploadedLength(created.Negotiation.SubmissionId, manifest.Files[0].Sha256));
        }

        // is_accepting existed from the first migration and nothing read it. Setting it to 0 now
        // closes a board - the lever for one being flooded.
        [Fact]
        public async Task A_board_closed_to_contributions_refuses_new_submissions()
        {
            var store = new FakeSubmissionStore();

            await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            store.ClosedSystems.Add("Commodore/C64/250407");

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "system.closed");
            Assert.Single(store.Submissions);
        }

        // THE finding the review opened with: a submission to one board carrying another board's
        // file. Refused at create, before a single byte is uploaded.
        [Fact]
        public async Task A_submission_carrying_ANOTHER_boards_file_is_refused_before_any_upload()
        {
            var store = new FakeSubmissionStore();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Files[0].Path = "Commodore/C64/250425/main.png";
            manifest.Rows.Schematics[0] = new BoardSchematicEntry
            {
                SchematicName = "Main",
                SchematicImageFile = "Commodore/C64/250425/main.png"
            };

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now, CancellationToken.None, PublishedTreeView.Empty);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "file.other_board");
            Assert.Empty(store.Submissions);
        }

        // A file no row uses would have been shown nowhere on the review screen.
        [Fact]
        public async Task A_submission_carrying_a_file_NO_ROW_USES_is_refused()
        {
            var store = new FakeSubmissionStore();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Files.Add(new SubmissionFile
            {
                Path = "Commodore/C64/250407/hidden.png",
                Sha256 = SubmissionFlowTests.HashOf("hidden"),
                SizeBytes = 10
            });

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, this.Blobs(), SubmissionFlowTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Contains(outcome.Findings, finding => finding.Code == "file.not_used");
        }

        // The bytes are checked against the name once they are all here. A ".png" that is not an
        // image is rejected before any reviewer sees it.
        [Fact]
        public async Task A_file_whose_BYTES_do_not_match_its_name_is_rejected_at_finalise()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            byte[] notAnImage = Encoding.UTF8.GetBytes("MZ this is not a picture");

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            manifest.Files[0].Sha256 = Convert.ToHexStringLower(SHA256.HashData(notAnImage));
            manifest.Files[0].SizeBytes = notAnImage.Length;

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            using (var content = new MemoryStream(notAnImage))
            {
                BlobUploadOutcome uploaded = await SubmissionFlows.UploadChunkAsync(
                    created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, 0, content,
                    created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

                Assert.Equal(BlobUploadStatus.Completed, uploaded.Status);
            }

            SubmissionResult result = await SubmissionFlows.FinaliseAsync(
                created.Negotiation.SubmissionId, created.Negotiation.UploadToken, store, blobs, SubmissionFlowTests.Now);

            Assert.False(result.IsAccepted);
            Assert.Equal(SubmissionState.Rejected, result.State);
            Assert.Contains(result.Findings, finding => finding.Code == "file.content_mismatch");
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: the upload-state question is scoped too.
        // -----------------------------------------------------------------------------------

        // It used to answer for ANY hash to anyone holding ANY token - an oracle for "has somebody
        // submitted this exact file". Now only for the hashes this submission listed.
        [Fact]
        public async Task The_upload_state_answers_only_for_a_hash_this_submission_listed()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            // Somebody else's completed file sits in the shared store.
            SubmissionManifest theirs = SubmissionFlowTests.Manifest("THEIRS");
            SubmissionCreationOutcome theirCreate = await SubmissionFlows.CreateAsync(
                theirs, Submitter.Anonymous("them@example.com", "198.51.100.1"), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);
            await this.UploadAsync(theirCreate.Negotiation!, theirs.Files[0], "THEIRS", store, blobs);

            SubmissionCreationOutcome mine = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now);

            UploadState? probe = await SubmissionFlows.GetUploadStateAsync(
                mine.Negotiation!.SubmissionId, theirs.Files[0].Sha256, mine.Negotiation.UploadToken, store, blobs);

            UploadState? own = await SubmissionFlows.GetUploadStateAsync(
                mine.Negotiation.SubmissionId, SubmissionFlowTests.Manifest().Files[0].Sha256,
                mine.Negotiation.UploadToken, store, blobs);

            Assert.Null(probe);
            Assert.NotNull(own);
        }

        [Fact]
        public async Task The_upload_state_needs_the_token()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);

            Assert.Null(await SubmissionFlows.GetUploadStateAsync(
                created.Negotiation!.SubmissionId, manifest.Files[0].Sha256, "wrong-token", store, blobs));
        }

        // -----------------------------------------------------------------------------------
        // Security review, 2026-09-25: what nothing used to collect.
        // -----------------------------------------------------------------------------------

        // Every completed blob stayed on disk for good, including every file of every rejected or
        // abandoned submission from anybody. Now a blob no live submission needs is deleted.
        [Fact]
        public async Task A_blob_only_a_REJECTED_submission_needed_is_collected_and_a_PENDING_ones_is_kept()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest rejected = SubmissionFlowTests.Manifest("REJECTED");
            SubmissionCreationOutcome rejectedCreate = await SubmissionFlows.CreateAsync(
                rejected, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);
            await this.UploadAsync(rejectedCreate.Negotiation!, rejected.Files[0], "REJECTED", store, blobs);
            await store.SetStateAsync(rejectedCreate.Negotiation!.SubmissionId, SubmissionState.Rejected, SubmissionFlowTests.Now);

            SubmissionManifest pending = SubmissionFlowTests.Manifest("PENDING");
            SubmissionCreationOutcome pendingCreate = await SubmissionFlows.CreateAsync(
                pending, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);
            await this.UploadAsync(pendingCreate.Negotiation!, pending.Files[0], "PENDING", store, blobs);
            await store.SetStateAsync(pendingCreate.Negotiation!.SubmissionId, SubmissionState.Pending, SubmissionFlowTests.Now);

            int deleted = await SubmissionFlows.CollectUnreferencedBlobsAsync(store, blobs);

            Assert.Equal(1, deleted);
            Assert.False(blobs.Contains(rejected.Files[0].Sha256));
            Assert.True(blobs.Contains(pending.Files[0].Sha256));
        }

        // MERGED blobs are kept: they are what lets the next submission to that board skip
        // re-uploading every file it did not change.
        [Fact]
        public async Task A_blob_a_MERGED_submission_used_is_kept()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            SubmissionManifest manifest = SubmissionFlowTests.Manifest();
            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(), store, blobs, SubmissionFlowTests.Now);
            await this.UploadAsync(created.Negotiation!, manifest.Files[0], "PNGDATA", store, blobs);
            await store.SetStateAsync(created.Negotiation!.SubmissionId, SubmissionState.Merged, SubmissionFlowTests.Now);

            Assert.Equal(0, await SubmissionFlows.CollectUnreferencedBlobsAsync(store, blobs));
            Assert.True(blobs.Contains(manifest.Files[0].Sha256));
        }

        // The collector and "you already have this" share one gate, so a submission told "already
        // held" can never have that blob deleted in between. Proved by holding the gate: a create
        // must WAIT for it.
        [Fact]
        public async Task A_submission_waits_for_the_collector_rather_than_racing_it()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            Task<SubmissionCreationOutcome> create;

            using (await blobs.EnterReferenceGateAsync())
            {
                create = SubmissionFlows.CreateAsync(
                    SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, blobs, SubmissionFlowTests.Now);

                await Task.Delay(200);

                Assert.False(create.IsCompleted, "creating a submission must wait for the reference gate");
            }

            Assert.True((await create).IsAccepted);
        }

        // ###########################################################################################
        // *** A FIRST SUBMISSION NO LONGER HOLDS THE GATE WHILE IT IMPORTS (code review,
        // 2026-09-25). *** The gate is one lock for the whole service. Importing a whole board from
        // the published tree used to happen inside it, so everybody else's create - and the
        // collector - waited for up to ~120 MB to be copied and hashed. The import now happens with
        // the gate held by somebody else; only the last step waits for it.
        //
        // And because the collector can now run between the import and the row being created, what
        // it takes is caught: the blob is confirmed inside the gate, and asked for when it is gone.
        // ###########################################################################################
        [Fact]
        public async Task A_first_submission_imports_outside_the_gate_and_asks_for_a_blob_collected_meanwhile()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            this.Publish(SubmissionFlowTests.MainPath, "PNGDATA");
            SubmissionManifest manifest = SubmissionFlowTests.Manifest("PNGDATA");
            string hash = manifest.Files[0].Sha256;

            Task<SubmissionCreationOutcome> create;

            using (await blobs.EnterReferenceGateAsync())
            {
                create = SubmissionFlows.CreateAsync(
                    manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                    store, blobs, SubmissionFlowTests.Now, CancellationToken.None, this.Tree());

                for (int i = 0; i < 200 && !blobs.Contains(hash); i++)
                    await Task.Delay(50);

                Assert.True(blobs.Contains(hash), "the import must not wait for the reference gate");
                Assert.False(create.IsCompleted, "creating the row must still wait for it");

                // The collector's doing, while the create waits.
                Assert.True(blobs.TryDeleteBlob(hash));
            }

            SubmissionCreationOutcome created = await create;

            Assert.True(created.IsAccepted);
            Assert.Equal(hash, Assert.Single(created.Negotiation!.MissingHashes));
            Assert.Equal(0, created.Negotiation.AlreadyHeldCount);
            Assert.Equal(manifest.Files[0].SizeBytes, store.Created[created.Negotiation.SubmissionId].BytesToUpload);
        }

        // ###########################################################################################
        // A sender whose budget a failed import tips over is refused - and the blobs THIS create
        // imported are taken back, rather than left on disk for the hourly collector (code review,
        // 2026-09-25). Nothing is recorded.
        // ###########################################################################################
        [Fact]
        public async Task A_budget_refusal_after_the_imports_takes_back_what_it_imported()
        {
            var store = new FakeSubmissionStore();
            BlobStore blobs = this.Blobs();

            // Earlier today, from the same address, all but five bytes of the budget.
            await store.CreateAsync(new NewSubmission(
                "Commodore/C64/250407", "Commodore", "C64", "250407", null, "dennis@example.com", "192.0.2.1",
                "hash", "r0", "Earlier.", 1, [], SubmissionFlowTests.Now.AddHours(-1), SubmissionFlowTests.Now,
                SubmissionRateLimitPolicy.MaxUploadBytesPerAddress - 5));

            const string second = "Commodore/C64/250407/second.png";
            this.Publish(SubmissionFlowTests.MainPath, "PNGDATA");
            this.Publish(second, "NOT WHAT THE TREE SAID");

            SubmissionManifest manifest = SubmissionFlowTests.Manifest("PNGDATA");
            manifest.Files.Add(new SubmissionFile
            {
                Path = second,
                Sha256 = SubmissionFlowTests.HashOf("SECOND"),
                SizeBytes = SubmissionFlowTests.Bytes("SECOND").Length
            });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = "Second", File = second });

            // The tree says both are published unchanged; the second's bytes on disk say otherwise,
            // so its import fails and it must be uploaded - fourteen bytes the budget has no room for.
            var tree = new PublishedTreeView(
                path => path == SubmissionFlowTests.MainPath ? manifest.Files[0].Sha256
                    : path == second ? manifest.Files[1].Sha256
                    : null,
                _ => null);

            SubmissionCreationOutcome outcome = await SubmissionFlows.CreateAsync(
                manifest, SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, blobs, SubmissionFlowTests.Now, CancellationToken.None, tree);

            Assert.True(outcome.IsRateLimited);
            Assert.False(blobs.Contains(manifest.Files[0].Sha256), "the import this create made is taken back");
            Assert.Single(store.Created);
        }

        // The bulk of an ended submission - its copy of the board's rows, and its file list - goes
        // after a month; its row and findings stay as the audit trail and the contributor's record.
        [Fact]
        public async Task An_old_ENDED_submissions_rows_are_cleared_but_its_record_is_kept()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            long id = created.Negotiation!.SubmissionId;
            await store.SetStateAsync(id, SubmissionState.Rejected, SubmissionFlowTests.Now);

            Assert.Equal(0, await SubmissionFlows.CollectRetiredAsync(store, SubmissionFlowTests.Now + TimeSpan.FromDays(1)));

            Assert.Equal(1, await SubmissionFlows.CollectRetiredAsync(
                store, SubmissionFlowTests.Now + SubmissionFlows.RetiredRetention + TimeSpan.FromDays(1)));

            Assert.False(store.Payloads.ContainsKey(id));
            Assert.False(store.Files.ContainsKey(id));
            Assert.True(store.Submissions.ContainsKey(id));
        }

        [Fact]
        public async Task A_PENDING_submission_is_never_cleared_however_old()
        {
            var store = new FakeSubmissionStore();

            SubmissionCreationOutcome created = await SubmissionFlows.CreateAsync(
                SubmissionFlowTests.Manifest(), SubmissionFlowTests.Contributor(), this.SystemFolder(),
                store, this.Blobs(), SubmissionFlowTests.Now);

            await store.SetStateAsync(created.Negotiation!.SubmissionId, SubmissionState.Pending, SubmissionFlowTests.Now);

            Assert.Equal(0, await SubmissionFlows.CollectRetiredAsync(store, SubmissionFlowTests.Now + TimeSpan.FromDays(3650)));
            Assert.True(store.Payloads.ContainsKey(created.Negotiation.SubmissionId));
        }

        // A finding's subject is often a label straight from the contributor's rows, and nothing
        // bounded it - a ten-thousand-character label failed the findings INSERT with a 500 after
        // the submission row had been written.
        [Fact]
        public void Findings_are_cut_to_the_columns_they_are_stored_in()
        {
            IReadOnlyList<ValidationFinding> fitted = SubmissionFlows.FitForStorage(
            [
                new ValidationFinding
                {
                    Severity = ValidationSeverity.Error,
                    Code = new string('c', 200),
                    Subject = new string('s', 10_000),
                    Message = new string('m', 10_000)
                }
            ]);

            ValidationFinding finding = Assert.Single(fitted);

            Assert.Equal(SubmissionFlows.MaximumFindingSubjectLength, finding.Subject.Length);
            Assert.Equal(SubmissionFlows.MaximumFindingCodeLength, finding.Code.Length);
            Assert.EndsWith("...", finding.Subject, StringComparison.Ordinal);

            // The message is TEXT - never cut, it is what the contributor reads.
            Assert.Equal(10_000, finding.Message.Length);
        }

        // Uploads a file's content in one go, asserting it completed. Takes the negotiation so it
        // carries the capability token - every call after Create needs it.
        private async Task UploadAsync(
            HashNegotiationResponse negotiation, SubmissionFile file, string content,
            FakeSubmissionStore store, BlobStore blobs)
        {
            using var stream = new MemoryStream(SubmissionFlowTests.Bytes(content));

            BlobUploadOutcome outcome = await SubmissionFlows.UploadChunkAsync(
                negotiation.SubmissionId, file.Sha256, 0, stream, negotiation.UploadToken,
                store, blobs, SubmissionFlowTests.Now);

            Assert.Equal(BlobUploadStatus.Completed, outcome.Status);
        }
    }
}
