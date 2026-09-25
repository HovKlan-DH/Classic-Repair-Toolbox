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

        // Contributing requires no account - the ordinary submitter is anonymous with a contact
        // address. See NewContributeStrategy.md, "CONTRIBUTING NEEDS NO ACCOUNT".
        private static Submitter Contributor(string email = "dennis@example.com") =>
            Submitter.Anonymous(email, "192.0.2.1");

        private static byte[] Bytes(string content) => Encoding.UTF8.GetBytes(content);

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
                        Path = "main.png",
                        Sha256 = SubmissionFlowTests.HashOf(imageContent),
                        SizeBytes = SubmissionFlowTests.Bytes(imageContent).Length
                    }
                },
                Rows = new SubmissionRows
                {
                    Schematics = { new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "main.png" } },
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
            Assert.Equal("main.png", finding.Subject);
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
