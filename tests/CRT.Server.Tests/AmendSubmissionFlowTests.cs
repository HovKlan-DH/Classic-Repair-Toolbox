using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // AmendSubmissionFlow - a maintainer's change to a submission, made in the Maintainer tab's
    // table (owner request, 2026-09-25: "The maintainer should be able to also edit whatever, if
    // he chooses to publish it afterwards").
    //
    // The negative tests come first: an amendment rewrites what will be published, so who may do
    // it, when, and over whose change are the properties that matter most.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class AmendSubmissionFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);
        private const string SystemId = "Commodore/C64/250407";
        private const string Manual = "Commodore/C64/250407/manual.pdf";
        private const string SheetImage = "Commodore/C64/250407/sheet-1.png";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-amend", Guid.NewGuid().ToString("N"));
        private readonly string thisData;
        private readonly string thisBlobs;

        public AmendSubmissionFlowTests()
        {
            this.thisData = Path.Combine(this.thisRoot, "Data");
            this.thisBlobs = Path.Combine(this.thisRoot, "blobs");
            Directory.CreateDirectory(this.thisData);
            Directory.CreateDirectory(this.thisBlobs);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(this.thisRoot, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private BlobStore Blobs() => new(this.thisBlobs, NullLogger<BlobStore>.Instance);

        private static ReviewAccess MaintainerOf(params string[] systems) =>
            ReviewAccess.For(AmendSubmissionFlowTests.Account(administrator: false), systems);

        private static ReviewAccess Administrator() =>
            ReviewAccess.For(AmendSubmissionFlowTests.Account(administrator: true));

        private static AccountRecord Account(bool administrator) =>
            new(7, "anna@example.com", "anna@example.com", "hash", "Anna",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false, AmendSubmissionFlowTests.Now, null);

        // A pending submission of one component, citing the manual as a board local file.
        private static async Task<(FakeSubmissionStore Store, long Id)> PendingAsync(string state = SubmissionState.Pending)
        {
            var store = new FakeSubmissionStore();
            var manifest = new SubmissionManifest
            {
                SystemId = AmendSubmissionFlowTests.SystemId,
                Manufacturer = "Commodore",
                Hardware = "C64",
                Board = "250407",
                Files =
                [
                    new SubmissionFile { Path = AmendSubmissionFlowTests.Manual, Sha256 = new string('a', 64), SizeBytes = 10 },
                    new SubmissionFile { Path = AmendSubmissionFlowTests.SheetImage, Sha256 = new string('b', 64), SizeBytes = 10 }
                ]
            };
            manifest.Rows.RevisionDate = "2026-September-25";
            manifest.Rows.Schematics.Add(new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = AmendSubmissionFlowTests.SheetImage });
            manifest.Rows.Components.Add(new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01" });
            manifest.Rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = "Manual", File = AmendSubmissionFlowTests.Manual });
            manifest.Rows.ComponentHighlights.Add(new ComponentHighlightEntry
            {
                SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4"
            });

            long id = await store.CreateAsync(
                new NewSubmission(
                    manifest.SystemId, manifest.Manufacturer, manifest.Hardware, manifest.Board,
                    null, "contributor@example.com", "192.0.2.1", "hash", "r0", "A change.", 1,
                    [.. manifest.Files], AmendSubmissionFlowTests.Now, AmendSubmissionFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SavePayloadAsync(id, manifest, CancellationToken.None);
            await store.SetStateAsync(id, state, AmendSubmissionFlowTests.Now, CancellationToken.None);

            return (store, id);
        }

        // The rows as the table would send them: everything as submitted, one part number changed.
        private static async Task<SubmissionRows> EditedRowsAsync(FakeSubmissionStore store, long id, Action<SubmissionRows>? edit = null)
        {
            SubmissionManifest current = (await store.LoadPayloadAsync(id, CancellationToken.None))!;
            SubmissionRows rows = SubmissionRowsBoard.FromBoard(SubmissionRowsBoard.ToBoard(current.Rows));
            rows.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-01" };
            edit?.Invoke(rows);
            return rows;
        }

        private Task<AmendOutcome> AmendAsync(
            FakeSubmissionStore store, long id, SubmissionRows rows, ReviewAccess? access = null, int expectedVersion = 0,
            FakeAccountStore? accounts = null, PublishLock? publishLock = null) =>
            AmendSubmissionFlow.AmendAsync(
                access ?? AmendSubmissionFlowTests.MaintainerOf(AmendSubmissionFlowTests.SystemId),
                id, expectedVersion, rows, this.thisData, store, accounts ?? new FakeAccountStore(), this.Blobs(),
                publishLock ?? new PublishLock(), AmendSubmissionFlowTests.Now, CancellationToken.None);

        // ###########################################################################################
        // *** THE SUBMISSION'S KiCad DATA SURVIVES AN AMENDMENT (2026-09-26). *** The amended file
        // list is rebuilt from what the rows cite, and no row cites a KiCad file - so without the
        // carry-over, a maintainer fixing one part number would silently strip the contributor's
        // KiCad data from the submission, and approving it would publish the board without its
        // traces.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainers_table_edit_keeps_the_submissions_KiCad_data()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();

            const string KiCadPath = AmendSubmissionFlowTests.SystemId + "/KiCad data/board.kicad_pcb";

            // Into the submission's FILE RECORDS, where the store rebuilds a payload's files from -
            // the fake's LoadPayloadAsync models the real store's join, not the saved object.
            store.Files[id].Add(new SubmissionFileRecord(99, KiCadPath, new string('c', 64), 10, IsUploaded: true));

            AmendOutcome outcome = await this.AmendAsync(store, id, await AmendSubmissionFlowTests.EditedRowsAsync(store, id));

            Assert.True(outcome.IsAmended, outcome.FullError);

            SubmissionManifest after = (await store.LoadPayloadAsync(id, CancellationToken.None))!;
            Assert.Contains(after.Files, file => file.Path == KiCadPath);
        }

        // ---- who, and when --------------------------------------------------------------------

        [Fact]
        public async Task A_maintainer_of_ANOTHER_board_cannot_amend_and_nothing_changes()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id);

            AmendOutcome outcome = await this.AmendAsync(store, id, rows, AmendSubmissionFlowTests.MaintainerOf("Commodore/C128/310378"));

            Assert.True(outcome.IsForbidden);
            Assert.Empty(store.Amendments);
            Assert.Equal("", (await store.LoadPayloadAsync(id, CancellationToken.None))!.Rows.Components[0].PartNumber ?? "");
        }

        [Theory]
        [InlineData(SubmissionState.Merged)]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.Uploading)]
        public async Task A_submission_that_is_decided_or_still_uploading_cannot_be_amended(string state)
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync(state);

            AmendOutcome outcome = await this.AmendAsync(store, id, await AmendSubmissionFlowTests.EditedRowsAsync(store, id));

            Assert.True(outcome.IsConflict);
            Assert.Empty(store.Amendments);
        }

        // Two maintainers at once: the second is told whose change they would overwrite.
        [Fact]
        public async Task An_amendment_made_after_the_maintainer_opened_it_is_refused_naming_who()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id);

            Assert.True((await this.AmendAsync(store, id, rows)).IsAmended);

            AmendOutcome second = await this.AmendAsync(store, id, rows, AmendSubmissionFlowTests.Administrator(), expectedVersion: 0);

            Assert.True(second.IsConflict);
            Assert.Contains("Anna", second.Error, StringComparison.Ordinal);
            Assert.Single(store.Amendments);
        }

        // ###########################################################################################
        // *** THE VERSION IS RE-CHECKED WHERE IT IS STORED (code review, 2026-09-25). *** The flow's
        // own check ran before the store's transaction, and the store never saw the version the
        // maintainer opened - so an amendment committed in between was silently overwritten, the
        // exact loss the version exists to prevent. Here the other maintainer's change lands in that
        // gap, after the flow has checked and before it saves.
        // ###########################################################################################
        [Fact]
        public async Task An_amendment_stored_between_the_version_check_and_the_save_is_not_overwritten()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id);
            SubmissionManifest theirs = (await store.LoadPayloadAsync(id, CancellationToken.None))!;

            store.BeforeAmend = async () =>
            {
                store.BeforeAmend = null;
                await store.AmendAsync(id, 0, theirs, false, 8, "Bo (bo@example.com)", AmendSubmissionFlowTests.Now, CancellationToken.None);
            };

            AmendOutcome outcome = await this.AmendAsync(store, id, rows);

            Assert.True(outcome.IsConflict, outcome.Error);
            Assert.Contains("Bo", outcome.Error, StringComparison.Ordinal);
            Assert.Equal("Bo (bo@example.com)", Assert.Single(store.Amendments).By);
        }

        // The same gap, closed by a decision rather than another amendment: a reject is not taken
        // under the publish lock, so only the store's own check can catch it.
        [Fact]
        public async Task A_submission_decided_between_the_state_check_and_the_save_is_not_amended()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id);

            store.BeforeAmend = async () =>
            {
                store.BeforeAmend = null;
                await store.SetStateAsync(id, SubmissionState.Rejected, AmendSubmissionFlowTests.Now, CancellationToken.None);
            };

            AmendOutcome outcome = await this.AmendAsync(store, id, rows);

            Assert.True(outcome.IsConflict, outcome.Error);
            Assert.Empty(store.Amendments);
            Assert.Equal("", (await store.LoadPayloadAsync(id, CancellationToken.None))!.Rows.Components[0].PartNumber ?? "");
        }

        // ###########################################################################################
        // *** AN AMENDMENT WAITS FOR A PUBLISH IN PROGRESS (code review, 2026-09-25). *** Without the
        // publish lock an amendment could land while an approval was writing the tree: the tree got
        // the rows as they were, the database the amended ones - and the contributor was mailed that
        // a maintainer's change had been published. Taking the lock, it waits, then finds the
        // submission decided.
        // ###########################################################################################
        [Fact]
        public async Task An_amendment_saved_while_a_publish_is_running_waits_for_it_and_then_finds_it_decided()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id);
            var publishLock = new PublishLock();

            Task<AmendOutcome> amending;

            using (await publishLock.EnterAsync(CancellationToken.None))
            {
                // A publish is under way while the maintainer saves...
                amending = this.AmendAsync(store, id, rows, publishLock: publishLock);

                // ...and finishes, recording the submission as published.
                await store.SetStateAsync(id, SubmissionState.Merged, AmendSubmissionFlowTests.Now, CancellationToken.None);
            }

            AmendOutcome outcome = await amending;

            Assert.True(outcome.IsConflict, outcome.Error);
            Assert.Empty(store.Amendments);
        }

        // ---- what it changes ------------------------------------------------------------------

        [Fact]
        public async Task The_edited_rows_become_the_submission_and_the_original_is_kept()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            var accounts = new FakeAccountStore();

            AmendOutcome outcome = await this.AmendAsync(store, id, await AmendSubmissionFlowTests.EditedRowsAsync(store, id), accounts: accounts);

            Assert.True(outcome.IsAmended, outcome.Error + " " + string.Join(" | ", outcome.Findings.Select(f => f.Code + ": " + f.Message)));
            Assert.Equal(1, outcome.Version);
            Assert.Equal("251715-01", (await store.LoadPayloadAsync(id, CancellationToken.None))!.Rows.Components[0].PartNumber);

            var kept = Assert.Single(store.Amendments);
            Assert.True(string.IsNullOrEmpty(kept.Replaced.Rows.Components[0].PartNumber));
            Assert.Equal("Anna (anna@example.com)", kept.By);

            AuditEntry audit = Assert.Single(accounts.Audit);
            Assert.Equal(AmendSubmissionFlow.AmendedAction, audit.Action);
            Assert.Equal($"#{id}", audit.Subject);
        }

        // The table does not show highlights, calibrations or the revision date, so its route cannot
        // change them - whatever a hand-made request carries.
        [Fact]
        public async Task Highlights_calibrations_and_the_revision_date_stay_as_submitted()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id, edit: rows =>
            {
                rows.ComponentHighlights.Clear();
                rows.RevisionDate = "1999-January-01";
                rows.KiCadCalibrations.Add(new KiCadCalibrationEntry());
            });

            Assert.True((await this.AmendAsync(store, id, rows)).IsAmended);

            SubmissionRows stored = (await store.LoadPayloadAsync(id, CancellationToken.None))!.Rows;
            Assert.Single(stored.ComponentHighlights);
            Assert.Equal("2026-September-25", stored.RevisionDate);
            Assert.Empty(stored.KiCadCalibrations);
        }

        // Files follow the rows: a deleted row drops the file it cited.
        // ###########################################################################################
        // *** A MAINTAINER CAN DELETE A SCHEMATIC THAT HAS HIGHLIGHTS (code review, 2026-09-25). ***
        // The highlights are not in the table, so the save used to be refused on
        // highlight.unknown_schematic with nothing the maintainer could do about it. They now go with
        // their schematic (SubmissionRowsBoard.WithTableSections), and so does its image file.
        // ###########################################################################################
        [Fact]
        public async Task Deleting_a_schematic_with_highlights_is_saved_and_its_highlights_go_with_it()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id, edit: rows => rows.Schematics.Clear());

            AmendOutcome outcome = await this.AmendAsync(store, id, rows);

            Assert.True(outcome.IsAmended, outcome.Error + " " + string.Join(" ", outcome.Findings.Select(f => f.Message)));

            SubmissionManifest stored = (await store.LoadPayloadAsync(id, CancellationToken.None))!;
            Assert.Empty(stored.Rows.Schematics);
            Assert.Empty(stored.Rows.ComponentHighlights);
            Assert.DoesNotContain(stored.Files, file => file.Path == AmendSubmissionFlowTests.SheetImage);
        }

        [Fact]
        public async Task A_deleted_row_drops_the_file_it_cited()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id, edit: rows => rows.BoardLocalFiles.Clear());

            Assert.True((await this.AmendAsync(store, id, rows)).IsAmended);

            Assert.Equal(
                [AmendSubmissionFlowTests.SheetImage],
                (await store.LoadPayloadAsync(id, CancellationToken.None))!.Files.Select(file => file.Path));
        }

        // A maintainer may point a row at a file that is already PUBLISHED - the server takes its own
        // copy, the way a new submission's unchanged files are taken.
        [Fact]
        public async Task A_row_may_cite_a_file_already_published_and_the_server_takes_its_own_copy()
        {
            const string sheet = "Commodore/C64/250407/sheet1.png";
            DataTreeBuilder.Files(this.thisData, sheet);

            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id, edit: rows =>
                rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = "Sheet", File = sheet }));

            AmendOutcome outcome = await this.AmendAsync(store, id, rows);

            Assert.True(outcome.IsAmended, outcome.Error + " " + string.Join(" ", outcome.Findings.Select(f => f.Message)));
            SubmissionFile added = Assert.Single((await store.LoadPayloadAsync(id, CancellationToken.None))!.Files, file => file.Path == sheet);
            Assert.True(this.Blobs().Contains(added.Sha256));
        }

        // A maintainer edits rows; they cannot bring in a file nobody has sent.
        [Fact]
        public async Task A_row_citing_a_file_nobody_has_is_refused_and_nothing_changes()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();
            SubmissionRows rows = await AmendSubmissionFlowTests.EditedRowsAsync(store, id, edit: rows =>
                rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = "X", File = "Commodore/C64/250407/nowhere.pdf" }));

            AmendOutcome outcome = await this.AmendAsync(store, id, rows);

            Assert.False(outcome.IsAmended);
            Assert.Contains(outcome.Findings, finding => finding.Code == "amend.file_unknown");

            // Said to the maintainer about "the Maintainer tab", not the application it replaced.
            ValidationFinding unknown = outcome.Findings.First(finding => finding.Code == "amend.file_unknown");
            Assert.Contains("the Maintainer tab", unknown.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("CRT Maintainer", unknown.Message, StringComparison.Ordinal);
            Assert.Empty(store.Amendments);
        }

        // An approval given before the change was given to other content.
        [Fact]
        public async Task Amending_clears_the_approvals_given_and_an_approved_submission_waits_again()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync(SubmissionState.Approved);
            await store.AddApprovalAsync(id, ApproverRole.Maintainer, 7, "Anna", AmendSubmissionFlowTests.Now, CancellationToken.None);

            Assert.True((await this.AmendAsync(store, id, await AmendSubmissionFlowTests.EditedRowsAsync(store, id))).IsAmended);

            Assert.Empty(await store.GetApprovalsAsync(id, CancellationToken.None));
            Assert.Equal(SubmissionState.Pending, (await store.FindAsync(id, CancellationToken.None))!.State);
        }

        [Fact]
        public async Task The_latest_amendment_names_who_made_it()
        {
            (FakeSubmissionStore store, long id) = await AmendSubmissionFlowTests.PendingAsync();

            Assert.Null(await store.GetLatestAmendmentAsync(id, CancellationToken.None));

            await this.AmendAsync(store, id, await AmendSubmissionFlowTests.EditedRowsAsync(store, id));

            SubmissionAmendment? latest = await store.GetLatestAmendmentAsync(id, CancellationToken.None);
            Assert.Equal(1, latest!.Version);
            Assert.Equal("Anna (anna@example.com)", latest.By);
        }
    }
}
