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
    // BoardEditFlow - a board's Board data on the Boards screen, and a maintainer's edit of it,
    // which goes STRAIGHT TO BETA (owner decision, 2026-10-03: "it should go directly to the next
    // queue, 'BETA > Stable', so it can directly be tested in BETA") as a submission from their
    // account that the ordinary approval publishes at once. Not while the board waits in BETA
    // (owner decision, same day: "it should simply disallow it, even if this is coming from a
    // maintainer"). Against a real temp tree, the real create/finalise path and the real
    // ApprovePublishFlow; the stores are the fakes.
    //
    // The refusals come first: who may change a board, and when, are what matter most.
    // ###########################################################################################
    [Collection("BoardFiles")]
    public sealed class BoardEditFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 3, 12, 0, 0, TimeSpan.Zero);
        private const string BoardId = "Commodore/C64/250407";
        private const string Manual = "Commodore/C64/250407/manual.pdf";
        private const string Shared = "Commodore/Shared files/74LS08.pdf";
        private const string Sheet = "Commodore/C64/250407/sheet-1.png";

        private readonly string thisRoot = Path.Combine(Path.GetTempPath(), "crt-board-edit", Guid.NewGuid().ToString("N"));
        private readonly string thisData;
        private readonly string thisBlobs;

        public BoardEditFlowTests()
        {
            this.thisData = Path.Combine(this.thisRoot, "Data");
            this.thisBlobs = Path.Combine(this.thisRoot, "blobs");
            Directory.CreateDirectory(this.thisData);
            Directory.CreateDirectory(this.thisBlobs);

            // BETA's board: one schematic and one component, citing the manual and a shared datasheet,
            // with a highlight and a KiCad calibration in its highlight file.
            BoardWorkbookWriter.Write(this.WorkbookPath, new BoardData
            {
                RevisionDate = "2026-August-21",
                Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = BoardEditFlowTests.Sheet }],
                Components = [new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-01" }],
                BoardLocalFiles =
                [
                    new BoardLocalFileEntry { Category = "Service", Name = "Manual", File = BoardEditFlowTests.Manual },
                    new BoardLocalFileEntry { Category = "Service", Name = "74LS08", File = BoardEditFlowTests.Shared }
                ]
            });

            foreach (string path in new[] { BoardEditFlowTests.Manual, BoardEditFlowTests.Shared })
                File.WriteAllText(DataTreeBuilder.Full(this.thisData, path), "%PDF-1.4\n% " + path + "\n");

            File.WriteAllBytes(
                DataTreeBuilder.Full(this.thisData, BoardEditFlowTests.Sheet),
                [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, .. new byte[32]]);

            BoardSidecarWriter.Write(
                this.WorkbookPath,
                [new ComponentHighlightEntry { SchematicName = "Sheet 1", BoardLabel = "U8", X = "1", Y = "2", Width = "3", Height = "4" }],
                [new KiCadCalibrationEntry { SchematicName = "Sheet 1", CadName = "board", OffsetX = 1, OffsetY = 2, ScaleX = 1, ScaleY = 1 }]);
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

        private string WorkbookPath => DataTreeBuilder.Full(this.thisData, DataTreeBuilder.Workbook);

        private BlobStore Blobs() => new(this.thisBlobs, NullLogger<BlobStore>.Instance);

        private static PublishedBoardReader Boards() => new(NullLogger<PublishedBoardReader>.Instance);

        // The ordinary approval - with a stable source configured or not, which is what decides
        // whether one submission in BETA per board applies (its step 3b).
        private ApprovePublishFlow Approvals(ISubmissionStore store, bool withStable) =>
            new(
                new PublishExecutor(this.Blobs(), store, NullLogger<PublishExecutor>.Instance),
                BoardEditFlowTests.Boards(),
                store,
                new FakeAccountStore(),
                NullLogger<ApprovePublishFlow>.Instance,
                publishLock: null,
                options: withStable
                    ? new ServerOptions
                    {
                        ProductionDataTreeRoot = "production",
                        ProductionManifestPath = "production/dataChecksums.json",
                        ProductionPublicDataBaseUrl = "https://example.com/app-data/Data"
                    }
                    : null);

        private static AccountRecord Account(long id = 7, bool administrator = false) =>
            new(id, $"anna{id}@example.com", $"anna{id}@example.com", "hash", "Anna",
                IsVerified: true, IsAdministrator: administrator, IsLocked: false, BoardEditFlowTests.Now, null);

        private static ReviewAccess MaintainerOf(params string[] boards) =>
            ReviewAccess.For(BoardEditFlowTests.Account(), boards);

        private Task<BoardTableOutcome> ReadAsync(FakeSubmissionStore store, ReviewAccess? access = null, bool withStable = true) =>
            BoardEditFlow.ReadTableAsync(
                access ?? BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId),
                BoardEditFlowTests.BoardId,
                this.thisData,
                store,
                BoardEditFlowTests.Boards(),
                "https://example.com/data",
                withStable,
                CancellationToken.None);

        private Task<BoardEditOutcome> CheckAsync(FakeSubmissionStore store, BoardEditRequest request, ReviewAccess? access = null) =>
            BoardEditFlow.CheckAsync(
                access ?? BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId),
                request,
                this.thisData,
                store,
                BoardEditFlowTests.Boards(),
                oneSubmissionInBeta: true,
                BoardEditFlowTests.Now,
                CancellationToken.None);

        // `checkedBy` is what the flow is told about the stable source; `approvedWith` what the
        // approval is - the same server setting, apart only to stand in for a change between the two.
        private Task<BoardEditOutcome> PublishAsync(
            FakeSubmissionStore store,
            BoardEditRequest request,
            ReviewAccess? access = null,
            bool withStable = true,
            bool? approvedWithStable = null) =>
            BoardEditFlow.PublishAsync(
                access ?? BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId),
                request,
                this.thisData,
                store,
                this.Blobs(),
                BoardEditFlowTests.Boards(),
                this.Approvals(store, approvedWithStable ?? withStable),
                withStable,
                "192.0.2.7",
                BoardEditFlowTests.Now,
                CancellationToken.None);

        // The table as read, with U8's part number changed - what the Maintainer tab sends after one
        // edit, with the reason asked for and the (empty) list of removals the check answered.
        private async Task<BoardEditRequest> EditedAsync(
            FakeSubmissionStore store,
            Action<SubmissionRows>? edit = null,
            string reason = "Corrected U8's part number.",
            IReadOnlyList<string>? removals = null)
        {
            BoardTableAnswer table = (await this.ReadAsync(store, withStable: false)).Answer!;

            SubmissionRows rows = SubmissionRowsBoard.FromBoard(SubmissionRowsBoard.ToBoard(table.Rows));
            rows.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA", TechnicalNameOrValue = "906114-01", PartNumber = "251715-02" };
            edit?.Invoke(rows);

            return new BoardEditRequest(BoardEditFlowTests.BoardId, table.Fingerprint, reason, rows, removals ?? []);
        }

        private async Task<BoardData> BetaBoardAsync() =>
            await BoardEditFlowTests.Boards().TryReadBoardAsync(this.thisData, BoardEditFlowTests.BoardId) ?? throw new InvalidOperationException();

        // ###########################################################################################
        // A maintainer of ANOTHER board reads it - every maintainer may - but is told it cannot be
        // changed from here, and a change sent anyway is refused with nothing made.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainer_of_another_board_may_read_its_table_but_not_change_it()
        {
            var store = new FakeSubmissionStore();
            ReviewAccess other = BoardEditFlowTests.MaintainerOf("Commodore/C128/310378");

            BoardTableOutcome table = await this.ReadAsync(store, other);

            Assert.NotNull(table.Answer);
            Assert.False(table.Answer!.MayEdit);
            Assert.Equal(BoardEditFlow.NotYoursMessage, table.Answer.MayNotEditReason);

            BoardEditRequest request = await this.EditedAsync(store);

            Assert.True((await this.CheckAsync(store, request, other)).IsForbidden);
            Assert.True((await this.PublishAsync(store, request, other)).IsForbidden);
            Assert.Empty(store.Submissions);
        }

        [Fact]
        public async Task An_account_that_may_review_nothing_is_refused_the_table_the_check_and_a_change()
        {
            var store = new FakeSubmissionStore();
            BoardEditRequest request = await this.EditedAsync(store);
            ReviewAccess nobody = ReviewAccess.For(BoardEditFlowTests.Account());

            Assert.True((await this.ReadAsync(store, nobody)).IsForbidden);
            Assert.True((await this.CheckAsync(store, request, nobody)).IsForbidden);
            Assert.True((await this.PublishAsync(store, request, nobody)).IsForbidden);
            Assert.Empty(store.Submissions);
        }

        // A board the administrator has closed to contributions takes no change from here either.
        [Fact]
        public async Task A_board_closed_to_contributions_cannot_be_changed()
        {
            var store = new FakeSubmissionStore();
            BoardEditRequest request = await this.EditedAsync(store);

            store.Boards[BoardEditFlowTests.BoardId] = new NewSubmission(
                BoardEditFlowTests.BoardId, "Commodore", "C64", "250407", null, "x@example.com", null, "hash", "r0", "", 1, [],
                BoardEditFlowTests.Now, BoardEditFlowTests.Now.AddHours(24));
            store.ClosedBoards.Add(BoardEditFlowTests.BoardId);

            BoardTableAnswer table = (await this.ReadAsync(store)).Answer!;
            Assert.False(table.MayEdit);
            Assert.Equal(BoardEditFlow.ClosedMessage, table.MayNotEditReason);

            Assert.True((await this.PublishAsync(store, request)).IsForbidden);
        }

        // ###########################################################################################
        // *** BETA PUBLISHED AGAIN AFTER THE TABLE WAS OPENED. *** Publishing the table as it was
        // read would put that publish's change back as it was - a board's rows are replaced whole.
        // Refused at the check and at the publish, nothing made, BETA as the other publish left it.
        // ###########################################################################################
        [Fact]
        public async Task A_change_made_on_a_table_read_before_BETA_changed_is_refused()
        {
            var store = new FakeSubmissionStore();
            BoardEditRequest request = await this.EditedAsync(store);

            BoardData beta = await this.BetaBoardAsync();
            beta.Components[0] = new ComponentEntry { BoardLabel = "U8", FriendlyName = "PLA (changed meanwhile)", TechnicalNameOrValue = "906114-01" };
            BoardWorkbookWriter.Write(this.WorkbookPath, beta);

            BoardEditOutcome check = await this.CheckAsync(store, request);
            BoardEditOutcome sent = await this.PublishAsync(store, request);

            Assert.True(check.IsConflict);
            Assert.True(sent.IsConflict);
            Assert.Equal(BoardEditFlow.ChangedSinceMessage, sent.Error);
            Assert.Empty(store.Submissions);
            Assert.Equal("PLA (changed meanwhile)", (await this.BetaBoardAsync()).Components[0].FriendlyName);
        }

        [Fact]
        public async Task A_table_that_changes_nothing_is_not_published()
        {
            var store = new FakeSubmissionStore();
            BoardTableAnswer table = (await this.ReadAsync(store)).Answer!;
            var request = new BoardEditRequest(BoardEditFlowTests.BoardId, table.Fingerprint, "Nothing.", table.Rows, []);

            BoardEditOutcome check = await this.CheckAsync(store, request);
            BoardEditOutcome sent = await this.PublishAsync(store, request);

            Assert.False(check.IsAccepted);
            Assert.False(sent.IsAccepted);
            Assert.Equal(BoardEditFlow.NothingChangedMessage, sent.Error);
            Assert.Empty(store.Submissions);
        }

        // A maintainer edits rows here, and cannot bring in a file nobody has sent.
        [Fact]
        public async Task A_row_citing_a_file_BETA_does_not_hold_is_refused()
        {
            var store = new FakeSubmissionStore();
            BoardEditRequest request = await this.EditedAsync(store, rows =>
                rows.BoardLocalFiles.Add(new BoardLocalFileEntry { Category = "Service", Name = "New", File = "Commodore/C64/250407/new.pdf" }));

            BoardEditOutcome check = await this.CheckAsync(store, request);
            BoardEditOutcome sent = await this.PublishAsync(store, request);

            Assert.False(check.IsAccepted);
            Assert.False(sent.IsAccepted);
            Assert.Contains(sent.Findings, finding => finding.Code == "board_edit.file_unknown" && finding.Subject == "Commodore/C64/250407/new.pdf");
            Assert.Empty(store.Submissions);
        }

        [Fact]
        public async Task A_board_BETA_holds_nothing_of_has_no_table()
        {
            var store = new FakeSubmissionStore();

            BoardTableOutcome table = await BoardEditFlow.ReadTableAsync(
                BoardEditFlowTests.MaintainerOf("Commodore/VIC-20/250403"),
                "Commodore/VIC-20/250403",
                this.thisData,
                store,
                BoardEditFlowTests.Boards(),
                null,
                oneSubmissionInBeta: true,
                CancellationToken.None);

            Assert.True(table.IsNotFound);
            Assert.Equal(BoardEditFlow.NotInBetaMessage, table.Error);
        }

        // ###########################################################################################
        // *** THE CHANGE IS IN BETA. *** One publish: BETA's workbook carries the changed row, at a
        // revision the server stamped; the submission it became is MERGED - from the maintainer's
        // account, with the reason as its description, approved and decided by that same account -
        // so BETA > Stable, the board's history and a push-back all see an ordinary submission.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainers_change_is_published_to_BETA_as_a_merged_submission_from_their_account()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome sent = await this.PublishAsync(store, await this.EditedAsync(store));

            Assert.True(sent.IsAccepted, sent.FullError);
            Assert.True(sent.IsPublished, sent.NotPublishedReason);
            Assert.Equal(BoardWorkbookStyle.FormatRevisionDate(BoardEditFlowTests.Now), sent.Revision);
            Assert.Empty(sent.Removals);

            Assert.Equal("251715-02", (await this.BetaBoardAsync()).Components.Single().PartNumber);

            SubmissionRecord record = store.Submissions[sent.SubmissionId];
            Assert.Equal(SubmissionState.Merged, record.State);
            Assert.Equal(7, record.AccountId);
            Assert.Equal(BoardEditFlowTests.BoardId, record.BoardId);
            Assert.Equal("Corrected U8's part number.", record.Summary);
            Assert.Equal(7, store.Decisions[sent.SubmissionId].DecidedByAccountId);

            SubmissionManifest payload = (await store.LoadPayloadAsync(sent.SubmissionId, CancellationToken.None))!;
            Assert.Equal("2026-August-21", payload.BaseRevision);
            Assert.Equal(
                new[] { BoardEditFlowTests.Shared, BoardEditFlowTests.Manual, BoardEditFlowTests.Sheet }.Order(StringComparer.Ordinal),
                payload.Files.Select(file => file.Path).Order(StringComparer.Ordinal));
        }

        // ###########################################################################################
        // What the table does not show stays as BETA has it: the highlights and the KiCad calibrations
        // go into the submission - and so into BETA - unchanged. Without them the publish would write
        // the board with its calibration work removed (PublishMerge.CalibrationsOf).
        // ###########################################################################################
        [Fact]
        public async Task Highlights_and_calibrations_stay_as_BETA_has_them()
        {
            var store = new FakeSubmissionStore();
            BoardEditRequest request = await this.EditedAsync(store, rows =>
            {
                rows.ComponentHighlights.Clear();
                rows.KiCadCalibrations.Clear();
            });

            BoardEditOutcome sent = await this.PublishAsync(store, request);

            Assert.True(sent.IsPublished, sent.FullError + sent.NotPublishedReason);

            SubmissionManifest payload = (await store.LoadPayloadAsync(sent.SubmissionId, CancellationToken.None))!;
            Assert.Equal("U8", Assert.Single(payload.Rows.ComponentHighlights).BoardLabel);
            Assert.Equal("board", Assert.Single(payload.Rows.KiCadCalibrations).CadName);
            Assert.Equal("U8", Assert.Single((await this.BetaBoardAsync()).ComponentHighlights).BoardLabel);
        }

        // The reason goes with the change, as a contribution's description does - none, nothing made.
        [Fact]
        public async Task A_change_without_a_reason_is_not_published()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome sent = await this.PublishAsync(store, await this.EditedAsync(store, reason: "   "));

            Assert.False(sent.IsAccepted);
            Assert.Equal(BoardEditFlow.NoReasonMessage, sent.Error);
            Assert.Empty(store.Submissions);
            Assert.Equal("251715-01", (await this.BetaBoardAsync()).Components.Single().PartNumber);
        }

        // ###########################################################################################
        // *** WHAT IT REMOVES IS SHOWN FIRST. *** Dropping the only row citing the manual makes the
        // manual a file nothing uses, which the publish removes (no orphan files). The check names
        // it; a publish sent with any other list is refused with nothing made - the approval's own
        // rule - and one sent with that list publishes and removes exactly it.
        // ###########################################################################################
        [Fact]
        public async Task The_check_names_what_the_change_removes_and_only_that_list_is_published()
        {
            var store = new FakeSubmissionStore();
            string manual = DataTreeBuilder.Full(this.thisData, BoardEditFlowTests.Manual);

            void DropTheManual(SubmissionRows rows) => rows.BoardLocalFiles.RemoveAll(file => file.Name == "Manual");

            BoardEditOutcome check = await this.CheckAsync(store, await this.EditedAsync(store, DropTheManual));

            Assert.True(check.IsAccepted, check.FullError);
            Assert.Equal([BoardEditFlowTests.Manual], check.Removals);
            Assert.Empty(store.Submissions);

            BoardEditOutcome unshown = await this.PublishAsync(store, await this.EditedAsync(store, DropTheManual, removals: []));

            Assert.True(unshown.IsConflict);
            Assert.Equal(BoardEditFlow.RemovalsChangedMessage, unshown.Error);
            Assert.Empty(store.Submissions);
            Assert.True(File.Exists(manual));

            BoardEditOutcome sent = await this.PublishAsync(
                store, await this.EditedAsync(store, DropTheManual, removals: [BoardEditFlowTests.Manual]));

            Assert.True(sent.IsPublished, sent.FullError + sent.NotPublishedReason);
            Assert.Equal([BoardEditFlowTests.Manual], sent.Removals);
            Assert.False(File.Exists(manual));
        }

        // An ordinary edit removes nothing, and its check says so.
        [Fact]
        public async Task An_edit_that_drops_no_file_removes_nothing()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome check = await this.CheckAsync(store, await this.EditedAsync(store));

            Assert.True(check.IsAccepted, check.FullError);
            Assert.Empty(check.Removals);
        }

        // ###########################################################################################
        // *** WHILE THE BOARD WAITS IN BETA > STABLE, NO CHANGE IS MADE TO IT - a maintainer's
        // included (owner decision, 2026-10-03). *** After one change is published, the board waits
        // for stable: the table says why it is read-only, and a second change - checked or sent - is
        // refused with nothing made and BETA as the first change left it.
        // ###########################################################################################
        [Fact]
        public async Task While_the_board_waits_in_BETA_no_further_change_is_made_to_it()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome first = await this.PublishAsync(store, await this.EditedAsync(store));
            Assert.True(first.IsPublished, first.FullError + first.NotPublishedReason);

            BoardTableAnswer table = (await this.ReadAsync(store)).Answer!;
            Assert.False(table.MayEdit);
            Assert.Equal(OneSubmissionInBeta.NoChangeMessage(BoardEditFlowTests.BoardId), table.MayNotEditReason);

            var second = new BoardEditRequest(
                BoardEditFlowTests.BoardId,
                table.Fingerprint,
                "U8 is the CPU.",
                BoardEditFlowTests.WithFriendlyName(SubmissionRowsBoard.FromBoard(SubmissionRowsBoard.ToBoard(table.Rows)), "CPU"),
                []);

            BoardEditOutcome check = await this.CheckAsync(store, second);
            BoardEditOutcome sent = await this.PublishAsync(store, second);

            Assert.True(check.IsForbidden);
            Assert.True(sent.IsForbidden);
            Assert.Equal(OneSubmissionInBeta.NoChangeMessage(BoardEditFlowTests.BoardId), sent.Error);
            Assert.Single(store.Submissions);
            Assert.Equal("PLA", (await this.BetaBoardAsync()).Components.Single().FriendlyName);
        }

        private static SubmissionRows WithFriendlyName(SubmissionRows rows, string name)
        {
            ComponentEntry u8 = rows.Components[0];
            rows.Components[0] = new ComponentEntry
            {
                BoardLabel = u8.BoardLabel,
                FriendlyName = name,
                TechnicalNameOrValue = u8.TechnicalNameOrValue,
                PartNumber = u8.PartNumber
            };

            return rows;
        }

        // Without a stable source nothing ever leaves BETA, so the rule is off - as for an approval -
        // and every change is published.
        [Fact]
        public async Task Without_a_stable_source_a_second_change_is_published_too()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome first = await this.PublishAsync(store, await this.EditedAsync(store), withStable: false);
            Assert.True(first.IsPublished, first.FullError + first.NotPublishedReason);

            Assert.True((await this.ReadAsync(store, withStable: false)).Answer!.MayEdit);

            BoardEditOutcome second = await this.PublishAsync(
                store,
                await this.EditedAsync(store, rows => BoardEditFlowTests.WithFriendlyName(rows, "CPU"), reason: "U8 is the CPU."),
                withStable: false);

            Assert.True(second.IsPublished, second.FullError + second.NotPublishedReason);
            Assert.Equal("CPU", (await this.BetaBoardAsync()).Components.Single().FriendlyName);
        }

        // ###########################################################################################
        // *** A PUBLISH THE APPROVAL REFUSES AFTER ALL LOSES NOTHING. *** Something changed between
        // the checks and the approval's lock - here the board began waiting in BETA - so the
        // submission is made but not published: it stays PENDING under Contributor Submissions, the
        // reason is said, and BETA is untouched.
        // ###########################################################################################
        [Fact]
        public async Task A_publish_the_approval_refuses_leaves_the_change_waiting_in_the_queue()
        {
            var store = new FakeSubmissionStore();

            BoardEditOutcome first = await this.PublishAsync(store, await this.EditedAsync(store), withStable: false);
            Assert.True(first.IsPublished, first.FullError + first.NotPublishedReason);

            BoardEditOutcome second = await this.PublishAsync(
                store,
                await this.EditedAsync(store, rows => BoardEditFlowTests.WithFriendlyName(rows, "CPU"), reason: "U8 is the CPU."),
                withStable: false,
                approvedWithStable: true);

            Assert.True(second.IsAccepted, second.FullError);
            Assert.False(second.IsPublished);
            Assert.Equal(OneSubmissionInBeta.BusyMessage(BoardEditFlowTests.BoardId), second.NotPublishedReason);
            Assert.Equal(SubmissionState.Pending, store.Submissions[second.SubmissionId].State);
            Assert.Equal("PLA", (await this.BetaBoardAsync()).Components.Single().FriendlyName);
        }

        // ###########################################################################################
        // *** A CHANGE WHILE THE ACCOUNT'S EARLIER ONE STILL WAITS IN THE QUEUE. *** It would be built
        // on BETA, which does not hold the earlier one, and as a newer submission from the same
        // account it would REPLACE it - quietly withdrawing that change. The table says so and opens
        // read-only; a send is refused and the earlier submission is left as it was. Another
        // contributor's waiting submission does not stand in the way.
        // ###########################################################################################
        [Fact]
        public async Task A_change_while_the_accounts_earlier_change_still_waits_is_refused_and_the_earlier_one_kept()
        {
            var store = new FakeSubmissionStore();
            long earlier = await BoardEditFlowTests.PendingFromAsync(store, accountId: 7);

            BoardTableAnswer table = (await this.ReadAsync(store)).Answer!;
            Assert.False(table.MayEdit);
            Assert.Equal(BoardEditWording.AlreadyWaitingMessage(earlier), table.MayNotEditReason);

            BoardEditOutcome sent = await this.PublishAsync(store, await this.EditedAsync(store));

            Assert.True(sent.IsForbidden);
            Assert.Single(store.Submissions);
            Assert.Equal(SubmissionState.Pending, store.Submissions[earlier].State);

            // Another account's waiting submission is not this account's to wait for.
            Assert.True((await this.ReadAsync(store, ReviewAccess.For(
                BoardEditFlowTests.Account(id: 8), [BoardEditFlowTests.BoardId]))).Answer!.MayEdit);
        }

        // A pending submission of the board from `accountId`, made straight in the store.
        private static async Task<long> PendingFromAsync(FakeSubmissionStore store, long accountId)
        {
            long id = await store.CreateAsync(
                new NewSubmission(
                    BoardEditFlowTests.BoardId, "Commodore", "C64", "250407",
                    accountId, null, "192.0.2.2", "hash", "2026-August-21",
                    "An earlier change.", 1, [], BoardEditFlowTests.Now, BoardEditFlowTests.Now.AddHours(24)),
                CancellationToken.None);

            await store.SetStateAsync(id, SubmissionState.Pending, BoardEditFlowTests.Now, CancellationToken.None);

            return id;
        }

        // ###########################################################################################
        // THE STABLE SOURCE'S BOARD (owner request, 2026-10-04: a BETA / Stable switch on Board data):
        // read for any maintainer, NEVER editable - not by its own maintainers, not by the
        // administrator - and opened from the stable source's public address.
        // ###########################################################################################
        [Fact]
        public async Task The_stable_sources_board_is_read_and_never_editable()
        {
            foreach (ReviewAccess access in new[] { BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId), BoardEditFlowTests.MaintainerOf("Commodore/C128/310378") })
            {
                BoardTableOutcome outcome = await BoardEditFlow.ReadStableTableAsync(
                    access, BoardEditFlowTests.BoardId, this.thisData, BoardEditFlowTests.Boards(), "https://example.com/app-data/Data");

                BoardTableAnswer answer = outcome.Answer!;
                Assert.False(answer.MayEdit);
                Assert.Equal(BoardEditFlow.StableReadOnlyMessage, answer.MayNotEditReason);
                Assert.Equal("https://example.com/app-data/Data", answer.ProductionDataUrl);
                Assert.Null(answer.BetaDataUrl);
                Assert.Equal("U8", Assert.Single(answer.Rows.Components).BoardLabel);
            }
        }

        [Fact]
        public async Task A_stable_source_without_the_board_or_none_at_all_says_so()
        {
            string empty = Path.Combine(this.thisRoot, "stable-empty");
            Directory.CreateDirectory(empty);

            BoardTableOutcome notThere = await BoardEditFlow.ReadStableTableAsync(
                BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId), BoardEditFlowTests.BoardId, empty, BoardEditFlowTests.Boards(), null);

            BoardTableOutcome noStable = await BoardEditFlow.ReadStableTableAsync(
                BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId), BoardEditFlowTests.BoardId, null, BoardEditFlowTests.Boards(), null);

            // And somebody who may review nothing is refused before anything is read.
            BoardTableOutcome refused = await BoardEditFlow.ReadStableTableAsync(
                BoardEditFlowTests.MaintainerOf(), BoardEditFlowTests.BoardId, this.thisData, BoardEditFlowTests.Boards(), null);

            Assert.True(refused.IsForbidden);

            Assert.True(notThere.IsNotFound);
            Assert.Equal(BoardEditFlow.NotInStableMessage, notThere.Error);
            Assert.True(noStable.IsNotFound);
            Assert.Equal(BoardEditFlow.NoStableSourceMessage, noStable.Error);
        }

        // ###########################################################################################
        // *** ONE STABLE ROOT FOR EVERY READER (code review, 2026-10-04). *** The overview judged
        // "in the stable source" - which offers CRT's "Data: Stable" switch - from the promotion's
        // root or else ProductionTreeRoot, while the stable table and files asked for publishing to
        // be switched on, so on a server with only ProductionTreeRoot every pick answered 404. All of
        // them read ServerOptions.StableSourceRoot now; reading needs no publishing.
        // ###########################################################################################
        [Fact]
        public async Task A_server_with_only_ProductionTreeRoot_reads_the_stable_board_from_there()
        {
            var options = new ServerOptions { ProductionTreeRoot = this.thisData };

            Assert.False(options.IsProductionPublishingConfigured);

            BoardTableOutcome outcome = await BoardEditFlow.ReadStableTableAsync(
                BoardEditFlowTests.MaintainerOf(BoardEditFlowTests.BoardId),
                BoardEditFlowTests.BoardId,
                options.StableSourceRoot,
                BoardEditFlowTests.Boards(),
                options.ProductionPublicDataBaseUrl);

            Assert.Equal("U8", Assert.Single(outcome.Answer!.Rows.Components).BoardLabel);
        }

        [Fact]
        public void The_stable_source_is_the_promotions_root_first_then_ProductionTreeRoot_else_none()
        {
            string promotion = Path.Combine(Path.GetTempPath(), "app-data", "Data");
            string guard = Path.Combine(Path.GetTempPath(), "app-data");

            Assert.Equal(promotion, new ServerOptions { ProductionDataTreeRoot = promotion, ProductionTreeRoot = guard }.StableSourceRoot);
            Assert.Equal(guard, new ServerOptions { ProductionDataTreeRoot = " ", ProductionTreeRoot = guard }.StableSourceRoot);
            Assert.Null(new ServerOptions { ProductionTreeRoot = " " }.StableSourceRoot);
            Assert.Null(new ServerOptions().StableSourceRoot);
        }

        // The same board read twice gives the same fingerprint - or no table could ever be sent.
        [Fact]
        public async Task The_same_board_read_twice_gives_the_same_fingerprint()
        {
            var store = new FakeSubmissionStore();

            string first = (await this.ReadAsync(store)).Answer!.Fingerprint;
            string second = (await this.ReadAsync(store)).Answer!.Fingerprint;

            Assert.Equal(first, second);
            Assert.Equal(64, first.Length);
        }

        // ###########################################################################################
        // The refusals CRT shows as they come name the Maintainer tab's screens as its tab strip does
        // (owner request, 2026-10-04: "Queue: Contributor submissions" and "Queue: Awaiting push from
        // BETA to stable") - quoted, since bare they run into the sentence. Written out, so the shared
        // words changing, or a refusal going back to a name of its own, fails here.
        // ###########################################################################################
        [Fact]
        public void The_refusals_name_the_maintainer_tabs_screens_as_its_tab_strip_does()
        {
            const string ContributorQueue = "\"Queue: Contributor submissions\"";
            const string BetaQueue = "\"Queue: Awaiting push from BETA to stable\"";

            Assert.Contains($"still waiting under {ContributorQueue}.", BoardEditWording.AlreadyWaitingMessage(57), StringComparison.Ordinal);
            // A board BETA lacks no longer claims its first submission is in the queue (owner report,
            // 2026-10-04: said also for one whose only submission was turned down long ago) - the
            // Maintainer tab's stage line says where it is.
            Assert.DoesNotContain("Queue", BoardEditFlow.NotInBetaMessage, StringComparison.Ordinal);
            Assert.Contains($"into BETA, {BetaQueue} and the board's history", BoardEditFlow.NoReasonMessage, StringComparison.Ordinal);
            Assert.Contains($"to the queue (under {BetaQueue}) first.", OneSubmissionInBeta.BusyMessage(BoardEditFlowTests.BoardId), StringComparison.Ordinal);
            Assert.Contains($"rejected under {BetaQueue} first.", OneSubmissionInBeta.NoChangeMessage(BoardEditFlowTests.BoardId), StringComparison.Ordinal);

            Assert.Equal(ContributorQueue, MaintainerScreenWording.ContributorQueueQuoted);
            Assert.Equal(BetaQueue, MaintainerScreenWording.BetaQueueQuoted);
        }
    }
}
