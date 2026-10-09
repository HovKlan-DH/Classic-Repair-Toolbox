using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // BoardListingFlow - placing a NEW board in CRT's drop-down lists from the Boards screen
    // (owner request, 2026-09-27): "The maintainer should order the new system, so it becomes visible
    // in the right location for the drop-down lists. This must be done before it can be pushed to
    // BETA." The list is BETA's newest main Excel data file; a real one is written for every test.
    // ###########################################################################################
    public sealed class BoardListingFlowTests : IDisposable
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private static readonly MasterListingRow C64 =
            new("Commodore 64", "250407 (long board)", "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", string.Empty);

        private static readonly MasterListingRow C128 =
            new("Commodore 128", "310378 (C128 & C128D)", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx", "Has 6581 SID.");

        private static readonly MasterListingRow Spectrum =
            new("ZX Spectrum 16K/48K", "Issue 4B", "ZX Spectrum/Spectrum 16K-48K/Issue 4B/Data ZX Issue 4B v2.0.0.xlsx", string.Empty);

        private const string Open128 = "Commodore/C128/310378 Open128";
        private const string Open128Workbook = "Commodore/C128/310378 Open128/Data C128 310378 Open128 v2.0.0.xlsx";

        private readonly string thisBeta = Path.Combine(Path.GetTempPath(), "crt-listing", Guid.NewGuid().ToString("N"));

        public BoardListingFlowTests()
        {
            Directory.CreateDirectory(this.thisBeta);
            DataTreeBuilder.ListingMaster(this.thisBeta, BoardListingFlowTests.C64, BoardListingFlowTests.C128, BoardListingFlowTests.Spectrum);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Path.GetDirectoryName(this.thisBeta)!, recursive: true);
            }
            catch (IOException)
            {
            }
        }

        private string MasterPath => Path.Combine(this.thisBeta, "Classic-Repair-Toolbox.v2.0.0.xlsx");

        private static BoardListingFlow Flow(FakeSubmissionStore store) =>
            new(store, new PublishLock(), NullLogger<BoardListingFlow>.Instance);

        private static ReviewAccess Admin() => ReviewAccess.For(BoardListingFlowTests.Account(administrator: true));

        private static ReviewAccess MaintainerOf(params string[] boards) =>
            ReviewAccess.For(BoardListingFlowTests.Account(administrator: false), boards);

        private static Handlers.Accounts.AccountRecord Account(bool administrator) =>
            new(
                Id: 7,
                Email: administrator ? "admin@example.com" : "maintainer@example.com",
                NormalisedEmail: administrator ? "admin@example.com" : "maintainer@example.com",
                PasswordHash: "hash",
                DisplayName: administrator ? "Admin" : "Maintainer",
                IsVerified: true,
                IsAdministrator: administrator,
                IsLocked: false,
                CreatedUtc: BoardListingFlowTests.Now,
                LastLoginUtc: null);

        // A submission for a board, left in `state` - with the contributor's notes from "Create
        // board", when given, sent `minutesLater` than Now.
        private static async Task<FakeSubmissionStore> SubmittedAsync(
            string boardId,
            string state = SubmissionState.Pending,
            FakeSubmissionStore? store = null,
            string? notes = null,
            int minutesLater = 0)
        {
            store ??= new FakeSubmissionStore();
            string[] parts = boardId.Split('/');
            DateTimeOffset sent = BoardListingFlowTests.Now.AddMinutes(minutesLater);

            long id = await store.CreateAsync(
                new NewSubmission(
                    boardId, parts[0], parts[1], parts[2],
                    null, "someone@example.com", "192.0.2.1", "hash", string.Empty,
                    "A new board.", 1, [], sent, sent.AddHours(24),
                    HardwareNotes: notes),
                CancellationToken.None);

            await store.SetStateAsync(id, state, BoardListingFlowTests.Now, CancellationToken.None);

            return store;
        }

        private static SetPlacementRequest Request(string? after = null) =>
            new(BoardListingFlowTests.Open128, "Commodore 128", "310378 Open128", "Open-source replica.", after ?? BoardListingFlowTests.C128.ExcelDataFile);

        // ------------------------------------------------------------------ the list

        [Fact]
        public async Task The_list_is_BETAs_in_CRTs_order_and_a_waiting_new_board_is_offered_to_place()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            BoardListingAnswer answer = outcome.Answer!;

            Assert.True(answer.HasList);
            Assert.Equal(["250407 (long board)", "310378 (C128 & C128D)", "Issue 4B"], answer.Rows.Select(row => row.BoardName));

            UnlistedBoardEntry unlisted = Assert.Single(answer.Unlisted);
            Assert.Equal(BoardListingFlowTests.Open128, unlisted.BoardId);
            Assert.False(unlisted.InBeta);
            Assert.True(unlisted.CanPlace);
            Assert.Null(unlisted.Placement);
        }

        // ###########################################################################################
        // Untouched, a new board of a hardware already listed joins it - under ITS name, not the
        // folder's: Open128 goes after the C128 boards as "Commodore 128", not as a hardware of its
        // own called "C128".
        // ###########################################################################################
        [Fact]
        public async Task A_new_board_is_suggested_at_the_end_of_its_hardware_under_that_hardwares_name()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            Assert.Equal(
                new BoardPlacement("Commodore 128", "310378 Open128", string.Empty, BoardListingFlowTests.C128.ExcelDataFile),
                Assert.Single(outcome.Answer!.Unlisted).Suggested);
        }

        // ------------------------------------------------------------------ the contributor's notes

        // ###########################################################################################
        // *** THE NOTES REACH THE MAIN EXCEL DATA FILE (owner request, 2026-10-05: "that note needs to
        // be sent also to the server, as this notes needs to go into the main Excel in the 'Hardware
        // and Board' sheet and in the column 'Hardware notes in "Overview" tab'"). *** The notes the
        // contributor typed in "Create board" are where the placement starts; saving that placement,
        // as the Maintainer tab sends it, writes them into BETA's list.
        // ###########################################################################################
        [Fact]
        public async Task The_contributors_notes_start_the_placement_and_saving_it_writes_them_into_the_list()
        {
            const string Notes = "Open-source replica of the C128 board.";

            DataTreeBuilder.Board(this.thisBeta, BoardListingFlowTests.Open128Workbook);
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(
                BoardListingFlowTests.Open128, SubmissionState.Merged, notes: Notes);

            BoardListingOutcome listed = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            BoardPlacement suggested = Assert.Single(listed.Answer!.Unlisted).Suggested;
            Assert.Equal(Notes, suggested.Notes);

            SetPlacementOutcome saved = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.MaintainerOf(BoardListingFlowTests.Open128),
                new SetPlacementRequest(BoardListingFlowTests.Open128, suggested.HardwareName, suggested.BoardName, suggested.Notes, suggested.AfterExcelDataFile),
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.True(saved.Answer!.ListedInBeta, saved.Error);
            Assert.Contains(
                new MasterListingRow("Commodore 128", "310378 Open128", BoardListingFlowTests.Open128Workbook, Notes),
                DataTreeBuilder.ListedIn(this.thisBeta));
        }

        // ###########################################################################################
        // Which notes count: the newest of a submission still waiting or published. A rejected,
        // withdrawn or abandoned one's were turned down with it; one still uploading nobody has seen.
        // ###########################################################################################
        [Theory]
        [InlineData(SubmissionState.Pending, "Newer.")]
        [InlineData(SubmissionState.Approved, "Newer.")]
        [InlineData(SubmissionState.ChangesRequested, "Newer.")]
        [InlineData(SubmissionState.Merged, "Newer.")]
        [InlineData(SubmissionState.Rejected, "Older.")]
        [InlineData(SubmissionState.Withdrawn, "Older.")]
        [InlineData(SubmissionState.Abandoned, "Older.")]
        [InlineData(SubmissionState.Uploading, "Older.")]
        public void The_placement_starts_with_the_newest_notes_of_a_submission_still_waiting_or_published(string newerState, string expected)
        {
            IReadOnlyList<SubmissionNotes> notes =
            [
                new(2, newerState, BoardListingFlowTests.Now.AddHours(1), "  Newer.  "),
                new(1, SubmissionState.Pending, BoardListingFlowTests.Now, "Older."),
            ];

            Assert.Equal(expected, BoardListingRules.SuggestedNotes(notes));
        }

        [Fact]
        public void With_no_notes_that_count_the_placement_starts_with_none()
        {
            Assert.Equal(string.Empty, BoardListingRules.SuggestedNotes(null));
            Assert.Equal(string.Empty, BoardListingRules.SuggestedNotes([]));
            Assert.Equal(string.Empty, BoardListingRules.SuggestedNotes([new(1, SubmissionState.Rejected, BoardListingFlowTests.Now, "Turned down.")]));
            Assert.Equal(string.Empty, BoardListingRules.SuggestedNotes([new(1, SubmissionState.Pending, BoardListingFlowTests.Now, "   ")]));
        }

        // Through the flow, a rejected submission's notes stay out while a later waiting one's are used.
        [Fact]
        public async Task A_rejected_submissions_notes_are_not_suggested_but_a_later_waiting_ones_are()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(
                BoardListingFlowTests.Open128, SubmissionState.Rejected, notes: "Turned down.");

            await BoardListingFlowTests.SubmittedAsync(
                BoardListingFlowTests.Open128, SubmissionState.Pending, store, notes: "Sent again.", minutesLater: 30);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            Assert.Equal("Sent again.", Assert.Single(outcome.Answer!.Unlisted).Suggested.Notes);
        }

        // A placement the maintainer SAVED is theirs - the contributor's notes only start one nobody
        // has saved, so clearing or rewording them sticks.
        [Fact]
        public async Task A_saved_placement_keeps_its_own_notes()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(
                BoardListingFlowTests.Open128, notes: "The contributor's words.");

            await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.MaintainerOf(BoardListingFlowTests.Open128),
                new SetPlacementRequest(BoardListingFlowTests.Open128, "Commodore 128", "310378 Open128", string.Empty, BoardListingFlowTests.C128.ExcelDataFile),
                this.thisBeta,
                BoardListingFlowTests.Now);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            UnlistedBoardEntry entry = Assert.Single(outcome.Answer!.Unlisted);
            Assert.Equal(string.Empty, entry.Placement!.Notes);
            Assert.Equal("The contributor's words.", entry.Suggested.Notes);
        }

        [Fact]
        public void A_hardware_nobody_has_listed_is_suggested_at_the_end_under_its_folder_name()
        {
            IReadOnlyList<MasterListingRow> rows = [BoardListingFlowTests.C64, BoardListingFlowTests.Spectrum];

            Assert.Equal(
                new BoardPlacement("Amiga 500", "Rev 6A", string.Empty, BoardListingFlowTests.Spectrum.ExcelDataFile),
                BoardListingRules.Suggest(rows, "Commodore", "Amiga 500", "Rev 6A"));

            Assert.Null(BoardListingRules.Suggest([], "Commodore", "Amiga 500", "Rev 6A").AfterExcelDataFile);
        }

        // Nothing to place until something of it could be published.
        [Theory]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.Abandoned)]
        [InlineData(SubmissionState.Withdrawn)]
        public async Task A_new_board_with_no_submission_waiting_is_not_offered(string state)
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128, state);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            Assert.Empty(outcome.Answer!.Unlisted);
        }

        // A board already in BETA that the list does not carry - one published before placing existed.
        [Fact]
        public async Task A_board_in_BETA_the_list_does_not_carry_is_offered_as_in_BETA()
        {
            DataTreeBuilder.Board(this.thisBeta, BoardListingFlowTests.Open128Workbook);

            BoardListingOutcome outcome = await BoardListingFlowTests.Flow(new FakeSubmissionStore()).ListAsync(
                BoardListingFlowTests.Admin(), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            UnlistedBoardEntry unlisted = Assert.Single(outcome.Answer!.Unlisted);
            Assert.Equal(BoardListingFlowTests.Open128, unlisted.BoardId);
            Assert.True(unlisted.InBeta);
        }

        // Everybody who may review sees the list; only who may publish a board may place it.
        [Fact]
        public async Task Placing_is_offered_only_to_a_maintainer_of_the_board_or_the_administrator()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            BoardListingOutcome other = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.MaintainerOf("Commodore/C64/250407"), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            BoardListingOutcome own = await BoardListingFlowTests.Flow(store).ListAsync(
                BoardListingFlowTests.MaintainerOf(BoardListingFlowTests.Open128), this.thisBeta, PublishedBoardLister.List(this.thisBeta));

            Assert.False(Assert.Single(other.Answer!.Unlisted).CanPlace);
            Assert.True(Assert.Single(own.Answer!.Unlisted).CanPlace);

            BoardListingOutcome nobody = await BoardListingFlowTests.Flow(store).ListAsync(
                ReviewAccess.For(BoardListingFlowTests.Account(administrator: false)), this.thisBeta, []);

            Assert.True(nobody.IsForbidden);
        }

        // ------------------------------------------------------------------ placing

        // Not in BETA yet: kept, and written when it is published - a row now would name a board
        // BETA does not have.
        [Fact]
        public async Task Placing_a_board_not_in_BETA_yet_saves_it_and_leaves_the_list_alone()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);
            byte[] before = File.ReadAllBytes(this.MasterPath);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(), BoardListingFlowTests.Request(), this.thisBeta, BoardListingFlowTests.Now);

            Assert.False(outcome.Answer!.ListedInBeta);
            Assert.Equal(
                new BoardPlacement("Commodore 128", "310378 Open128", "Open-source replica.", BoardListingFlowTests.C128.ExcelDataFile),
                store.Placements[BoardListingFlowTests.Open128]);
            Assert.Equal(before, File.ReadAllBytes(this.MasterPath));
        }

        // Already in BETA: written at once, where it was placed - the case of a board published
        // before placing existed.
        [Fact]
        public async Task Placing_a_board_already_in_BETA_adds_its_row_at_once_where_it_was_placed()
        {
            DataTreeBuilder.Board(this.thisBeta, BoardListingFlowTests.Open128Workbook);
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128, SubmissionState.Merged);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.MaintainerOf(BoardListingFlowTests.Open128),
                BoardListingFlowTests.Request(),
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.True(outcome.Answer!.ListedInBeta, outcome.Error);
            Assert.Equal(
                [
                    BoardListingFlowTests.C64,
                    BoardListingFlowTests.C128,
                    new MasterListingRow("Commodore 128", "310378 Open128", BoardListingFlowTests.Open128Workbook, "Open-source replica."),
                    BoardListingFlowTests.Spectrum,
                ],
                DataTreeBuilder.ListedIn(this.thisBeta));
        }

        // A board in BETA with no `boards` row yet - copied there by hand - gets one to keep the
        // placement on.
        [Fact]
        public async Task A_board_in_BETA_with_no_row_is_given_one_and_placed()
        {
            DataTreeBuilder.Board(this.thisBeta, BoardListingFlowTests.Open128Workbook);
            var store = new FakeSubmissionStore();

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(), BoardListingFlowTests.Request(), this.thisBeta, BoardListingFlowTests.Now);

            Assert.True(outcome.Answer!.ListedInBeta, outcome.Error);
            Assert.True(store.Boards.ContainsKey(BoardListingFlowTests.Open128));
        }

        [Fact]
        public async Task A_board_the_list_already_carries_cannot_be_placed()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync("Commodore/C128/310378");

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(),
                BoardListingFlowTests.Request() with { BoardId = "Commodore/C128/310378" },
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.True(outcome.IsConflict);
            Assert.Equal(BoardListingRules.AlreadyListedMessage, outcome.Error);
        }

        [Fact]
        public async Task Placing_after_a_row_the_list_does_not_carry_is_refused()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(),
                BoardListingFlowTests.Request(after: "Commodore/VIC-20/999999/Data.xlsx"),
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.Null(outcome.Answer);
            Assert.False(outcome.IsConflict);
            Assert.Empty(store.Placements);
        }

        // ###########################################################################################
        // CRT keys a board by "hardware name|board name", case-insensitively: a second board under
        // the same pair would be one board twice in everybody's drop-downs, sharing its settings.
        // Refused at the placement, where the maintainer can still pick another name.
        // ###########################################################################################
        [Fact]
        public async Task A_board_cannot_be_placed_under_names_another_board_is_listed_under()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(),
                BoardListingFlowTests.Request() with { HardwareName = "COMMODORE 128", BoardName = "310378 (c128 & c128d)" },
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.True(outcome.IsConflict);
            Assert.Contains("is already in the drop-down lists, for Commodore/C128/310378", outcome.Error);
            Assert.Empty(store.Placements);
        }

        [Fact]
        public async Task A_maintainer_of_another_board_cannot_place_it()
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.MaintainerOf("Commodore/C64/250407"),
                BoardListingFlowTests.Request(),
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.True(outcome.IsForbidden);
            Assert.Empty(store.Placements);
        }

        // What is saved must always be writable - the writer's own rule is applied to the request.
        [Theory]
        [InlineData("", "310378 Open128")]
        [InlineData("Commodore 128", "")]
        [InlineData("Commodore\t128", "310378 Open128")]
        public async Task Names_that_cannot_be_a_row_are_refused(string hardware, string board)
        {
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(),
                BoardListingFlowTests.Request() with { HardwareName = hardware, BoardName = board },
                this.thisBeta,
                BoardListingFlowTests.Now);

            Assert.Null(outcome.Answer);
            Assert.Empty(store.Placements);
        }

        [Fact]
        public async Task With_no_list_in_BETA_nothing_can_be_placed()
        {
            File.Delete(this.MasterPath);
            FakeSubmissionStore store = await BoardListingFlowTests.SubmittedAsync(BoardListingFlowTests.Open128);

            SetPlacementOutcome outcome = await BoardListingFlowTests.Flow(store).SetAsync(
                BoardListingFlowTests.Admin(), BoardListingFlowTests.Request(), this.thisBeta, BoardListingFlowTests.Now);

            Assert.True(outcome.IsConflict);
            Assert.Equal(BoardListingRules.NoMasterMessage, outcome.Error);
        }

        // ------------------------------------------------------------------ the request

        [Fact]
        public void A_request_is_trimmed_and_a_blank_after_means_first()
        {
            Assert.True(BoardListingRules.TryNormalise(
                new SetPlacementRequest("  Commodore/C128/310378 Open128 ", " Commodore 128 ", " 310378 Open128 ", " notes ", "  "),
                out string boardId,
                out BoardPlacement placement,
                out string problem), problem);

            Assert.Equal(BoardListingFlowTests.Open128, boardId);
            Assert.Equal(new BoardPlacement("Commodore 128", "310378 Open128", "notes", null), placement);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Commodore/C128")]
        [InlineData("Commodore/../C128/x")]
        public void A_request_not_naming_a_board_is_refused(string? boardId)
        {
            Assert.False(BoardListingRules.TryNormalise(
                new SetPlacementRequest(boardId, "Commodore 128", "B", string.Empty, null), out _, out _, out _));
        }
    }
}
