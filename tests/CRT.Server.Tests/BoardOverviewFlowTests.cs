using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers BoardOverviewFlow - the Maintainer tab's "Boards" screen (owner request, 2026-09-27):
    // every board, and one board's maintainers, contributors and recent submissions.
    //
    // THE RULE THAT MATTERS MOST is who may see it: ANY account that may review anything, for
    // EVERY board (owner decision, "Everything for everyone") - but its email addresses only for
    // its own maintainers and the administrator (owner request, 2026-10-05). A maintainer of one
    // board reads another board's facts by name; an account in no pool reads nothing.
    //
    // No database, no filesystem: the trees are handed in as lists, the way the endpoint hands them
    // in after PublishedBoardLister has read them.
    // ###########################################################################################
    public sealed class BoardOverviewFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";
        private const string Vic = "Commodore/VIC-20/250403";

        private static readonly IReadOnlyList<PublishedBoardLister.KnownBoard> Beta =
        [
            new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407"),
            new(BoardOverviewFlowTests.C128, "Commodore", "C128", "310378")
        ];

        private static readonly IReadOnlyList<PublishedBoardLister.KnownBoard> Production =
        [
            new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407")
        ];

        private static SubmissionRecord Record(
            long id,
            string? email = "dennis@example.com",
            long? account = null,
            string state = SubmissionState.Pending,
            string board = BoardOverviewFlowTests.C64,
            DateTimeOffset? created = null,
            DateTimeOffset? decided = null,
            string? comment = null) =>
            new(id, board, account, email, "hash", "", state, $"Submission {id}", 1,
                created ?? BoardOverviewFlowTests.Now.AddDays(-id), null, decided, comment);

        private static BoardSubmissionRecord Sent(SubmissionRecord record, bool byMaintainer = false) => new(record, byMaintainer);

        private static async Task<(FakeAccountStore Accounts, ReviewAccess Anna, long AnnaId, long BoId)> MaintainersAsync()
        {
            var accounts = new FakeAccountStore();

            long anna = await accounts.CreateAccountAsync(new NewAccount(
                "anna@example.com", "anna@example.com", "hash", "Anna", BoardOverviewFlowTests.Now));
            accounts.Accounts[anna] = accounts.Accounts[anna] with { IsVerified = true };

            long bo = await accounts.CreateAccountAsync(new NewAccount(
                "bo@example.com", "bo@example.com", "hash", "Bo", BoardOverviewFlowTests.Now));
            accounts.Accounts[bo] = accounts.Accounts[bo] with { IsVerified = true };

            // Anna maintains the C128 ONLY - the screen still shows her the C64.
            await accounts.AddMaintainerAsync(BoardOverviewFlowTests.C128, anna, 1, BoardOverviewFlowTests.Now);
            await accounts.AddMaintainerAsync(BoardOverviewFlowTests.C64, bo, 1, BoardOverviewFlowTests.Now);

            ReviewAccess access = ReviewAccess.For(
                accounts.Accounts[anna],
                await accounts.GetReviewedBoardIdsAsync(anna));

            return (accounts, access, anna, bo);
        }

        // -----------------------------------------------------------------------------------
        // Who may see it
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** EVERY BOARD FOR EVERYONE (owner decision, 2026-09-27) - BUT ITS ADDRESSES FOR ITS OWN
        // MAINTAINERS ONLY (owner request, 2026-10-05: "I do not think that normal maintainer should be
        // able to see other email addresses if they are not set as maintainer for that system"). ***
        // Anna maintains only the C128: she is shown the C64's maintainers, contributors, submissions
        // and history - by name, with not one address anywhere in the answer. Bo maintains the C64,
        // and the administrator every board: both get every address.
        // ###########################################################################################
        private static async Task<(FakeAccountStore Accounts, FakeSubmissionStore Store, ReviewAccess Anna, ReviewAccess Bo, ReviewAccess Admin)> PeopleOnTheC64Async()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, long bo) = await BoardOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            long admin = await accounts.CreateAccountAsync(new NewAccount(
                "admin@example.com", "admin@example.com", "hash", "Dennis", BoardOverviewFlowTests.Now));
            accounts.Accounts[admin] = accounts.Accounts[admin] with { IsVerified = true, IsAdministrator = true };

            long dora = await accounts.CreateAccountAsync(new NewAccount(
                "dora@example.com", "dora@example.com", "hash", "Dora", BoardOverviewFlowTests.Now));
            accounts.Accounts[dora] = accounts.Accounts[dora] with { IsVerified = true };

            await store.EnsureBoardAsync(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", BoardOverviewFlowTests.Now);

            // Carl sent one without an account; Dora one signed in.
            store.Submissions[1] = BoardOverviewFlowTests.Record(1, "carl@example.com", state: SubmissionState.Merged);
            store.Submissions[2] = BoardOverviewFlowTests.Record(2, email: null, account: dora);

            // The history: Bo made a maintainer, somebody invited, and Carl discarding his draft.
            await accounts.WriteAuditAsync(new AuditEntry(
                admin, "admin@example.com", BoardHistoryEvents.MaintainerAdded, BoardOverviewFlowTests.C64,
                $"account {bo} (bo@example.com)", BoardOverviewFlowTests.Now.AddDays(-3)));
            await accounts.WriteAuditAsync(new AuditEntry(
                admin, "admin@example.com", BoardHistoryEvents.Invited, BoardOverviewFlowTests.C64,
                "eve@example.com", BoardOverviewFlowTests.Now.AddDays(-2)));
            await accounts.WriteAuditAsync(new AuditEntry(
                null, "carl@example.com", BoardHistoryEvents.DraftDiscarded, "#1", null, BoardOverviewFlowTests.Now.AddDays(-1)));

            return (
                accounts,
                store,
                anna,
                ReviewAccess.For(accounts.Accounts[bo], await accounts.GetReviewedBoardIdsAsync(bo)),
                ReviewAccess.For(accounts.Accounts[admin]));
        }

        private static async Task<BoardDetailAnswer> C64DetailAsync(ReviewAccess access, FakeAccountStore accounts, FakeSubmissionStore store) =>
            (await BoardOverviewFlow.DetailAsync(
                access, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, BoardOverviewFlowTests.Production, store, accounts)).Detail!;

        [Fact]
        public async Task A_maintainer_sees_a_board_they_do_not_maintain_but_not_one_of_its_email_addresses()
        {
            (FakeAccountStore accounts, FakeSubmissionStore store, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.PeopleOnTheC64Async();

            BoardDetailAnswer detail = await BoardOverviewFlowTests.C64DetailAsync(anna, accounts, store);

            Assert.True(detail.AddressesHidden);

            // Everybody is still there - by name.
            PoolMaintainerEntry maintainer = Assert.Single(detail.Maintainers);
            Assert.Equal(("Bo", string.Empty), (maintainer.DisplayName, maintainer.Email));
            Assert.Equal(2, detail.Contributors.Count);
            Assert.Contains(detail.Contributors, contributor => contributor.Name == "Dora");
            Assert.All(detail.Contributors, contributor => Assert.Null(contributor.Email));
            Assert.All(detail.Submissions, submission => Assert.Null(submission.ContactEmail));

            // The history names the people by their accounts: who made Bo a maintainer, and whom.
            BoardHistoryEntry granted = detail.History!.Single(entry => entry.Event == BoardHistoryEvents.MaintainerAdded);
            Assert.Equal(("Dennis", "Bo"), (granted.Who, granted.Detail));
            Assert.Null(detail.History!.Single(entry => entry.Event == BoardHistoryEvents.Invited).Detail);
            Assert.Equal("Dora", detail.History!.Single(entry => entry.Event == BoardHistoryEvents.Sent && entry.SubmissionId == 2).Who);

            // *** And not one address anywhere in what goes over the wire. *** The catch-all: a new
            // field carrying one would fail here, whatever it is called.
            Assert.DoesNotContain("@", System.Text.Json.JsonSerializer.Serialize(detail), StringComparison.Ordinal);
        }

        [Fact]
        public async Task The_boards_own_maintainer_and_the_administrator_see_every_address()
        {
            (FakeAccountStore accounts, FakeSubmissionStore store, _, ReviewAccess bo, ReviewAccess admin) = await BoardOverviewFlowTests.PeopleOnTheC64Async();

            foreach (ReviewAccess access in new[] { bo, admin })
            {
                BoardDetailAnswer detail = await BoardOverviewFlowTests.C64DetailAsync(access, accounts, store);

                Assert.False(detail.AddressesHidden);
                Assert.Equal("bo@example.com", Assert.Single(detail.Maintainers).Email);
                Assert.Contains(detail.Contributors, contributor => contributor.Email == "carl@example.com");
                Assert.Contains(detail.Contributors, contributor => contributor.Email == "dora@example.com");
                Assert.Contains(detail.Submissions, submission => submission.ContactEmail == "carl@example.com");
                // A maintainer made is named by the name on the account (owner request, 2026-10-09);
                // an invitation's person has no account yet, so the address it went to.
                Assert.Equal("Bo", detail.History!.Single(entry => entry.Event == BoardHistoryEvents.MaintainerAdded).Detail);
                Assert.Equal("eve@example.com", detail.History!.Single(entry => entry.Event == BoardHistoryEvents.Invited).Detail);
            }
        }

        // The detail names contributors, deciders, audit actors and the people a pool change named -
        // asked one by one, that was a database round trip per person on a request made under the
        // overlay and again every minute (code review, 2026-10-09). One query for them all, and the
        // names still arrive.
        [Fact]
        public async Task A_boards_detail_looks_up_every_account_it_names_in_one_query()
        {
            (FakeAccountStore accounts, FakeSubmissionStore store, _, ReviewAccess bo, _) = await BoardOverviewFlowTests.PeopleOnTheC64Async();

            BoardDetailAnswer detail = await BoardOverviewFlowTests.C64DetailAsync(bo, accounts, store);

            Assert.Equal((0, 1), (accounts.FindByIdCount, accounts.FindByIdsCount));
            Assert.Contains(detail.Contributors, contributor => contributor.Name == "Dora");

            BoardHistoryEntry granted = detail.History!.Single(entry => entry.Event == BoardHistoryEvents.MaintainerAdded);
            Assert.Equal(("Dennis", "Bo"), (granted.Who, granted.Detail));
        }

        // An administrator named a board's maintainer (2026-10-05) is shown in its pool to everybody
        // - the point of naming them - by name to a maintainer of another board.
        [Fact]
        public async Task An_administrator_in_a_pool_is_shown_as_its_maintainer()
        {
            (FakeAccountStore accounts, FakeSubmissionStore store, ReviewAccess anna, _, ReviewAccess admin) = await BoardOverviewFlowTests.PeopleOnTheC64Async();

            await accounts.AddMaintainerAsync(BoardOverviewFlowTests.C64, admin.Account.Id, admin.Account.Id, BoardOverviewFlowTests.Now);

            BoardDetailAnswer detail = await BoardOverviewFlowTests.C64DetailAsync(anna, accounts, store);

            Assert.Equal(["Bo", "Dennis"], detail.Maintainers.Select(maintainer => maintainer.DisplayName).Order(StringComparer.Ordinal));
        }

        // An account in no pool may not review anything, and is refused both - the queue's own rule.
        [Fact]
        public async Task An_account_that_reviews_nothing_is_refused_the_list_and_the_detail()
        {
            var accounts = new FakeAccountStore();
            long id = await accounts.CreateAccountAsync(new NewAccount(
                "nobody@example.com", "nobody@example.com", "hash", "Nobody", BoardOverviewFlowTests.Now));
            accounts.Accounts[id] = accounts.Accounts[id] with { IsVerified = true };

            ReviewAccess nobody = ReviewAccess.For(accounts.Accounts[id]);
            var store = new FakeSubmissionStore();

            Assert.True((await BoardOverviewFlow.ListAsync(nobody, BoardOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
            Assert.True((await BoardOverviewFlow.DetailAsync(nobody, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
            Assert.True((await BoardOverviewFlow.ListAsync(null, BoardOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
        }

        // A board that is neither a row nor a board in the tree is not a board.
        [Fact]
        public async Task A_board_that_exists_nowhere_is_not_found()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();

            BoardOverviewOutcome outcome = await BoardOverviewFlow.DetailAsync(
                anna, "Acme/Nothing/1", BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts);

            Assert.True(outcome.IsNotFound);
        }

        // -----------------------------------------------------------------------------------
        // The list
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The rows AND the tree, once each: a shipped board nobody has touched has no row and is
        // still a board, and a new board waiting for its first review has a row and no board.
        // ###########################################################################################
        [Fact]
        public void The_list_is_every_row_and_every_board_in_the_tree_once_each()
        {
            BoardRecord[] rows =
            [
                new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", "2026-September-20", true, "beta-hash", "2026-May-14", "prod-hash", BoardOverviewFlowTests.Now.AddDays(-9)),
                new(BoardOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403", null, true)
            ];

            MaintainerRecord[] pool =
            [
                new(BoardOverviewFlowTests.C64, 1, "Bo", "bo@example.com"),
                new(BoardOverviewFlowTests.C64, 2, "Cy", "cy@example.com")
            ];

            IReadOnlyList<BoardOverviewEntry> list = BoardOverviewFlow.Entries(
                rows, BoardOverviewFlowTests.Beta, BoardOverviewFlowTests.Production, pool);

            Assert.Equal([BoardOverviewFlowTests.C128, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Vic], list.Select(entry => entry.BoardId));

            BoardOverviewEntry c64 = list.Single(entry => entry.BoardId == BoardOverviewFlowTests.C64);
            Assert.True(c64.InBeta);
            Assert.True(c64.InProduction);
            Assert.True(c64.IsAwaitingProduction);
            Assert.Equal("2026-September-20", c64.BetaRevision);
            Assert.Equal("2026-May-14", c64.ProductionRevision);
            Assert.Equal(2, c64.MaintainerCount);

            // In the BETA tree, no row: a shipped board. Not in production's list.
            BoardOverviewEntry c128 = list.Single(entry => entry.BoardId == BoardOverviewFlowTests.C128);
            Assert.True(c128.InBeta);
            Assert.False(c128.InProduction);
            Assert.False(c128.IsAwaitingProduction);
            Assert.Equal("C128", c128.Hardware);
            Assert.Equal(0, c128.MaintainerCount);

            // A row, no board anywhere: a new board nothing has published yet.
            BoardOverviewEntry vic = list.Single(entry => entry.BoardId == BoardOverviewFlowTests.Vic);
            Assert.False(vic.InBeta);
            Assert.False(vic.InProduction);
        }

        // ###########################################################################################
        // BETA's content hash rides on each board (2026-10-04) - the row's record of what BETA holds,
        // which every publish to BETA and every push-back moves. The Maintainer tab's Boards screen
        // reads its open table again when it moves; a board no row records has none to say.
        // ###########################################################################################
        [Fact]
        public void Each_board_carries_BETA_s_content_hash_from_its_row()
        {
            BoardRecord row = new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", "2026-September-20", true, "beta-hash", "2026-May-14", "prod-hash", null);

            Assert.Equal("beta-hash", BoardOverviewFlow.Entry(BoardOverviewFlowTests.C64, row, BoardOverviewFlowTests.Beta[0], null, 0).BetaContentHash);
            Assert.Null(BoardOverviewFlow.Entry(BoardOverviewFlowTests.C128, null, BoardOverviewFlowTests.Beta[1], null, 0).BetaContentHash);
        }

        // ###########################################################################################
        // *** WHAT THE STABLE SOURCE ALONE HOLDS OR LISTS IS ON THE LIST (owner request, 2026-10-04:
        // "I do not expect there should be cases where something can only be listed in stable? If so,
        // it must be flagged in the left-sided menu 'Boards' list"). *** Such a board used to be
        // missing from the screen altogether, so nothing could flag it. Each entry also says whether
        // each source's drop-down list names it.
        // ###########################################################################################
        [Fact]
        public void The_list_carries_what_each_drop_down_list_names_and_what_only_stable_has()
        {
            const string Pet = "Commodore/PET/2001";

            IReadOnlyList<PublishedBoardLister.KnownBoard> stableTree =
            [
                new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407"),
                new(BoardOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403")
            ];

            var listings = new BoardListings(
                new HashSet<string>([BoardOverviewFlowTests.C64], StringComparer.OrdinalIgnoreCase),
                new HashSet<string>([BoardOverviewFlowTests.C64, BoardOverviewFlowTests.C128, BoardOverviewFlowTests.Vic, Pet], StringComparer.OrdinalIgnoreCase));

            IReadOnlyList<BoardOverviewEntry> list = BoardOverviewFlow.Entries([], BoardOverviewFlowTests.Beta, stableTree, [], listings);

            Assert.Equal([BoardOverviewFlowTests.C128, BoardOverviewFlowTests.C64, Pet, BoardOverviewFlowTests.Vic], list.Select(entry => entry.BoardId));

            BoardOverviewEntry c128 = list.Single(entry => entry.BoardId == BoardOverviewFlowTests.C128);
            Assert.Equal((false, true), (c128.ListedInBeta, c128.ListedInStable));

            // In the stable source's tree only: on the list, named from that tree, not in BETA.
            BoardOverviewEntry vic = list.Single(entry => entry.BoardId == BoardOverviewFlowTests.Vic);
            Assert.False(vic.InBeta);
            Assert.True(vic.InProduction);
            Assert.Equal("VIC-20", vic.Hardware);

            // In the stable source's drop-down list only, no board anywhere: named from its id.
            BoardOverviewEntry pet = list.Single(entry => entry.BoardId == Pet);
            Assert.Equal(("Commodore", "PET", "2001"), (pet.Manufacturer, pet.Hardware, pet.Board));
            Assert.Equal((false, true), (pet.ListedInBeta, pet.ListedInStable));
        }

        // A list that could not be read says nothing - never "not listed" without looking.
        [Fact]
        public void A_list_that_could_not_be_read_says_nothing_about_listing()
        {
            IReadOnlyList<BoardOverviewEntry> list = BoardOverviewFlow.Entries(
                [], BoardOverviewFlowTests.Beta, null, [], new BoardListings(null, null));

            Assert.All(list, entry => Assert.Equal((null, null), (entry.ListedInBeta, entry.ListedInStable)));
        }

        // ###########################################################################################
        // The listings are read from each source's newest main Excel data file - the real files, by
        // board id in any case; a source with no such file is null. Read once per version of the
        // file: a rewrite is read afresh.
        // ###########################################################################################
        [Fact]
        public void The_listings_are_read_from_each_sources_main_Excel_data_file()
        {
            string root = Path.Combine(Path.GetTempPath(), "crt-listings", Guid.NewGuid().ToString("N"));
            string beta = Path.Combine(root, "beta");
            string stable = Path.Combine(root, "stable");

            Directory.CreateDirectory(beta);
            Directory.CreateDirectory(stable);

            try
            {
                DataTreeBuilder.Master(beta, "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx");
                DataTreeBuilder.Master(stable, "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx");

                BoardListings listings = BoardOverviewFlow.ReadListings(beta, stable);

                Assert.Equal([BoardOverviewFlowTests.C64], listings.Beta!);
                Assert.True(listings.Stable!.SetEquals([BoardOverviewFlowTests.C64, "commodore/c128/310378"]));
                Assert.Null(BoardOverviewFlow.ReadListings(beta, Path.Combine(root, "none")).Stable);

                // Rewritten with another board: read again.
                File.SetLastWriteTimeUtc(Path.Combine(beta, "Classic-Repair-Toolbox.v2.0.0.xlsx"), DateTime.UtcNow.AddDays(-1));
                DataTreeBuilder.Master(beta, "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx");

                Assert.Equal(2, BoardOverviewFlow.ReadListings(beta, null).Beta!.Count);
            }
            finally
            {
                Directory.Delete(root, recursive: true);
            }
        }

        // A board the stable source alone holds can be opened - it is on the list, and a list entry
        // that answers "no such board" would be a dead end.
        [Fact]
        public async Task A_board_only_the_stable_source_holds_has_a_detail()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();

            IReadOnlyList<PublishedBoardLister.KnownBoard> stableTree = [new(BoardOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403")];

            BoardOverviewOutcome outcome = await BoardOverviewFlow.DetailAsync(
                anna, BoardOverviewFlowTests.Vic, BoardOverviewFlowTests.Beta, stableTree, new FakeSubmissionStore(), accounts,
                listings: new BoardListings(new HashSet<string>(), new HashSet<string>([BoardOverviewFlowTests.Vic])));

            Assert.False(outcome.IsNotFound);
            Assert.Equal(("VIC-20", false, true), (outcome.Detail!.Board.Hardware, outcome.Detail.Board.InBeta, outcome.Detail.Board.ListedInStable));
        }

        // With no production tree to ask, the list never claims a board is not in production.
        [Fact]
        public void Without_a_production_tree_nothing_is_said_about_production()
        {
            IReadOnlyList<BoardOverviewEntry> list = BoardOverviewFlow.Entries([], BoardOverviewFlowTests.Beta, inProduction: null, []);

            Assert.All(list, entry => Assert.Null(entry.InProduction));
        }

        // A closed board says so; a board with no row has nothing closing it.
        [Fact]
        public void A_closed_board_reads_as_not_accepting()
        {
            BoardRecord closed = new(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", null, IsAccepting: false);

            Assert.False(BoardOverviewFlow.Entry(BoardOverviewFlowTests.C64, closed, null, null, 0).IsAccepting);
            Assert.True(BoardOverviewFlow.Entry(BoardOverviewFlowTests.C128, null, BoardOverviewFlowTests.Beta[1], null, 0).IsAccepting);
        }

        // -----------------------------------------------------------------------------------
        // Contributors
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** ONE CONTRIBUTOR IS ONE PERSON, COUNTED THE WAY ContributorHistory COUNTS. *** The same
        // email in any case and with stray spaces is one person; Accepted is merged; a rejection
        // counts only when a maintainer made it; a replaced submission is not counted at all.
        // ###########################################################################################
        [Fact]
        public void A_contributors_submissions_are_counted_by_how_they_ended()
        {
            BoardSubmissionRecord[] sent =
            [
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(1, "Dennis@Example.com", state: SubmissionState.Merged), true),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(2, " dennis@example.com ", state: SubmissionState.Pending)),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(3, "dennis@example.com", state: SubmissionState.Approved), true),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(4, "dennis@example.com", state: SubmissionState.ChangesRequested), true),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(5, "dennis@example.com", state: SubmissionState.Rejected), true),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(6, "dennis@example.com", state: SubmissionState.Rejected)),   // the automatic checks
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(7, "dennis@example.com", state: SubmissionState.Withdrawn))   // replaced
            ];

            BoardContributorEntry dennis = Assert.Single(BoardOverviewFlow.Contributors(sent));

            Assert.Equal(new BoardContributorEntry("Dennis@Example.com", null, Accepted: 1, Waiting: 2, ChangesRequested: 1, Rejected: 1,
                LastSubmittedUtc: BoardOverviewFlowTests.Now.AddDays(-1)), dennis);
        }

        // ###########################################################################################
        // A signed-in contributor is their ACCOUNT - named by it, whatever the submissions carry -
        // and is never merged with an anonymous submission from the same address, the rule
        // SubmissionReplacementRules states. On purpose (confirmed in the code review, 2026-09-27):
        // an anonymous sender's address is unverified, so merging would let anybody add submissions
        // to an account holder's record. Newest contributor first.
        // ###########################################################################################
        [Fact]
        public void A_signed_in_contributor_is_their_account_and_the_newest_comes_first()
        {
            BoardSubmissionRecord[] sent =
            [
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(1, null, account: 7, created: BoardOverviewFlowTests.Now.AddDays(-1))),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(2, "anna@example.com", created: BoardOverviewFlowTests.Now.AddDays(-5))),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(3, null, account: 7, created: BoardOverviewFlowTests.Now.AddDays(-9)))
            ];

            var accounts = new Dictionary<long, AccountRecord>
            {
                [7] = new(7, "anna@example.com", "anna@example.com", "hash", "Anna", true, false, false, BoardOverviewFlowTests.Now, null)
            };

            IReadOnlyList<BoardContributorEntry> contributors = BoardOverviewFlow.Contributors(sent, accounts);

            Assert.Equal(2, contributors.Count);
            Assert.Equal(("anna@example.com", "Anna", 2), (contributors[0].Email, contributors[0].Name, contributors[0].Waiting));
            Assert.Equal(("anna@example.com", (string?)null, 1), (contributors[1].Email, contributors[1].Name, contributors[1].Waiting));
        }

        // A submission naming nobody cannot be anybody's, and is left out rather than listed blank.
        [Fact]
        public void A_submission_naming_nobody_is_no_contributor()
        {
            Assert.Empty(BoardOverviewFlow.Contributors([BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(1, "  "))]));
        }

        // -----------------------------------------------------------------------------------
        // Submissions
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE WORD THE CONTRIBUTOR IS TOLD. *** A merged submission decided after the board's
        // last promotion is in BETA ("merged"), one decided before it has reached everyone
        // ("published"), and a pending one a rollback returned was pushed back ("returned") - the
        // same words CRT shows, so maintainer and contributor describe it alike. Newest first.
        //
        // "returned" comes from the rollback's RECORD (migration 0015, code review 2026-09-29), no
        // longer from "pending with a comment": submission 5 carries a comment and no record, and is
        // simply pending.
        // ###########################################################################################
        [Fact]
        public void Each_submission_carries_the_word_its_contributor_is_told_newest_first()
        {
            DateTimeOffset promoted = BoardOverviewFlowTests.Now.AddDays(-3);

            BoardSubmissionRecord[] sent =
            [
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(1, state: SubmissionState.Merged, decided: promoted.AddDays(-1))),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(2, state: SubmissionState.Merged, decided: promoted.AddDays(1))),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(3, state: SubmissionState.Pending, comment: "U8 is wrong.")),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(4, state: SubmissionState.Pending)),
                BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(5, state: SubmissionState.Pending, comment: "A note, no rollback."))
            ];

            var returns = new Dictionary<long, DateTimeOffset> { [3] = BoardOverviewFlowTests.Now };

            IReadOnlyList<BoardSubmissionEntry> listed = BoardOverviewFlow.Submissions(sent, promoted, returns: returns);

            Assert.Equal([5L, 4L, 3L, 2L, 1L], listed.Select(entry => entry.Id));
            Assert.Equal(["pending", "pending", "returned", "merged", "published"], listed.Select(entry => entry.State));
            Assert.Equal("U8 is wrong.", listed[2].DecisionComment);
        }

        // Only the newest ListedSubmissions are listed; the contributors are counted over them all.
        [Fact]
        public void Only_the_newest_submissions_are_listed()
        {
            BoardSubmissionRecord[] sent = Enumerable.Range(1, BoardOverviewFlow.ListedSubmissions + 5)
                .Select(id => BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(id)))
                .ToArray();

            IReadOnlyList<BoardSubmissionEntry> listed = BoardOverviewFlow.Submissions(sent, null);

            Assert.Equal(BoardOverviewFlow.ListedSubmissions, listed.Count);
            Assert.Equal(BoardOverviewFlow.ListedSubmissions + 5, listed[0].Id);
            Assert.Equal(BoardOverviewFlow.ListedSubmissions + 5, Assert.Single(BoardOverviewFlow.Contributors(sent)).Waiting);
        }

        // A signed-in contributor's submission stores no address - their account's is shown.
        [Fact]
        public void A_signed_in_contributors_submission_shows_their_accounts_address()
        {
            var accounts = new Dictionary<long, AccountRecord>
            {
                [7] = new(7, "anna@example.com", "anna@example.com", "hash", "Anna", true, false, false, BoardOverviewFlowTests.Now, null)
            };

            BoardSubmissionEntry entry = Assert.Single(BoardOverviewFlow.Submissions(
                [BoardOverviewFlowTests.Sent(BoardOverviewFlowTests.Record(1, null, account: 7))], null, accounts));

            Assert.Equal("anna@example.com", entry.ContactEmail);
        }

        // -----------------------------------------------------------------------------------
        // Through the store
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The detail end to end over the fakes: only THIS board's submissions, never one still
        // uploading or abandoned (never anybody's to review), a maintainer's rejection counted and
        // the automatic checks' not, and the account's name looked up.
        // ###########################################################################################
        [Fact]
        public async Task The_detail_reads_only_this_boards_arrived_submissions()
        {
            (FakeAccountStore accounts, ReviewAccess anna, long annaId, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            await store.EnsureBoardAsync(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", BoardOverviewFlowTests.Now);

            store.Submissions[1] = BoardOverviewFlowTests.Record(1, "carl@example.com", state: SubmissionState.Rejected);
            store.Decisions[1] = new RecordedDecision(SubmissionState.Rejected, annaId, "No.", BoardOverviewFlowTests.Now);
            store.Submissions[2] = BoardOverviewFlowTests.Record(2, "carl@example.com", state: SubmissionState.Rejected);
            store.Submissions[3] = BoardOverviewFlowTests.Record(3, "carl@example.com", state: SubmissionState.Uploading);
            store.Submissions[4] = BoardOverviewFlowTests.Record(4, "carl@example.com", state: SubmissionState.Abandoned);
            store.Submissions[5] = BoardOverviewFlowTests.Record(5, "carl@example.com", board: BoardOverviewFlowTests.C128);
            store.Submissions[6] = BoardOverviewFlowTests.Record(6, null, account: annaId, state: SubmissionState.Merged);

            // Read as one of the C64's own maintainers, who sees its addresses (2026-10-05) - this
            // test is about which submissions count, and finds the contributors by address.
            ReviewAccess maintainerOfC64 = ReviewAccess.For(anna.Account, [BoardOverviewFlowTests.C64]);

            BoardDetailAnswer detail = (await BoardOverviewFlow.DetailAsync(
                maintainerOfC64, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, BoardOverviewFlowTests.Production, store, accounts)).Detail!;

            Assert.Equal([6L, 2L, 1L], detail.Submissions.Select(entry => entry.Id));

            BoardContributorEntry carl = detail.Contributors.Single(contributor => contributor.Email == "carl@example.com");
            Assert.Equal(1, carl.Rejected);

            BoardContributorEntry annaAsContributor = detail.Contributors.Single(contributor => contributor.Name == "Anna");
            Assert.Equal(("anna@example.com", 1), (annaAsContributor.Email, annaAsContributor.Accepted));

            Assert.Equal(BoardOverviewFlowTests.C64, detail.Board.BoardId);
            Assert.True(detail.Board.InProduction);
        }

        // ###########################################################################################
        // What each submission changed as it went into BETA (owner request, 2026-10-04) rides on its
        // entry, for the History view - and one never published, or published before the record was
        // kept, carries none.
        // ###########################################################################################
        [Fact]
        public async Task The_detail_carries_what_each_published_submission_changed()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            await store.EnsureBoardAsync(BoardOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", BoardOverviewFlowTests.Now);

            store.Submissions[1] = BoardOverviewFlowTests.Record(1, state: SubmissionState.Merged);
            store.Submissions[2] = BoardOverviewFlowTests.Record(2, state: SubmissionState.Merged);
            store.Submissions[3] = BoardOverviewFlowTests.Record(3);

            var changes = new SubmissionChanges(
                false,
                [new SectionChanges("Components", 0, 1, 0, 0, [], [new ChangedRowFact("U8", ["Part-number"])], [], [])],
                FileChanges.None);

            await store.SetChangesAsync(1, changes, BoardOverviewFlowTests.Now, CancellationToken.None);

            BoardDetailAnswer detail = (await BoardOverviewFlow.DetailAsync(
                anna, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, null, store, accounts)).Detail!;

            Assert.Same(changes, detail.Submissions.Single(entry => entry.Id == 1).Changes);
            Assert.Null(detail.Submissions.Single(entry => entry.Id == 2).Changes);
            Assert.Null(detail.Submissions.Single(entry => entry.Id == 3).Changes);
        }

        // -----------------------------------------------------------------------------------
        // Board views (owner request, 2026-09-27)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The list's count is the last 30 whole days, BETA-source views left out - and a board nobody
        // looked at reads 0, not nothing: it was counted.
        // ###########################################################################################
        [Fact]
        public async Task The_list_carries_each_boards_views_in_the_last_30_days_beta_left_out()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore();

            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-1)));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-29)));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-31)));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-1), fromBeta: true));

            BoardOverviewOutcome outcome = await BoardOverviewFlow.ListAsync(
                anna, BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts,
                boardViews: views, now: BoardOverviewFlowTests.Now);

            Assert.Equal(
                [(BoardOverviewFlowTests.C128, (int?)0), (BoardOverviewFlowTests.C64, 2)],
                outcome.Boards!.Select(board => (board.BoardId, board.ViewsLast30Days)));
        }

        // The detail: the numbers, and the same 30-day count on its entry as the list shows.
        [Fact]
        public async Task The_detail_carries_the_boards_view_statistics()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore();

            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-1), "DK"));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-1), "DK"));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-20), "US"));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Now.AddDays(-2), "DE", fromBeta: true));
            views.Rows.Add(FakeBoardViewStore.Row(BoardOverviewFlowTests.C128, BoardOverviewFlowTests.Now.AddDays(-1), "SE"));

            BoardDetailAnswer detail = (await BoardOverviewFlow.DetailAsync(
                anna, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts,
                now: BoardOverviewFlowTests.Now, boardViews: views)).Detail!;

            BoardViewStatistics stats = detail.Views!;
            Assert.Equal((2, 3, 3, 1), (stats.Last7Days, stats.Last30Days, stats.Last365Days, stats.FromBetaLast30Days));
            Assert.Equal(["Denmark 2", "United States 1"], stats.TopCountries.Select(country => $"{country.CountryName} {country.Views}"));
            Assert.Equal(3, detail.Board.ViewsLast30Days);
        }

        // ###########################################################################################
        // *** THE VIEW COUNTS NEVER FAIL THE SCREEN. *** A database error reading them leaves them out
        // - the list and the detail still answer, with no numbers - rather than a 500 for everything.
        // ###########################################################################################
        [Fact]
        public async Task Views_that_cannot_be_read_are_left_out_and_the_screen_still_answers()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore { FailReads = true };

            BoardOverviewOutcome list = await BoardOverviewFlow.ListAsync(
                anna, BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts, boardViews: views);

            BoardOverviewOutcome detail = await BoardOverviewFlow.DetailAsync(
                anna, BoardOverviewFlowTests.C64, BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts, boardViews: views);

            Assert.All(list.Boards!, board => Assert.Null(board.ViewsLast30Days));
            Assert.Null(detail.Detail!.Views);
            Assert.Null(detail.Detail.Board.ViewsLast30Days);
        }

        // Without a store (a caller that does not count views) there are no numbers either.
        [Fact]
        public async Task Without_a_view_store_there_are_no_view_numbers()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();

            BoardOverviewOutcome list = await BoardOverviewFlow.ListAsync(
                anna, BoardOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts);

            Assert.All(list.Boards!, board => Assert.Null(board.ViewsLast30Days));
        }
    
        // ###########################################################################################
        // *** HOW A BOARD'S SUBMISSIONS WENT, ON ITS LINE IN THE LIST (owner request, 2026-10-09:
        // "19 submissions in total; 2 rejected, 1 in BETA, 17 in stable"). *** Each in the word its
        // contributor is told: merged and decided before the last publish to stable is in stable,
        // after it in BETA. Waiting covers review and changes requested; a rejection counts only when
        // a maintainer made it, as on the Contributor view - and Total is the four added up.
        // ###########################################################################################
        [Fact]
        public void A_boards_submissions_are_counted_by_where_they_stand_now()
        {
            BoardSubmissionCounts counts = BoardOverviewFlow.SubmissionCounts(
                [
                    new SubmissionStateCount(C64, SubmissionState.Merged, true, InStable: true, Count: 2),
                    new SubmissionStateCount(C64, SubmissionState.Merged, true, InStable: false, Count: 1),
                    new SubmissionStateCount(C64, SubmissionState.Pending, false, false, 1),
                    new SubmissionStateCount(C64, SubmissionState.Approved, false, false, 1),
                    new SubmissionStateCount(C64, SubmissionState.ChangesRequested, true, false, 1),
                    new SubmissionStateCount(C64, SubmissionState.Rejected, true, false, 1),
                    new SubmissionStateCount(C64, SubmissionState.Rejected, false, false, 4),
                ]);

            Assert.Equal(new BoardSubmissionCounts(Total: 7, Waiting: 3, Rejected: 1, InBeta: 1, InStable: 2), counts);
        }

        // ###########################################################################################
        // The store counts in SQL (code review, 2026-10-09: it read every submission of every board
        // on each minute's check); "in stable" is ContributorFacingState's own test - merged and
        // decided no later than the board's last publish to stable - which the fake keeps as the
        // query does.
        // ###########################################################################################
        [Fact]
        public async Task The_store_counts_merged_submissions_as_in_stable_up_to_the_last_publish_and_in_BETA_after_it()
        {
            DateTimeOffset published = BoardOverviewFlowTests.Now.AddDays(-10);
            var store = new FakeSubmissionStore();
            await store.EnsureBoardAsync(C64, "Commodore", "C64", "250407", "shipped", BoardOverviewFlowTests.Now);
            store.ProductionBoards[C64] = new PublishedBoardRow("2026-September-29", new string('a', 64), published);

            store.Submissions[1] = BoardOverviewFlowTests.Record(1, state: SubmissionState.Merged, decided: published.AddDays(-1));
            store.Submissions[2] = BoardOverviewFlowTests.Record(2, state: SubmissionState.Merged, decided: published);
            store.Submissions[3] = BoardOverviewFlowTests.Record(3, state: SubmissionState.Merged, decided: published.AddHours(1));

            IReadOnlyList<SubmissionStateCount> counts = await store.GetSubmissionStateCountsAsync();

            Assert.Equal(new BoardSubmissionCounts(3, 0, 0, 1, 2), BoardOverviewFlow.SubmissionCounts(counts));
        }

        // ###########################################################################################
        // Through the store: the list's every board carries its counts - one with no submission at
        // 0 - and the store leaves out uploads under way, abandoned ones and replaced ones.
        // ###########################################################################################
        [Fact]
        public async Task Every_board_in_the_list_carries_its_counts_and_the_store_leaves_out_what_nobody_reviewed()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await BoardOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            long id = 1;
            foreach (string state in new[] { SubmissionState.Pending, SubmissionState.Uploading, SubmissionState.Abandoned, SubmissionState.Withdrawn, SubmissionState.Merged })
            {
                store.Submissions[id] = BoardOverviewFlowTests.Record(id, state: state);
                id++;
            }

            BoardOverviewOutcome outcome = await BoardOverviewFlow.ListAsync(
                anna, BoardOverviewFlowTests.Beta, null, store, accounts, now: BoardOverviewFlowTests.Now);

            Assert.Equal(
                [(BoardOverviewFlowTests.C128, (BoardSubmissionCounts?)new BoardSubmissionCounts(0, 0, 0, 0, 0)), (BoardOverviewFlowTests.C64, new BoardSubmissionCounts(2, 1, 0, 1, 0))],
                outcome.Boards!.Select(board => (board.BoardId, board.SubmissionCounts)));
        }
}
}
