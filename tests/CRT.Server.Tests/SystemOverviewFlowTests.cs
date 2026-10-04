using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SystemOverviewFlow - the Maintainer tab's "Systems" screen (owner request, 2026-09-27):
    // every system, and one system's maintainers, contributors and recent submissions.
    //
    // THE RULE THAT MATTERS MOST is who may see it: ANY account that may review anything, for
    // EVERY system - contributor addresses included (owner decision, "Everything for everyone").
    // A maintainer of one board reads another board's facts; an account in no pool reads nothing.
    //
    // No database, no filesystem: the trees are handed in as lists, the way the endpoint hands them
    // in after PublishedSystemLister has read them.
    // ###########################################################################################
    public sealed class SystemOverviewFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";
        private const string Vic = "Commodore/VIC-20/250403";

        private static readonly IReadOnlyList<PublishedSystemLister.KnownSystem> Beta =
        [
            new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407"),
            new(SystemOverviewFlowTests.C128, "Commodore", "C128", "310378")
        ];

        private static readonly IReadOnlyList<PublishedSystemLister.KnownSystem> Production =
        [
            new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407")
        ];

        private static SubmissionRecord Record(
            long id,
            string? email = "dennis@example.com",
            long? account = null,
            string state = SubmissionState.Pending,
            string system = SystemOverviewFlowTests.C64,
            DateTimeOffset? created = null,
            DateTimeOffset? decided = null,
            string? comment = null) =>
            new(id, system, account, email, "hash", "", state, $"Submission {id}", 1,
                created ?? SystemOverviewFlowTests.Now.AddDays(-id), null, decided, comment);

        private static SystemSubmissionRecord Sent(SubmissionRecord record, bool byMaintainer = false) => new(record, byMaintainer);

        private static async Task<(FakeAccountStore Accounts, ReviewAccess Anna, long AnnaId, long BoId)> MaintainersAsync()
        {
            var accounts = new FakeAccountStore();

            long anna = await accounts.CreateAccountAsync(new NewAccount(
                "anna@example.com", "anna@example.com", "hash", "Anna", SystemOverviewFlowTests.Now));
            accounts.Accounts[anna] = accounts.Accounts[anna] with { IsVerified = true };

            long bo = await accounts.CreateAccountAsync(new NewAccount(
                "bo@example.com", "bo@example.com", "hash", "Bo", SystemOverviewFlowTests.Now));
            accounts.Accounts[bo] = accounts.Accounts[bo] with { IsVerified = true };

            // Anna maintains the C128 ONLY - the screen still shows her the C64.
            await accounts.AddMaintainerAsync(SystemOverviewFlowTests.C128, anna, 1, SystemOverviewFlowTests.Now);
            await accounts.AddMaintainerAsync(SystemOverviewFlowTests.C64, bo, 1, SystemOverviewFlowTests.Now);

            ReviewAccess access = ReviewAccess.For(
                accounts.Accounts[anna],
                await accounts.GetReviewedSystemIdsAsync(anna));

            return (accounts, access, anna, bo);
        }

        // -----------------------------------------------------------------------------------
        // Who may see it
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** EVERYTHING FOR EVERYONE (owner decision, 2026-09-27). *** Anna maintains only the C128,
        // and is shown the C64's contributors and their addresses all the same.
        // ###########################################################################################
        [Fact]
        public async Task A_maintainer_sees_the_contributors_of_a_system_they_do_not_maintain()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            await store.EnsureSystemAsync(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", SystemOverviewFlowTests.Now);
            store.Submissions[1] = SystemOverviewFlowTests.Record(1, "carl@example.com", state: SubmissionState.Merged);

            SystemOverviewOutcome outcome = await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, SystemOverviewFlowTests.Production, store, accounts);

            Assert.False(outcome.IsForbidden);
            Assert.Equal("carl@example.com", Assert.Single(outcome.Detail!.Contributors).Email);
            Assert.Equal("Bo", Assert.Single(outcome.Detail.Maintainers).DisplayName);
        }

        // An account in no pool may not review anything, and is refused both - the queue's own rule.
        [Fact]
        public async Task An_account_that_reviews_nothing_is_refused_the_list_and_the_detail()
        {
            var accounts = new FakeAccountStore();
            long id = await accounts.CreateAccountAsync(new NewAccount(
                "nobody@example.com", "nobody@example.com", "hash", "Nobody", SystemOverviewFlowTests.Now));
            accounts.Accounts[id] = accounts.Accounts[id] with { IsVerified = true };

            ReviewAccess nobody = ReviewAccess.For(accounts.Accounts[id]);
            var store = new FakeSubmissionStore();

            Assert.True((await SystemOverviewFlow.ListAsync(nobody, SystemOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
            Assert.True((await SystemOverviewFlow.DetailAsync(nobody, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
            Assert.True((await SystemOverviewFlow.ListAsync(null, SystemOverviewFlowTests.Beta, null, store, accounts)).IsForbidden);
        }

        // A system that is neither a row nor a board in the tree is not a system.
        [Fact]
        public async Task A_system_that_exists_nowhere_is_not_found()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();

            SystemOverviewOutcome outcome = await SystemOverviewFlow.DetailAsync(
                anna, "Acme/Nothing/1", SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts);

            Assert.True(outcome.IsNotFound);
        }

        // -----------------------------------------------------------------------------------
        // The list
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The rows AND the tree, once each: a shipped board nobody has touched has no row and is
        // still a system, and a new system waiting for its first review has a row and no board.
        // ###########################################################################################
        [Fact]
        public void The_list_is_every_row_and_every_board_in_the_tree_once_each()
        {
            SystemRecord[] rows =
            [
                new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", "2026-September-20", true, "beta-hash", "2026-May-14", "prod-hash", SystemOverviewFlowTests.Now.AddDays(-9)),
                new(SystemOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403", null, true)
            ];

            MaintainerRecord[] pool =
            [
                new(SystemOverviewFlowTests.C64, 1, "Bo", "bo@example.com"),
                new(SystemOverviewFlowTests.C64, 2, "Cy", "cy@example.com")
            ];

            IReadOnlyList<SystemOverviewEntry> list = SystemOverviewFlow.Entries(
                rows, SystemOverviewFlowTests.Beta, SystemOverviewFlowTests.Production, pool);

            Assert.Equal([SystemOverviewFlowTests.C128, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Vic], list.Select(entry => entry.SystemId));

            SystemOverviewEntry c64 = list.Single(entry => entry.SystemId == SystemOverviewFlowTests.C64);
            Assert.True(c64.InBeta);
            Assert.True(c64.InProduction);
            Assert.True(c64.IsAwaitingProduction);
            Assert.Equal("2026-September-20", c64.BetaRevision);
            Assert.Equal("2026-May-14", c64.ProductionRevision);
            Assert.Equal(2, c64.MaintainerCount);

            // In the BETA tree, no row: a shipped board. Not in production's list.
            SystemOverviewEntry c128 = list.Single(entry => entry.SystemId == SystemOverviewFlowTests.C128);
            Assert.True(c128.InBeta);
            Assert.False(c128.InProduction);
            Assert.False(c128.IsAwaitingProduction);
            Assert.Equal("C128", c128.Hardware);
            Assert.Equal(0, c128.MaintainerCount);

            // A row, no board anywhere: a new system nothing has published yet.
            SystemOverviewEntry vic = list.Single(entry => entry.SystemId == SystemOverviewFlowTests.Vic);
            Assert.False(vic.InBeta);
            Assert.False(vic.InProduction);
        }

        // ###########################################################################################
        // BETA's content hash rides on each system (2026-10-04) - the row's record of what BETA holds,
        // which every publish to BETA and every push-back moves. The Maintainer tab's Systems screen
        // reads its open table again when it moves; a board no row records has none to say.
        // ###########################################################################################
        [Fact]
        public void Each_system_carries_BETA_s_content_hash_from_its_row()
        {
            SystemRecord row = new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", "2026-September-20", true, "beta-hash", "2026-May-14", "prod-hash", null);

            Assert.Equal("beta-hash", SystemOverviewFlow.Entry(SystemOverviewFlowTests.C64, row, SystemOverviewFlowTests.Beta[0], null, 0).BetaContentHash);
            Assert.Null(SystemOverviewFlow.Entry(SystemOverviewFlowTests.C128, null, SystemOverviewFlowTests.Beta[1], null, 0).BetaContentHash);
        }

        // ###########################################################################################
        // *** WHAT THE STABLE SOURCE ALONE HOLDS OR LISTS IS ON THE LIST (owner request, 2026-10-04:
        // "I do not expect there should be cases where something can only be listed in stable? If so,
        // it must be flagged in the left-sided menu 'Systems' list"). *** Such a system used to be
        // missing from the screen altogether, so nothing could flag it. Each entry also says whether
        // each source's drop-down list names it.
        // ###########################################################################################
        [Fact]
        public void The_list_carries_what_each_drop_down_list_names_and_what_only_stable_has()
        {
            const string Pet = "Commodore/PET/2001";

            IReadOnlyList<PublishedSystemLister.KnownSystem> stableTree =
            [
                new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407"),
                new(SystemOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403")
            ];

            var listings = new SystemListings(
                new HashSet<string>([SystemOverviewFlowTests.C64], StringComparer.OrdinalIgnoreCase),
                new HashSet<string>([SystemOverviewFlowTests.C64, SystemOverviewFlowTests.C128, SystemOverviewFlowTests.Vic, Pet], StringComparer.OrdinalIgnoreCase));

            IReadOnlyList<SystemOverviewEntry> list = SystemOverviewFlow.Entries([], SystemOverviewFlowTests.Beta, stableTree, [], listings);

            Assert.Equal([SystemOverviewFlowTests.C128, SystemOverviewFlowTests.C64, Pet, SystemOverviewFlowTests.Vic], list.Select(entry => entry.SystemId));

            SystemOverviewEntry c128 = list.Single(entry => entry.SystemId == SystemOverviewFlowTests.C128);
            Assert.Equal((false, true), (c128.ListedInBeta, c128.ListedInStable));

            // In the stable source's tree only: on the list, named from that tree, not in BETA.
            SystemOverviewEntry vic = list.Single(entry => entry.SystemId == SystemOverviewFlowTests.Vic);
            Assert.False(vic.InBeta);
            Assert.True(vic.InProduction);
            Assert.Equal("VIC-20", vic.Hardware);

            // In the stable source's drop-down list only, no board anywhere: named from its id.
            SystemOverviewEntry pet = list.Single(entry => entry.SystemId == Pet);
            Assert.Equal(("Commodore", "PET", "2001"), (pet.Manufacturer, pet.Hardware, pet.Board));
            Assert.Equal((false, true), (pet.ListedInBeta, pet.ListedInStable));
        }

        // A list that could not be read says nothing - never "not listed" without looking.
        [Fact]
        public void A_list_that_could_not_be_read_says_nothing_about_listing()
        {
            IReadOnlyList<SystemOverviewEntry> list = SystemOverviewFlow.Entries(
                [], SystemOverviewFlowTests.Beta, null, [], new SystemListings(null, null));

            Assert.All(list, entry => Assert.Equal((null, null), (entry.ListedInBeta, entry.ListedInStable)));
        }

        // ###########################################################################################
        // The listings are read from each source's newest main Excel data file - the real files, by
        // system id in any case; a source with no such file is null. Read once per version of the
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

                SystemListings listings = SystemOverviewFlow.ReadListings(beta, stable);

                Assert.Equal([SystemOverviewFlowTests.C64], listings.Beta!);
                Assert.True(listings.Stable!.SetEquals([SystemOverviewFlowTests.C64, "commodore/c128/310378"]));
                Assert.Null(SystemOverviewFlow.ReadListings(beta, Path.Combine(root, "none")).Stable);

                // Rewritten with another system: read again.
                File.SetLastWriteTimeUtc(Path.Combine(beta, "Classic-Repair-Toolbox.v2.0.0.xlsx"), DateTime.UtcNow.AddDays(-1));
                DataTreeBuilder.Master(beta, "Commodore/C64/250407/Data C64 250407 v2.0.0.xlsx", "Commodore/C128/310378/Data C128 310378 v2.0.0.xlsx");

                Assert.Equal(2, SystemOverviewFlow.ReadListings(beta, null).Beta!.Count);
            }
            finally
            {
                Directory.Delete(root, recursive: true);
            }
        }

        // A system the stable source alone holds can be opened - it is on the list, and a list entry
        // that answers "no such system" would be a dead end.
        [Fact]
        public async Task A_system_only_the_stable_source_holds_has_a_detail()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();

            IReadOnlyList<PublishedSystemLister.KnownSystem> stableTree = [new(SystemOverviewFlowTests.Vic, "Commodore", "VIC-20", "250403")];

            SystemOverviewOutcome outcome = await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.Vic, SystemOverviewFlowTests.Beta, stableTree, new FakeSubmissionStore(), accounts,
                listings: new SystemListings(new HashSet<string>(), new HashSet<string>([SystemOverviewFlowTests.Vic])));

            Assert.False(outcome.IsNotFound);
            Assert.Equal(("VIC-20", false, true), (outcome.Detail!.System.Hardware, outcome.Detail.System.InBeta, outcome.Detail.System.ListedInStable));
        }

        // With no production tree to ask, the list never claims a board is not in production.
        [Fact]
        public void Without_a_production_tree_nothing_is_said_about_production()
        {
            IReadOnlyList<SystemOverviewEntry> list = SystemOverviewFlow.Entries([], SystemOverviewFlowTests.Beta, inProduction: null, []);

            Assert.All(list, entry => Assert.Null(entry.InProduction));
        }

        // A closed system says so; a system with no row has nothing closing it.
        [Fact]
        public void A_closed_system_reads_as_not_accepting()
        {
            SystemRecord closed = new(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", null, IsAccepting: false);

            Assert.False(SystemOverviewFlow.Entry(SystemOverviewFlowTests.C64, closed, null, null, 0).IsAccepting);
            Assert.True(SystemOverviewFlow.Entry(SystemOverviewFlowTests.C128, null, SystemOverviewFlowTests.Beta[1], null, 0).IsAccepting);
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
            SystemSubmissionRecord[] sent =
            [
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(1, "Dennis@Example.com", state: SubmissionState.Merged), true),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(2, " dennis@example.com ", state: SubmissionState.Pending)),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(3, "dennis@example.com", state: SubmissionState.Approved), true),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(4, "dennis@example.com", state: SubmissionState.ChangesRequested), true),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(5, "dennis@example.com", state: SubmissionState.Rejected), true),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(6, "dennis@example.com", state: SubmissionState.Rejected)),   // the automatic checks
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(7, "dennis@example.com", state: SubmissionState.Withdrawn))   // replaced
            ];

            SystemContributorEntry dennis = Assert.Single(SystemOverviewFlow.Contributors(sent));

            Assert.Equal(new SystemContributorEntry("Dennis@Example.com", null, Accepted: 1, Waiting: 2, ChangesRequested: 1, Rejected: 1,
                LastSubmittedUtc: SystemOverviewFlowTests.Now.AddDays(-1)), dennis);
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
            SystemSubmissionRecord[] sent =
            [
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(1, null, account: 7, created: SystemOverviewFlowTests.Now.AddDays(-1))),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(2, "anna@example.com", created: SystemOverviewFlowTests.Now.AddDays(-5))),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(3, null, account: 7, created: SystemOverviewFlowTests.Now.AddDays(-9)))
            ];

            var accounts = new Dictionary<long, AccountRecord>
            {
                [7] = new(7, "anna@example.com", "anna@example.com", "hash", "Anna", true, false, false, SystemOverviewFlowTests.Now, null)
            };

            IReadOnlyList<SystemContributorEntry> contributors = SystemOverviewFlow.Contributors(sent, accounts);

            Assert.Equal(2, contributors.Count);
            Assert.Equal(("anna@example.com", "Anna", 2), (contributors[0].Email, contributors[0].Name, contributors[0].Waiting));
            Assert.Equal(("anna@example.com", (string?)null, 1), (contributors[1].Email, contributors[1].Name, contributors[1].Waiting));
        }

        // A submission naming nobody cannot be anybody's, and is left out rather than listed blank.
        [Fact]
        public void A_submission_naming_nobody_is_no_contributor()
        {
            Assert.Empty(SystemOverviewFlow.Contributors([SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(1, "  "))]));
        }

        // -----------------------------------------------------------------------------------
        // Submissions
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE WORD THE CONTRIBUTOR IS TOLD. *** A merged submission decided after the system's
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
            DateTimeOffset promoted = SystemOverviewFlowTests.Now.AddDays(-3);

            SystemSubmissionRecord[] sent =
            [
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(1, state: SubmissionState.Merged, decided: promoted.AddDays(-1))),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(2, state: SubmissionState.Merged, decided: promoted.AddDays(1))),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(3, state: SubmissionState.Pending, comment: "U8 is wrong.")),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(4, state: SubmissionState.Pending)),
                SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(5, state: SubmissionState.Pending, comment: "A note, no rollback."))
            ];

            var returns = new Dictionary<long, DateTimeOffset> { [3] = SystemOverviewFlowTests.Now };

            IReadOnlyList<SystemSubmissionEntry> listed = SystemOverviewFlow.Submissions(sent, promoted, returns: returns);

            Assert.Equal([5L, 4L, 3L, 2L, 1L], listed.Select(entry => entry.Id));
            Assert.Equal(["pending", "pending", "returned", "merged", "published"], listed.Select(entry => entry.State));
            Assert.Equal("U8 is wrong.", listed[2].DecisionComment);
        }

        // Only the newest ListedSubmissions are listed; the contributors are counted over them all.
        [Fact]
        public void Only_the_newest_submissions_are_listed()
        {
            SystemSubmissionRecord[] sent = Enumerable.Range(1, SystemOverviewFlow.ListedSubmissions + 5)
                .Select(id => SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(id)))
                .ToArray();

            IReadOnlyList<SystemSubmissionEntry> listed = SystemOverviewFlow.Submissions(sent, null);

            Assert.Equal(SystemOverviewFlow.ListedSubmissions, listed.Count);
            Assert.Equal(SystemOverviewFlow.ListedSubmissions + 5, listed[0].Id);
            Assert.Equal(SystemOverviewFlow.ListedSubmissions + 5, Assert.Single(SystemOverviewFlow.Contributors(sent)).Waiting);
        }

        // A signed-in contributor's submission stores no address - their account's is shown.
        [Fact]
        public void A_signed_in_contributors_submission_shows_their_accounts_address()
        {
            var accounts = new Dictionary<long, AccountRecord>
            {
                [7] = new(7, "anna@example.com", "anna@example.com", "hash", "Anna", true, false, false, SystemOverviewFlowTests.Now, null)
            };

            SystemSubmissionEntry entry = Assert.Single(SystemOverviewFlow.Submissions(
                [SystemOverviewFlowTests.Sent(SystemOverviewFlowTests.Record(1, null, account: 7))], null, accounts));

            Assert.Equal("anna@example.com", entry.ContactEmail);
        }

        // -----------------------------------------------------------------------------------
        // Through the store
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The detail end to end over the fakes: only THIS system's submissions, never one still
        // uploading or abandoned (never anybody's to review), a maintainer's rejection counted and
        // the automatic checks' not, and the account's name looked up.
        // ###########################################################################################
        [Fact]
        public async Task The_detail_reads_only_this_systems_arrived_submissions()
        {
            (FakeAccountStore accounts, ReviewAccess anna, long annaId, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            await store.EnsureSystemAsync(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", SystemOverviewFlowTests.Now);

            store.Submissions[1] = SystemOverviewFlowTests.Record(1, "carl@example.com", state: SubmissionState.Rejected);
            store.Decisions[1] = new RecordedDecision(SubmissionState.Rejected, annaId, "No.", SystemOverviewFlowTests.Now);
            store.Submissions[2] = SystemOverviewFlowTests.Record(2, "carl@example.com", state: SubmissionState.Rejected);
            store.Submissions[3] = SystemOverviewFlowTests.Record(3, "carl@example.com", state: SubmissionState.Uploading);
            store.Submissions[4] = SystemOverviewFlowTests.Record(4, "carl@example.com", state: SubmissionState.Abandoned);
            store.Submissions[5] = SystemOverviewFlowTests.Record(5, "carl@example.com", system: SystemOverviewFlowTests.C128);
            store.Submissions[6] = SystemOverviewFlowTests.Record(6, null, account: annaId, state: SubmissionState.Merged);

            SystemDetailAnswer detail = (await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, SystemOverviewFlowTests.Production, store, accounts)).Detail!;

            Assert.Equal([6L, 2L, 1L], detail.Submissions.Select(entry => entry.Id));

            SystemContributorEntry carl = detail.Contributors.Single(contributor => contributor.Email == "carl@example.com");
            Assert.Equal(1, carl.Rejected);

            SystemContributorEntry annaAsContributor = detail.Contributors.Single(contributor => contributor.Name == "Anna");
            Assert.Equal(("anna@example.com", 1), (annaAsContributor.Email, annaAsContributor.Accepted));

            Assert.Equal(SystemOverviewFlowTests.C64, detail.System.SystemId);
            Assert.True(detail.System.InProduction);
        }

        // ###########################################################################################
        // What each submission changed as it went into BETA (owner request, 2026-10-04) rides on its
        // entry, for the History view - and one never published, or published before the record was
        // kept, carries none.
        // ###########################################################################################
        [Fact]
        public async Task The_detail_carries_what_each_published_submission_changed()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var store = new FakeSubmissionStore();

            await store.EnsureSystemAsync(SystemOverviewFlowTests.C64, "Commodore", "C64", "250407", "shipped", SystemOverviewFlowTests.Now);

            store.Submissions[1] = SystemOverviewFlowTests.Record(1, state: SubmissionState.Merged);
            store.Submissions[2] = SystemOverviewFlowTests.Record(2, state: SubmissionState.Merged);
            store.Submissions[3] = SystemOverviewFlowTests.Record(3);

            var changes = new SubmissionChanges(
                false,
                [new SectionChanges("Components", 0, 1, 0, 0, [], [new ChangedRowFact("U8", ["Part-number"])], [], [])],
                FileChanges.None);

            await store.SetChangesAsync(1, changes, SystemOverviewFlowTests.Now, CancellationToken.None);

            SystemDetailAnswer detail = (await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, null, store, accounts)).Detail!;

            Assert.Same(changes, detail.Submissions.Single(entry => entry.Id == 1).Changes);
            Assert.Null(detail.Submissions.Single(entry => entry.Id == 2).Changes);
            Assert.Null(detail.Submissions.Single(entry => entry.Id == 3).Changes);
        }

        // -----------------------------------------------------------------------------------
        // Board views (owner request, 2026-09-27)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The list's count is the last 30 whole days, BETA-source views left out - and a system nobody
        // looked at reads 0, not nothing: it was counted.
        // ###########################################################################################
        [Fact]
        public async Task The_list_carries_each_systems_views_in_the_last_30_days_beta_left_out()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore();

            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-1)));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-29)));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-31)));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-1), fromBeta: true));

            SystemOverviewOutcome outcome = await SystemOverviewFlow.ListAsync(
                anna, SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts,
                boardViews: views, now: SystemOverviewFlowTests.Now);

            Assert.Equal(
                [(SystemOverviewFlowTests.C128, (int?)0), (SystemOverviewFlowTests.C64, 2)],
                outcome.Systems!.Select(system => (system.SystemId, system.ViewsLast30Days)));
        }

        // The detail: the numbers, and the same 30-day count on its entry as the list shows.
        [Fact]
        public async Task The_detail_carries_the_systems_view_statistics()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore();

            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-1), "DK"));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-1), "DK"));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-20), "US"));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Now.AddDays(-2), "DE", fromBeta: true));
            views.Rows.Add(FakeBoardViewStore.Row(SystemOverviewFlowTests.C128, SystemOverviewFlowTests.Now.AddDays(-1), "SE"));

            SystemDetailAnswer detail = (await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts,
                now: SystemOverviewFlowTests.Now, boardViews: views)).Detail!;

            BoardViewStatistics stats = detail.Views!;
            Assert.Equal((2, 3, 3, 1), (stats.Last7Days, stats.Last30Days, stats.Last365Days, stats.FromBetaLast30Days));
            Assert.Equal(["Denmark 2", "United States 1"], stats.TopCountries.Select(country => $"{country.CountryName} {country.Views}"));
            Assert.Equal(3, detail.System.ViewsLast30Days);
        }

        // ###########################################################################################
        // *** THE VIEW COUNTS NEVER FAIL THE SCREEN. *** A database error reading them leaves them out
        // - the list and the detail still answer, with no numbers - rather than a 500 for everything.
        // ###########################################################################################
        [Fact]
        public async Task Views_that_cannot_be_read_are_left_out_and_the_screen_still_answers()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();
            var views = new FakeBoardViewStore { FailReads = true };

            SystemOverviewOutcome list = await SystemOverviewFlow.ListAsync(
                anna, SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts, boardViews: views);

            SystemOverviewOutcome detail = await SystemOverviewFlow.DetailAsync(
                anna, SystemOverviewFlowTests.C64, SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts, boardViews: views);

            Assert.All(list.Systems!, system => Assert.Null(system.ViewsLast30Days));
            Assert.Null(detail.Detail!.Views);
            Assert.Null(detail.Detail.System.ViewsLast30Days);
        }

        // Without a store (a caller that does not count views) there are no numbers either.
        [Fact]
        public async Task Without_a_view_store_there_are_no_view_numbers()
        {
            (FakeAccountStore accounts, ReviewAccess anna, _, _) = await SystemOverviewFlowTests.MaintainersAsync();

            SystemOverviewOutcome list = await SystemOverviewFlow.ListAsync(
                anna, SystemOverviewFlowTests.Beta, null, new FakeSubmissionStore(), accounts);

            Assert.All(list.Systems!, system => Assert.Null(system.ViewsLast30Days));
        }
    }
}
