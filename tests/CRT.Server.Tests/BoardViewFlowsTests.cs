using System.Net;
using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers BoardViewFlows - storing a batch of board views from CRT (owner request, 2026-09-27:
    // "catch every time a user selects a board", "see usage of boards in countries").
    //
    // WHAT MUST HOLD: one row per countable view, carrying the listing's names (never the sender's),
    // the report's version/OS/CPU and the country; nothing identifying the sender; a batch counted
    // once however often it is sent; and a report CRT should drop is refused, while a view that
    // cannot be counted is only ignored - it never costs the rest of its batch.
    // ###########################################################################################
    public sealed class BoardViewFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";
        private const string Vic = "Commodore/VIC-20/250403";

        private static readonly IPAddress Sender = IPAddress.Parse("85.184.162.75");

        private static readonly CountryAnswer Denmark = new("DK", "Denmark");

        private static BoardViewNames? Listed(string systemId) =>
            systemId switch
            {
                BoardViewFlowsTests.C64 => new BoardViewNames("Commodore 64", "250407"),
                BoardViewFlowsTests.Vic => new BoardViewNames("Commodore VIC-20", "250403"),
                _ => null
            };

        private static BoardViewReport Report(params BoardView[] views) =>
            new(Guid.NewGuid().ToString("N"), "CRT 2026.10.0", "Windows", "Microsoft Windows 10.0.19045", "64-bit", views);

        private static BoardView View(string systemId, double minutesAgo = 5, bool fromBeta = false) =>
            new(systemId, BoardViewFlowsTests.Now.AddMinutes(-minutesAgo), fromBeta);

        private static Task<BoardViewOutcome> RecordAsync(BoardViewReport? report, FakeBoardViewStore store, FakeCountryLookup? countries = null) =>
            BoardViewFlowsTests.RecordFromAsync(BoardViewFlowsTests.Sender, report, store, countries);

        private static Task<BoardViewOutcome> RecordFromAsync(
            IPAddress sender,
            BoardViewReport? report,
            FakeBoardViewStore store,
            FakeCountryLookup? countries = null,
            bool countLocalNetwork = true) =>
            BoardViewFlows.RecordAsync(
                report,
                sender,
                BoardViewFlowsTests.Listed,
                countries ?? new FakeCountryLookup(BoardViewFlowsTests.Denmark),
                store,
                countLocalNetwork,
                BoardViewFlowsTests.Now);

        // ###########################################################################################
        // ONE ROW PER VIEW - the same board twice is two views (owner: "every time a user selects a
        // board"), each with everything the table was agreed to hold.
        // ###########################################################################################
        [Fact]
        public async Task Each_view_is_one_row_with_the_listings_names_the_machine_and_the_country()
        {
            var store = new FakeBoardViewStore();

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(
                    BoardViewFlowsTests.View(BoardViewFlowsTests.C64, 10),
                    BoardViewFlowsTests.View(BoardViewFlowsTests.C64, 3),
                    BoardViewFlowsTests.View(BoardViewFlowsTests.Vic, 1, fromBeta: true)),
                store);

            Assert.False(outcome.IsRefused);
            Assert.Equal((3, 0), (outcome.StoredViews, outcome.IgnoredViews));
            Assert.Equal(3, store.Rows.Count);

            BoardViewRow first = store.Rows[0];
            Assert.Equal(BoardViewFlowsTests.C64, first.SystemId);
            Assert.Equal(("Commodore 64", "250407"), (first.HardwareName, first.BoardName));
            Assert.Equal("CRT 2026.10.0", first.Version);
            Assert.Equal(("Windows", "Microsoft Windows 10.0.19045", "64-bit"), (first.OsHighlevel, first.OsVersion, first.Cpu));
            Assert.Equal(("DK", "Denmark"), (first.CountryCode, first.CountryName));
            Assert.Equal(BoardViewFlowsTests.Now.AddMinutes(-10), first.ViewedUtc);
            Assert.False(first.FromBeta);
            Assert.False(first.FromLocalNetwork);

            Assert.True(store.Rows[2].FromBeta);
            Assert.Equal("Commodore VIC-20", store.Rows[2].HardwareName);
        }

        // The country comes from the SENDER's address - the only thing the address is used for.
        [Fact]
        public async Task The_country_is_looked_up_once_per_report_from_the_senders_address()
        {
            var countries = new FakeCountryLookup(BoardViewFlowsTests.Denmark);

            await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64), BoardViewFlowsTests.View(BoardViewFlowsTests.Vic)),
                new FakeBoardViewStore(),
                countries);

            Assert.Equal(BoardViewFlowsTests.Sender, Assert.Single(countries.Asked));
            Assert.Equal(0, countries.AskedOwn);
        }

        // *** NOTHING STORED IDENTIFIES THE SENDER. *** Every text a row holds is the listing's, the
        // report's machine facts or the country - asserted by value, so a field added later that
        // carried the address would fail here.
        [Fact]
        public async Task No_stored_value_carries_the_senders_address()
        {
            var store = new FakeBoardViewStore();

            await BoardViewFlowsTests.RecordAsync(BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64)), store);

            BoardViewRow row = Assert.Single(store.Rows);

            string[] stored =
                typeof(BoardViewRow).GetProperties()
                    .Select(property => property.GetValue(row)?.ToString() ?? string.Empty)
                    .ToArray();

            Assert.DoesNotContain(stored, value => value.Contains("85.184", StringComparison.Ordinal));
        }

        // A failed lookup costs the country, never the view.
        [Fact]
        public async Task A_view_whose_country_cannot_be_told_is_stored_without_one()
        {
            var store = new FakeBoardViewStore();

            await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64)), store, new FakeCountryLookup(answer: null));

            BoardViewRow row = Assert.Single(store.Rows);
            Assert.Null(row.CountryCode);
            Assert.Null(row.CountryName);
        }

        // ###########################################################################################
        // *** A BATCH IS COUNTED ONCE. *** CRT sends the same batch again when it never heard the
        // answer; the second copy is accepted (so CRT stops sending it) and stores nothing.
        // ###########################################################################################
        [Fact]
        public async Task A_batch_sent_again_is_accepted_and_counts_nothing_twice()
        {
            var store = new FakeBoardViewStore();
            BoardViewReport report = BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64));

            BoardViewOutcome first = await BoardViewFlowsTests.RecordAsync(report, store);
            BoardViewOutcome again = await BoardViewFlowsTests.RecordAsync(report, store);

            Assert.False(first.IsRepeat);
            Assert.True(again.IsRepeat);
            Assert.False(again.IsRefused);
            Assert.Single(store.Rows);
        }

        // ###########################################################################################
        // *** ONLY PUBLISHED BOARDS ARE COUNTED. *** An id no listing has is ignored - which is what
        // keeps made-up text out of the table and off the Fun facts page - and so is a view too old
        // or too far ahead. Either way the rest of the batch is stored.
        // ###########################################################################################
        [Fact]
        public async Task Views_of_unlisted_boards_and_views_out_of_time_are_ignored_and_the_rest_stored()
        {
            var store = new FakeBoardViewStore();

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(
                    BoardViewFlowsTests.View("<script>alert(1)</script>/x/y"),
                    BoardViewFlowsTests.View("Commodore/C64/999999"),
                    new BoardView(BoardViewFlowsTests.C64, BoardViewFlowsTests.Now - BoardViewRules.MaxAge - TimeSpan.FromMinutes(1), false),
                    new BoardView(BoardViewFlowsTests.C64, BoardViewFlowsTests.Now + BoardViewRules.MaxAhead + TimeSpan.FromMinutes(1), false),
                    new BoardView(null, BoardViewFlowsTests.Now, false),
                    BoardViewFlowsTests.View(BoardViewFlowsTests.Vic)),
                store);

            Assert.Equal((1, 5), (outcome.StoredViews, outcome.IgnoredViews));
            Assert.Equal(BoardViewFlowsTests.Vic, Assert.Single(store.Rows).SystemId);
        }

        // A fast clock's view is kept, at the server's time - a view is never stored in the future.
        [Fact]
        public async Task A_view_from_a_clock_slightly_ahead_is_stored_at_the_servers_time()
        {
            var store = new FakeBoardViewStore();

            await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(new BoardView(BoardViewFlowsTests.C64, BoardViewFlowsTests.Now.AddHours(3), false)),
                store);

            Assert.Equal(BoardViewFlowsTests.Now, Assert.Single(store.Rows).ViewedUtc);
        }

        // ###########################################################################################
        // *** A REPORT FROM THE LOCAL NETWORK IS THE PROJECT OWNER'S OWN, AND COUNTS "FOR NOW" ***
        // (owner request, 2026-09-27: "For now I would like my own home usage also to count, as we
        // then can check the numbers and see how it works and looks like"). With the setting on, its
        // views are stored like anybody's, MARKED so they can be left out or deleted later, with the
        // country of the SERVER's own address - the sender's private address places nobody, and it
        // is never handed to the lookup.
        // ###########################################################################################
        [Theory]
        [InlineData("192.168.20.15")]
        [InlineData("10.0.0.2")]
        [InlineData("127.0.0.1")]
        public async Task A_report_from_a_local_network_is_stored_marked_with_the_servers_own_country(string address)
        {
            var store = new FakeBoardViewStore();
            var countries = new FakeCountryLookup(BoardViewFlowsTests.Denmark, own: FakeCountryLookup.Sweden);

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordFromAsync(
                IPAddress.Parse(address),
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64), BoardViewFlowsTests.View(BoardViewFlowsTests.Vic)),
                store,
                countries,
                countLocalNetwork: true);

            Assert.False(outcome.IsRefused);
            Assert.True(outcome.IsFromLocalNetwork);
            Assert.Equal((2, 0), (outcome.StoredViews, outcome.IgnoredViews));

            Assert.Equal(2, store.Rows.Count);
            Assert.All(store.Rows, row => Assert.True(row.FromLocalNetwork));
            Assert.All(store.Rows, row => Assert.Equal(("SE", "Sweden"), (row.CountryCode, row.CountryName)));

            Assert.Empty(countries.Asked);
            Assert.Equal(1, countries.AskedOwn);
        }

        // "For now" is the default: a server whose settings do not mention it counts home views, and
        // the switch that stops it is one line in appsettings.Production.json.
        [Fact]
        public void Home_views_count_unless_the_setting_turns_them_off()
        {
            Assert.True(new CRT.Server.Configuration.ServerOptions().CountLocalNetworkBoardViews);
        }

        // Home views are ordinary views in every other way: a batch sent again counts once.
        [Fact]
        public async Task A_batch_from_a_local_network_sent_again_counts_nothing_twice()
        {
            var store = new FakeBoardViewStore();
            BoardViewReport report = BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64));

            await BoardViewFlowsTests.RecordFromAsync(IPAddress.Parse("192.168.20.15"), report, store);
            BoardViewOutcome again = await BoardViewFlowsTests.RecordFromAsync(IPAddress.Parse("192.168.20.15"), report, store);

            Assert.True(again.IsRepeat);
            Assert.True(again.IsFromLocalNetwork);
            Assert.Single(store.Rows);
        }

        // ###########################################################################################
        // With the setting OFF - the launch check-in's rule - a report from the local network stores
        // nothing: accepted (CRT stops sending it), looked up nowhere, and said in the outcome so the
        // endpoint can log it.
        // ###########################################################################################
        [Theory]
        [InlineData("192.168.20.15")]
        [InlineData("10.0.0.2")]
        [InlineData("127.0.0.1")]
        public async Task With_local_counting_off_a_report_from_a_local_network_is_accepted_and_stores_nothing(string address)
        {
            var store = new FakeBoardViewStore();
            var countries = new FakeCountryLookup(BoardViewFlowsTests.Denmark);

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordFromAsync(
                IPAddress.Parse(address),
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64), BoardViewFlowsTests.View(BoardViewFlowsTests.Vic)),
                store,
                countries,
                countLocalNetwork: false);

            Assert.False(outcome.IsRefused);
            Assert.True(outcome.IsFromLocalNetwork);
            Assert.Equal((0, 2), (outcome.StoredViews, outcome.IgnoredViews));
            Assert.Empty(store.Rows);
            Assert.Empty(store.Batches);
            Assert.Empty(countries.Asked);
            Assert.Equal(0, countries.AskedOwn);
        }

        // The setting is about the LOCAL network only: a public sender is stored either way, unmarked.
        [Fact]
        public async Task With_local_counting_off_a_public_sender_is_still_stored()
        {
            var store = new FakeBoardViewStore();

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordFromAsync(
                BoardViewFlowsTests.Sender,
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64)),
                store,
                countLocalNetwork: false);

            Assert.False(outcome.IsFromLocalNetwork);
            Assert.False(Assert.Single(store.Rows).FromLocalNetwork);
        }

        // A malformed report is refused wherever it comes from, whichever way the setting is.
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public async Task A_malformed_report_from_a_local_network_is_still_refused(bool countLocalNetwork)
        {
            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordFromAsync(
                IPAddress.Loopback,
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64)) with { BatchId = "x" },
                new FakeBoardViewStore(),
                new FakeCountryLookup(),
                countLocalNetwork);

            Assert.True(outcome.IsRefused);
        }

        // With nothing to store, nothing is looked up and no batch is remembered.
        [Fact]
        public async Task A_report_with_nothing_countable_looks_nothing_up_and_stores_nothing()
        {
            var store = new FakeBoardViewStore();
            var countries = new FakeCountryLookup(BoardViewFlowsTests.Denmark);

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(BoardViewFlowsTests.View("Acme/Nothing/1")), store, countries);

            Assert.False(outcome.IsRefused);
            Assert.Empty(store.Rows);
            Assert.Empty(store.Batches);
            Assert.Empty(countries.Asked);
        }

        // ###########################################################################################
        // WHAT IS REFUSED - a report CRT should drop rather than send again: none, no usable batch
        // id, no version, no list, or more views than one report may carry.
        // ###########################################################################################
        [Fact]
        public async Task A_report_crt_should_drop_is_refused_and_nothing_is_stored()
        {
            var store = new FakeBoardViewStore();
            BoardView view = BoardViewFlowsTests.View(BoardViewFlowsTests.C64);
            BoardViewReport good = BoardViewFlowsTests.Report(view);

            BoardViewReport?[] refused =
            [
                null,
                good with { BatchId = null },
                good with { BatchId = "not-a-guid" },
                good with { BatchId = Guid.Empty.ToString("N") },
                good with { Version = "   " },
                good with { Views = null },
                good with { Views = Enumerable.Repeat(view, BoardViewRules.MaxViewsPerReport + 1).ToList() }
            ];

            foreach (BoardViewReport? report in refused)
            {
                BoardViewOutcome outcome = await BoardViewFlowsTests.RecordAsync(report, store);

                Assert.True(outcome.IsRefused);
                Assert.False(string.IsNullOrWhiteSpace(outcome.Reason));
            }

            Assert.Empty(store.Rows);
        }

        // The most one report may carry is accepted whole.
        [Fact]
        public async Task A_report_of_exactly_the_most_views_allowed_is_stored_whole()
        {
            var store = new FakeBoardViewStore();

            BoardViewOutcome outcome = await BoardViewFlowsTests.RecordAsync(
                BoardViewFlowsTests.Report(Enumerable.Repeat(BoardViewFlowsTests.View(BoardViewFlowsTests.C64), BoardViewRules.MaxViewsPerReport).ToArray()),
                store);

            Assert.Equal(BoardViewRules.MaxViewsPerReport, outcome.StoredViews);
        }

        // The sender's text is made fit to store: control characters out, cut to its column.
        [Fact]
        public async Task The_machine_facts_are_cleaned_and_cut_to_their_columns()
        {
            var store = new FakeBoardViewStore();

            BoardViewReport report = BoardViewFlowsTests.Report(BoardViewFlowsTests.View(BoardViewFlowsTests.C64)) with
            {
                OsHighlevel = "Windows\r\nX-Injected: yes and a great deal more",
                OsVersion = new string('v', 500),
                Cpu = "  "
            };

            await BoardViewFlowsTests.RecordAsync(report, store);

            BoardViewRow row = Assert.Single(store.Rows);
            Assert.Equal(BoardViewRules.OsHighlevelLength, row.OsHighlevel!.Length);
            Assert.DoesNotContain('\n', row.OsHighlevel);
            Assert.Equal(BoardViewRules.OsVersionLength, row.OsVersion!.Length);
            Assert.Null(row.Cpu);
        }
    }
}
