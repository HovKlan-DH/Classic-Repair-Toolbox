using System.Net;
using CRT.Server.Handlers.Usage;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers CheckInFlow - storing CRT's launch check-in in crt_update, the app-checkin PHP page's
    // job since 2026-10-03.
    //
    // WHAT MUST HOLD: a row the Fun facts pages count exactly as they counted the PHP page's - the
    // sender's public address as text (every chart counts DISTINCT ipaddr and filters
    // "192.168.%"), the version, the operating system and a two-letter country; nothing stored
    // from the server's own network; and anything not from CRT refused without a lookup.
    // ###########################################################################################
    public sealed class CheckInFlowTests
    {
        private static readonly IPAddress Sender = IPAddress.Parse("85.184.162.75");

        private static readonly CountryAnswer Denmark = new("DK", "Denmark");

        private static CheckInRequest Request(
            string? userAgent = "CRT 2026.10.0",
            string? control = CheckInContract.ControlValue,
            string? os = "Windows",
            string? osVersion = "Microsoft Windows 10.0.19045",
            string? cpu = "64-bit") =>
            new(userAgent, control, os, osVersion, cpu);

        private static Task<CheckInOutcome> RecordAsync(
            CheckInRequest request,
            FakeCheckInStore store,
            FakeCountryLookup? countries = null,
            IPAddress? sender = null) =>
            CheckInFlow.RecordAsync(
                request,
                sender ?? CheckInFlowTests.Sender,
                countries ?? new FakeCountryLookup(CheckInFlowTests.Denmark),
                store);

        [Fact]
        public async Task A_check_in_is_one_row_as_the_PHP_page_wrote_it()
        {
            var store = new FakeCheckInStore();

            CheckInOutcome outcome = await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(), store);

            Assert.True(outcome.IsStored);
            Assert.False(outcome.IsRefused);

            CheckInRow row = Assert.Single(store.Rows);
            Assert.Equal("85.184.162.75", row.IpAddress);
            Assert.Equal("CRT 2026.10.0", row.Version);
            Assert.Equal(("Windows", "Microsoft Windows 10.0.19045", "64-bit"), (row.OsHighlevel, row.OsVersion, row.Cpu));
            Assert.Equal(("DK", "Denmark"), (row.CountryCode, row.CountryName));
            Assert.Equal("""{"status":"success","countryCode":"DK","country":"Denmark"}""", row.ApiJson);
        }

        // An IPv4 sender seen through an IPv6 socket is stored as the plain IPv4 address - "::ffff:..."
        // would count as a second user in every DISTINCT ipaddr, and slip past "NOT LIKE '192.168.%'".
        [Fact]
        public async Task An_ipv4_address_carried_in_ipv6_is_stored_as_ipv4()
        {
            var store = new FakeCheckInStore();

            await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(), store, sender: IPAddress.Parse("::ffff:85.184.162.75"));

            Assert.Equal("85.184.162.75", Assert.Single(store.Rows).IpAddress);
        }

        [Fact]
        public async Task The_country_is_looked_up_for_the_senders_address()
        {
            var countries = new FakeCountryLookup(CheckInFlowTests.Denmark);

            await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(), new FakeCheckInStore(), countries);

            Assert.Equal([CheckInFlowTests.Sender], countries.Asked);
            Assert.Equal(0, countries.AskedOwn);
        }

        // A lookup that finds nothing costs the country, never the row - stored as the PHP page
        // stored it then: an empty code and name, and "{}".
        [Fact]
        public async Task A_failed_lookup_stores_the_row_without_a_country()
        {
            var store = new FakeCheckInStore();

            await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(), store, new FakeCountryLookup(answer: null));

            CheckInRow row = Assert.Single(store.Rows);
            Assert.Equal((string.Empty, string.Empty, "{}"), (row.CountryCode, row.CountryName, row.ApiJson));
        }

        // ###########################################################################################
        // The project owner's own machines: answered, so CRT is content, and never stored or looked
        // up - the PHP page's "192.168." rule, widened to every address that places nobody.
        // ###########################################################################################
        [Theory]
        [InlineData("192.168.30.11")]
        [InlineData("10.0.0.5")]
        [InlineData("127.0.0.1")]
        [InlineData("::ffff:192.168.30.11")]
        public async Task A_check_in_from_a_local_network_is_answered_and_not_stored(string address)
        {
            var store = new FakeCheckInStore();
            var countries = new FakeCountryLookup(CheckInFlowTests.Denmark);

            CheckInOutcome outcome = await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(), store, countries, IPAddress.Parse(address));

            Assert.False(outcome.IsRefused);
            Assert.False(outcome.IsStored);
            Assert.True(outcome.IsFromLocalNetwork);
            Assert.Empty(store.Rows);
            Assert.Empty(countries.Asked);
            Assert.Equal(0, countries.AskedOwn);
        }

        [Fact]
        public async Task No_sender_address_at_all_is_not_stored()
        {
            var store = new FakeCheckInStore();

            CheckInOutcome outcome = await CheckInFlow.RecordAsync(
                CheckInFlowTests.Request(), sender: null, new FakeCountryLookup(CheckInFlowTests.Denmark), store);

            Assert.True(outcome.IsFromLocalNetwork);
            Assert.Empty(store.Rows);
        }

        // ###########################################################################################
        // Not a CRT check-in: refused, with nothing stored and nobody looked up. The control field
        // must say exactly "CRT", as the PHP page compared it.
        // ###########################################################################################
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("crt")]
        [InlineData("CRT ")]
        [InlineData("Commodore Repair Toolbox")]
        public async Task A_check_in_whose_control_is_not_CRT_is_refused(string? control)
        {
            var store = new FakeCheckInStore();
            var countries = new FakeCountryLookup(CheckInFlowTests.Denmark);

            CheckInOutcome outcome = await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(control: control), store, countries);

            Assert.True(outcome.IsRefused);
            Assert.Empty(store.Rows);
            Assert.Empty(countries.Asked);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("curl/8.5.0")]
        [InlineData("CRT 2026.10.0+abc123")]
        public async Task A_check_in_without_a_CRT_version_is_refused(string? userAgent)
        {
            var store = new FakeCheckInStore();

            CheckInOutcome outcome = await CheckInFlowTests.RecordAsync(CheckInFlowTests.Request(userAgent: userAgent), store);

            Assert.True(outcome.IsRefused);
            Assert.Empty(store.Rows);
        }

        // The Fun facts page prints these fields: markup goes, a long value is cut to the length
        // board views keep, and a missing one is stored empty, as the PHP page stored it.
        [Fact]
        public async Task The_machine_fields_are_stripped_cut_and_never_missing()
        {
            var store = new FakeCheckInStore();

            await CheckInFlowTests.RecordAsync(
                CheckInFlowTests.Request(os: "<b>Linux</b>", osVersion: new string('x', 300), cpu: null),
                store);

            CheckInRow row = Assert.Single(store.Rows);
            Assert.Equal("Linux", row.OsHighlevel);
            Assert.Equal(BoardViewRules.OsVersionLength, row.OsVersion.Length);
            Assert.Equal(string.Empty, row.Cpu);
        }
    }
}
