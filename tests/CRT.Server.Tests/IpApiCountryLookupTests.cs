using System.Net;
using CRT.Server.Handlers.Usage;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the country lookup's pure parts: which addresses are public (SenderAddress - only those
    // are looked up), the question put to ip-api.com, and how its answer is read. The lookup itself
    // is a network call (rule 6) and untested.
    //
    // WHY THE FIRST MATTERS: a private or local address places nobody, and behind Apache a request
    // from the server's own LAN - the project owner's machines - arrives with exactly such an
    // address. Its views take the server's OWN country instead, or are not stored at all
    // (ServerOptions.CountLocalNetworkBoardViews).
    // ###########################################################################################
    public sealed class IpApiCountryLookupTests
    {
        [Theory]
        [InlineData("127.0.0.1")]
        [InlineData("10.1.2.3")]
        [InlineData("172.16.0.1")]
        [InlineData("172.31.255.255")]
        [InlineData("192.168.20.10")]
        [InlineData("169.254.1.1")]
        [InlineData("100.64.0.1")]
        [InlineData("0.0.0.0")]
        [InlineData("224.0.0.1")]
        [InlineData("::1")]
        [InlineData("fe80::1")]
        [InlineData("fd12:3456::1")]
        [InlineData("::ffff:192.168.1.5")]
        public void A_local_or_private_address_is_not_looked_up(string address)
        {
            Assert.Null(SenderAddress.PublicOrNull(IPAddress.Parse(address)));
        }

        [Theory]
        [InlineData("85.184.162.75", "85.184.162.75")]
        [InlineData("172.32.0.1", "172.32.0.1")]
        [InlineData("100.128.0.1", "100.128.0.1")]
        [InlineData("2a00:1450:4001:80b::200e", "2a00:1450:4001:80b::200e")]
        [InlineData("::ffff:85.184.162.75", "85.184.162.75")]
        public void A_public_address_is_looked_up_as_itself(string address, string expected)
        {
            Assert.Equal(expected, SenderAddress.PublicOrNull(IPAddress.Parse(address))?.ToString());
        }

        [Fact]
        public void No_address_is_not_looked_up()
        {
            Assert.Null(SenderAddress.PublicOrNull(null));
        }

        // ###########################################################################################
        // The question asked. With an address, about that address; with NONE, ip-api.com answers
        // about the address asking - the server's own, which is where a sender on its own network is
        // too. A wrong URL here fails silently (no country), so it is pinned.
        // ###########################################################################################
        [Fact]
        public void A_sender_is_asked_about_by_its_address_and_the_server_by_none()
        {
            Assert.Equal(
                "http://ip-api.com/json/85.184.162.75?fields=status,countryCode,country",
                IpApiCountryLookup.QueryUrl(IPAddress.Parse("85.184.162.75")));

            Assert.Equal(
                "http://ip-api.com/json/?fields=status,countryCode,country",
                IpApiCountryLookup.QueryUrl(null));
        }

        [Fact]
        public void A_successful_answer_gives_the_code_in_capitals_and_the_name()
        {
            Assert.Equal(
                new CountryAnswer("DK", "Denmark"),
                IpApiCountryLookup.ReadAnswer("""{"status":"success","country":"Denmark","countryCode":"dk"}"""));
        }

        // ip-api.com says "fail" for a reserved or unknown address; anything else unusable is null
        // too, and never throws - a view is stored without a country instead.
        [Theory]
        [InlineData("""{"status":"fail","message":"reserved range"}""")]
        [InlineData("""{"status":"success","country":"Denmark"}""")]
        [InlineData("""{"status":"success","country":"Denmark","countryCode":"DNK"}""")]
        [InlineData("""{"status":"success","country":"Denmark","countryCode":"D1"}""")]
        [InlineData("""{"status":"success","country":"","countryCode":"DK"}""")]
        [InlineData("""{"country":"Denmark","countryCode":"DK"}""")]
        [InlineData("""[]""")]
        [InlineData("""not json""")]
        [InlineData("")]
        [InlineData(null)]
        public void An_unusable_answer_is_no_country(string? json)
        {
            Assert.Null(IpApiCountryLookup.ReadAnswer(json));
        }

        // The column holds 50 characters.
        [Fact]
        public void A_long_country_name_is_cut_to_its_column()
        {
            CountryAnswer? answer = IpApiCountryLookup.ReadAnswer(
                "{\"status\":\"success\",\"countryCode\":\"GS\",\"country\":\"" + new string('S', 80) + "\"}");

            Assert.Equal(50, answer!.Name.Length);
        }
    
        // ###########################################################################################
        // *** A FAILED LOOKUP PAUSES LOOKING UP, FOR EVERY ADDRESS (code review, 2026-09-29). ***
        // Failures were never remembered, so while ip-api.com was down every board-view request waited
        // its full 3-second timeout. After one failure the lookup is skipped - for any address - until
        // FailureBackoff has passed; then it asks again. Driven through the test seam: a fetch that
        // counts and fails, and a clock the test moves (no network - test rule 6).
        // ###########################################################################################
        [Fact]
        public async Task After_a_failure_no_address_is_looked_up_until_the_pause_has_passed()
        {
            int asked = 0;
            DateTimeOffset now = new(2026, 9, 29, 12, 0, 0, TimeSpan.Zero);

            var lookup = new IpApiCountryLookup(
                NullLogger<IpApiCountryLookup>.Instance,
                (_, _) =>
                {
                    asked++;
                    throw new HttpRequestException("ip-api.com did not answer");
                },
                () => now);

            Assert.Null(await lookup.LookupAsync(IPAddress.Parse("8.8.8.8")));
            Assert.Equal(1, asked);

            // Another address, a minute later: not asked - no 3-second wait.
            now = now.AddMinutes(1);
            Assert.Null(await lookup.LookupAsync(IPAddress.Parse("1.1.1.1")));
            Assert.Equal(1, asked);

            // Once the pause has passed, it asks again.
            now = now.Add(IpApiCountryLookup.FailureBackoff);
            Assert.Null(await lookup.LookupAsync(IPAddress.Parse("1.1.1.1")));
            Assert.Equal(2, asked);
        }

        // A lookup cancelled because the SENDER went away says nothing about ip-api.com, and must not
        // stop everybody else's views being placed.
        [Fact]
        public async Task A_lookup_the_sender_cancelled_pauses_nothing()
        {
            int asked = 0;
            using var gone = new CancellationTokenSource();
            gone.Cancel();

            var lookup = new IpApiCountryLookup(
                NullLogger<IpApiCountryLookup>.Instance,
                (_, token) =>
                {
                    asked++;
                    return token.IsCancellationRequested
                        ? Task.FromException<string>(new TaskCanceledException())
                        : Task.FromResult("{\"status\":\"success\",\"countryCode\":\"DK\",\"country\":\"Denmark\"}");
                },
                () => new DateTimeOffset(2026, 9, 29, 12, 0, 0, TimeSpan.Zero));

            Assert.Null(await lookup.LookupAsync(IPAddress.Parse("8.8.8.8"), gone.Token));
            Assert.Equal(new CountryAnswer("DK", "Denmark"), await lookup.LookupAsync(IPAddress.Parse("1.1.1.1")));
            Assert.Equal(2, asked);
        }
    }
}
