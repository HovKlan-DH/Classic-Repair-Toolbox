using System.Text.Json;
using CRT.Server.Handlers.Usage;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers CheckInRules - the check-in's text rules, kept when the check-in moved to the service
    // (2026-10-03). crt_update is read by the Fun facts page and a nightly statistics job, so a
    // version that was stored before must still be stored as it was, one that was refused must
    // still be refused, and no markup may reach the table (the Fun facts page prints these
    // fields).
    // ###########################################################################################
    public sealed class CheckInRulesTests
    {
        [Theory]
        [InlineData("CRT 2026.10.0", "CRT 2026.10.0")]
        [InlineData("CRT 2.5.0-beta.1", "CRT 2.5.0-beta.1")]
        [InlineData("  CRT 2026.10.0 \n", "CRT 2026.10.0")]
        [InlineData("<b>CRT 2026.10.0</b>", "CRT 2026.10.0")]
        // "CRT " has always been looked for in any case, and stored as it was sent.
        [InlineData("crt 2026.10.0", "crt 2026.10.0")]
        public void A_CRT_version_is_kept_as_it_was_always_stored(string userAgent, string expected)
        {
            Assert.Equal(expected, CheckInRules.VersionFrom(userAgent));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("Mozilla/5.0 (Windows NT 10.0; Win64; x64)")]
        [InlineData("CRT")]
        [InlineData("CRT2026.10.0")]
        // Outside the version pattern: build metadata, an underscore, a semicolon, a quote.
        [InlineData("CRT 2026.10.0+abc123")]
        [InlineData("CRT 2026_10")]
        [InlineData("CRT 2026.10.0; DROP TABLE crt_update")]
        [InlineData("CRT 2026.10.0'")]
        public void Anything_else_is_no_version(string? userAgent)
        {
            Assert.Null(CheckInRules.VersionFrom(userAgent));
        }

        [Fact]
        public void A_version_longer_than_the_column_kept_for_it_is_no_version()
        {
            Assert.Null(CheckInRules.VersionFrom("CRT " + new string('1', 200)));
        }

        // ###########################################################################################
        // Markup stripped as it always has been, as far as these fields can carry: a tag goes, an
        // unclosed one takes the rest of the text with it, and a "<" that opens nothing stays.
        // ###########################################################################################
        [Theory]
        [InlineData("Windows", "Windows")]
        [InlineData("<b>Windows</b>", "Windows")]
        [InlineData("Linux <script>alert(1)</script>6.8", "Linux alert(1)6.8")]
        [InlineData("Windows <img src=x onerror=alert(1)", "Windows ")]
        [InlineData("5 < 6", "5 < 6")]
        [InlineData("ends with <", "ends with <")]
        [InlineData("a > b", "a > b")]
        [InlineData("", "")]
        public void Markup_is_stripped_as_it_always_was(string value, string expected)
        {
            Assert.Equal(expected, CheckInRules.StripTags(value));
        }

        [Fact]
        public void A_field_is_stripped_trimmed_and_cut_and_never_null()
        {
            Assert.Equal("Windows", CheckInRules.Text("  <i>Windows</i>  ", 20));
            Assert.Equal("Darwin", CheckInRules.Text("Darwin Kernel", 6));
            Assert.Equal(string.Empty, CheckInRules.Text(null, 20));
            Assert.Equal(string.Empty, CheckInRules.Text("<b></b>", 20));
        }

        // ###########################################################################################
        // apiJson holds the country the lookup gave - and nothing more, where it used to hold
        // ip-api.com's whole answer (city, coordinates, provider), which nothing reads.
        // ###########################################################################################
        [Fact]
        public void The_api_json_holds_the_country_or_is_empty()
        {
            Assert.Equal("{}", CheckInRules.ApiJson(null));
            Assert.Equal(
                """{"status":"success","countryCode":"DK","country":"Denmark"}""",
                CheckInRules.ApiJson(new CountryAnswer("DK", "Denmark")));
        }

        [Fact]
        public void A_country_name_with_quotes_or_accents_stays_valid_json()
        {
            string json = CheckInRules.ApiJson(new CountryAnswer("CI", "Côte d'Ivoire \"x\""));

            using JsonDocument document = JsonDocument.Parse(json);
            Assert.Equal("Côte d'Ivoire \"x\"", document.RootElement.GetProperty("country").GetString());
            Assert.Equal("CI", document.RootElement.GetProperty("countryCode").GetString());
        }
    }
}
