using System.Text.Json;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for ClientVersionContract, CrtVersion and ApiRefusal - how a request names the CRT that
// sent it, and how CRT reads the server's "update CRT" answer and any other refusal.
//
// These decide whether an installed CRT is served or told to update (owner request, 2026-10-04:
// "all older versions will continue to work"), so the ordering rules are pinned case by case: a
// wrong comparison either locks out a CRT that should work, or lets through one the server can no
// longer serve - and either way the person at the bench sees something they cannot act on.
public class ClientVersionContractTests
{
    // -------------------------------------------------------------- CrtVersion.TryParse

    [Theory]
    [InlineData("3.0.0", "3.0.0")]
    [InlineData("2.4.0-beta.16", "2.4.0-beta.16")]
    [InlineData("3.0.0-alpha.2", "3.0.0-alpha.2")]
    [InlineData(" v2.5.0 ", "2.5.0")]
    [InlineData("3.0.0+abc1234", "3.0.0")]
    [InlineData("3.0.0-beta.1+abc1234", "3.0.0-beta.1")]
    public void Every_version_CRT_has_released_parses_and_prints_back(string text, string expected)
    {
        Assert.True(CrtVersion.TryParse(text, out CrtVersion? version));
        Assert.Equal(expected, version!.ToString());
    }

    // Anything that is not a three-part version is refused, so a stray token in a User-Agent can
    // never be taken for one.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("3")]
    [InlineData("3.0")]
    [InlineData("3.0.0.1")]
    [InlineData("3.x.0")]
    [InlineData("-1.0.0")]
    [InlineData("3.0.0-")]
    [InlineData("3.0.0-beta..1")]
    [InlineData("3.0.0-be ta")]
    [InlineData("(Windows)")]
    public void Anything_that_is_not_a_three_part_version_does_not_parse(string? text)
    {
        Assert.False(CrtVersion.TryParse(text, out CrtVersion? version));
        Assert.Null(version);
    }

    // -------------------------------------------------------------- CrtVersion ordering

    // Each pair is lower < higher. The pre-release cases are the ones a plain string or
    // System.Version comparison gets wrong - and CRT's betas and alphas are installed too.
    [Theory]
    [InlineData("2.4.0", "2.5.0")]
    [InlineData("2.5.0", "3.0.0")]
    [InlineData("2.9.0", "2.10.0")]
    [InlineData("3.0.9", "3.0.10")]
    [InlineData("3.0.0-alpha.2", "3.0.0")]
    [InlineData("3.0.0-alpha.2", "3.0.0-beta.1")]
    [InlineData("2.4.0-beta.2", "2.4.0-beta.16")]
    [InlineData("3.0.0-beta", "3.0.0-beta.1")]
    [InlineData("3.0.0-1", "3.0.0-alpha")]
    [InlineData("2.5.0", "3.0.0-alpha.1")]
    public void Versions_order_as_SemVer_orders_them(string lower, string higher)
    {
        CrtVersion low = CrtVersion.Parse(lower);
        CrtVersion high = CrtVersion.Parse(higher);

        Assert.True(low < high);
        Assert.True(high > low);
        Assert.True(low.CompareTo(high) < 0);
        Assert.True(high.CompareTo(low) > 0);
    }

    [Fact]
    public void A_version_equals_itself_however_it_was_written()
    {
        Assert.Equal(CrtVersion.Parse("3.0.0"), CrtVersion.Parse("v3.0.0+build"));
        Assert.True(CrtVersion.Parse("3.0.0-beta.1") >= CrtVersion.Parse("3.0.0-beta.1"));
        Assert.True(CrtVersion.Parse("3.0.0-beta.1") <= CrtVersion.Parse("3.0.0-beta.1"));
        Assert.Equal(CrtVersion.Parse("3.0.0").GetHashCode(), CrtVersion.Parse("v3.0.0+build").GetHashCode());
    }

    // ###########################################################################################
    // Equal versions hash alike (code review, 2026-10-04). Numeric identifiers compare as numbers,
    // so "beta.01" IS "beta.1" - and the hash used the raw text, so a set keyed by version held the
    // two as different versions.
    // ###########################################################################################
    [Theory]
    [InlineData("3.0.0-beta.01", "3.0.0-beta.1")]
    [InlineData("3.00.0", "3.0.0")]
    [InlineData("3.0.0-007", "3.0.0-7")]
    public void Equal_versions_written_differently_hash_alike(string one, string other)
    {
        CrtVersion first = CrtVersion.Parse(one);
        CrtVersion second = CrtVersion.Parse(other);

        Assert.Equal(first, second);
        Assert.Equal(first.GetHashCode(), second.GetHashCode());
        Assert.Single(new HashSet<CrtVersion> { first, second });
    }

    // -------------------------------------------------------------- CrtVersion.FromUserAgent

    [Theory]
    [InlineData("CRT 3.0.0-alpha.2", "3.0.0-alpha.2")]
    [InlineData("CRT 2.4.0", "2.4.0")]
    [InlineData("crt 2.5.0", "2.5.0")]
    [InlineData("CRT/3.1.0", "3.1.0")]
    [InlineData("SomeProxy/1.0 CRT 3.1.0 (Windows)", "3.1.0")]
    public void The_version_is_read_from_CRTs_user_agent(string userAgent, string expected)
    {
        Assert.Equal(expected, CrtVersion.FromUserAgent(userAgent)?.ToString());
    }

    // A browser opening a mailed link, curl, or a CRT whose version does not parse names no
    // version - and the server lets those through rather than refusing them.
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Mozilla/5.0 (Windows NT 10.0; Win64; x64)")]
    [InlineData("curl/8.4.0")]
    [InlineData("CRT")]
    [InlineData("CRT dev-build")]
    [InlineData("CRTX 3.0.0")]
    public void A_user_agent_naming_no_CRT_version_gives_none(string? userAgent)
    {
        Assert.Null(CrtVersion.FromUserAgent(userAgent));
    }

    [Fact]
    public void The_user_agent_CRT_sends_reads_back_as_its_version()
    {
        string userAgent = ClientVersionContract.UserAgentFor("3.0.0-alpha.2");

        Assert.Equal("CRT 3.0.0-alpha.2", userAgent);
        Assert.Equal(CrtVersion.Parse("3.0.0-alpha.2"), CrtVersion.FromUserAgent(userAgent));
    }

    // -------------------------------------------------------------- the outdated answer

    [Fact]
    public void The_outdated_answer_names_both_versions_and_carries_the_minimum()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.Outdated(
            CrtVersion.Parse("3.0.0"), CrtVersion.Parse("3.2.0-beta.1"));

        Assert.Equal(ClientVersionContract.OutdatedCode, answer.Code);
        Assert.Equal("3.2.0-beta.1", answer.MinimumVersion);
        Assert.Contains("[3.0.0]", answer.Message);
        Assert.Contains("[3.2.0-beta.1]", answer.Message);
        Assert.Contains("update CRT", answer.Message);
        Assert.Equal(426, ClientVersionContract.OutdatedStatus);
    }

    // THE ROUND TRIP BOTH ENDS DEPEND ON: the answer as the server serialises it (the wire
    // settings), read back by ApiRefusal as CRT reads it. A renamed field on either side fails here.
    [Fact]
    public void The_outdated_answer_as_the_server_writes_it_is_recognised_by_ApiRefusal()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.Outdated(
            CrtVersion.Parse("3.0.0"), CrtVersion.Parse("3.2.0"));

        string body = JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings);
        ApiRefusal refusal = ApiRefusal.Read(body);

        Assert.True(refusal.IsClientOutdated);
        Assert.Equal(answer.Message, refusal.Message);
        Assert.Equal("3.2.0", refusal.MinimumVersion);
    }

    // -------------------------------------------------------------- the API revision

    // ###########################################################################################
    // The refusal for a CRT built for an older API revision (2026-10-04). The server knows no version
    // that has its revision - only that the newest CRT does - so it names none: "update to the
    // newest", with the CRT's own version when the request named one. Through the wire and back as
    // CRT reads it, still an "update CRT".
    // ###########################################################################################
    [Fact]
    public void The_api_revision_refusal_says_update_to_the_newest_and_names_no_minimum()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.OutdatedApi(CrtVersion.Parse("3.0.0-alpha.2"));

        Assert.Equal(ClientVersionContract.OutdatedCode, answer.Code);
        Assert.Null(answer.MinimumVersion);
        Assert.Equal(
            "This version of CRT [3.0.0-alpha.2] was made for an older version of the server - please update CRT to the newest version.",
            answer.Message);

        ApiRefusal refusal = ApiRefusal.Read(JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings));

        Assert.True(refusal.IsClientOutdated);
        Assert.Equal(answer.Message, refusal.Message);
        Assert.Null(refusal.MinimumVersion);
    }

    [Fact]
    public void The_api_revision_refusal_for_a_request_naming_no_version_names_none()
    {
        Assert.Equal(
            "This version of CRT was made for an older version of the server - please update CRT to the newest version.",
            ClientVersionContract.OutdatedApi(null).Message);
    }

    // What CRT sends - the revision as a plain number - reads back; anything else names no revision,
    // which the server lets through like a request naming no version.
    [Theory]
    [InlineData("1", 1)]
    [InlineData(" 12 ", 12)]
    [InlineData(null, null)]
    [InlineData("", null)]
    [InlineData("0", null)]
    [InlineData("-3", null)]
    [InlineData("+3", null)]
    [InlineData("2.5", null)]
    [InlineData("two", null)]
    [InlineData("99999999999", null)]
    public void The_api_revision_is_read_from_its_header(string? header, int? expected)
    {
        Assert.Equal(expected, ClientVersionContract.ApiRevisionFrom(header));
    }

    [Fact]
    public void The_api_revision_CRT_sends_reads_back()
    {
        Assert.True(ClientVersionContract.ApiRevision >= 1);
        Assert.Equal(
            ClientVersionContract.ApiRevision,
            ClientVersionContract.ApiRevisionFrom(ClientVersionContract.ApiRevision.ToString(System.Globalization.CultureInfo.InvariantCulture)));
    }

    // -------------------------------------------------------------- ApiRefusal.Read

    [Theory]
    [InlineData("""{"message":"Write at least 10 characters."}""", "Write at least 10 characters.")]
    [InlineData("""{"error":"This system is closed to contributions."}""", "This system is closed to contributions.")]
    [InlineData("""{"message":"  Spaced.  "}""", "Spaced.")]
    [InlineData("""{"message":"First.","error":"Second."}""", "First.")]
    [InlineData("""{"message":"   ","error":"The error, since the message is blank."}""", "The error, since the message is blank.")]
    public void The_servers_sentence_is_read_from_message_or_error(string body, string expected)
    {
        ApiRefusal refusal = ApiRefusal.Read(body);

        Assert.Equal(expected, refusal.Message);
        Assert.False(refusal.IsClientOutdated);
    }

    // The submission path shows "errors" as a list of findings beside the message; folding them
    // into the message too would show every finding twice.
    [Fact]
    public void An_errors_list_is_not_made_into_a_message()
    {
        ApiRefusal refusal = ApiRefusal.Read("""{"errors":[{"code":"x","message":"A finding."}]}""");

        Assert.Null(refusal.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("Success")]
    [InlineData("<html>502 Bad Gateway</html>")]
    [InlineData("[1,2,3]")]
    [InlineData("""{"message":42}""")]
    [InlineData("""{"message":""" )]
    public void A_body_saying_nothing_readable_gives_an_empty_refusal_and_never_throws(string? body)
    {
        ApiRefusal refusal = ApiRefusal.Read(body);

        Assert.Null(refusal.Code);
        Assert.Null(refusal.Message);
        Assert.Null(refusal.MinimumVersion);
        Assert.False(refusal.IsClientOutdated);
    }

    // Only the exact code means "update CRT" - a similar-looking one is an ordinary refusal.
    [Theory]
    [InlineData("CLIENT.OUTDATED")]
    [InlineData("client.outdated.soon")]
    [InlineData("outdated")]
    public void Only_the_exact_outdated_code_is_taken_as_update_CRT(string code)
    {
        ApiRefusal refusal = ApiRefusal.Read($$"""{"code":"{{code}}","message":"Something."}""");

        Assert.False(refusal.IsClientOutdated);
        Assert.Equal("Something.", refusal.Message);
    }

    // -------------------------------------------------------------- tolerance

    // *** CLIENTS MUST IGNORE WHAT THEY DO NOT KNOW. *** A newer server adds fields to its answers;
    // an installed CRT reading them with the wire settings must not fail on one it has never seen.
    // System.Text.Json skips unknown members by default - this holds the shared settings to it, so
    // nobody can switch on UnmappedMemberHandling.Disallow without this failing.
    [Fact]
    public void The_wire_settings_skip_fields_a_reader_does_not_know()
    {
        Assert.Equal(
            System.Text.Json.Serialization.JsonUnmappedMemberHandling.Skip,
            ReviewApiContract.WireSettings.UnmappedMemberHandling);

        ClientOutdatedAnswer? answer = JsonSerializer.Deserialize<ClientOutdatedAnswer>(
            """{"code":"client.outdated","message":"M","minimumVersion":"3.2.0","addedLater":{"x":[1]}}""",
            ReviewApiContract.WireSettings);

        Assert.Equal("M", answer!.Message);
    }

    [Fact]
    public void Reading_a_refusal_ignores_fields_it_does_not_know()
    {
        ApiRefusal refusal = ApiRefusal.Read("""{"message":"M","retryAfterSeconds":30,"details":[{"a":1}]}""");

        Assert.Equal("M", refusal.Message);
        Assert.Null(refusal.Code);
    }
}
