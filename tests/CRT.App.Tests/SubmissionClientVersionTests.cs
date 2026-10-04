using System.Text.Json;
using Handlers.DataHandling;
using Handlers.Online;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What the Drafts tab's Submit tells the server about itself, and how it reads a refusal
// (2026-10-04, "Installed CRTs keep working").
//
// The server can tell a CRT too old for a changed API to update - in words - only if the request
// names its version, and the contributor sees those words only if the client shows them. Both were
// missing from submissions: the HttpClient sent no User-Agent at all, and a refusal read nothing
// but the findings list, so a 426 reached the contributor as "The server refused the submission
// (426)." No network here: the client is built and read, never sent (test rule 6).
// ###########################################################################################
public sealed class SubmissionClientVersionTests
{
    [Fact]
    public void Every_submission_request_names_the_CRT_that_sent_it()
    {
        using HttpClient http = SubmissionClient.CreateHttpClient(TimeSpan.FromSeconds(5));

        string userAgent = http.DefaultRequestHeaders.UserAgent.ToString();

        Assert.Equal(OnlineServices.VersionForServer, userAgent);
        Assert.NotNull(CrtVersion.FromUserAgent(userAgent));
    }

    // And the API revision it was built for (2026-10-04), so a server that has moved on says "update
    // CRT" instead of failing the submission some other way.
    [Fact]
    public void Every_submission_request_names_the_api_revision_it_was_built_for()
    {
        using HttpClient http = SubmissionClient.CreateHttpClient(TimeSpan.FromSeconds(5));

        Assert.True(http.DefaultRequestHeaders.TryGetValues(ClientVersionContract.ApiRevisionHeader, out IEnumerable<string>? values));
        Assert.Equal(ClientVersionContract.ApiRevision, ClientVersionContract.ApiRevisionFrom(values!.Single()));
    }

    // The answer for a CRT built for an older API revision reaches the contributor in its own words.
    [Fact]
    public void An_update_CRT_answer_for_the_api_revision_is_shown_in_the_servers_words()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.OutdatedApi(CrtVersion.Parse("3.0.0-alpha.2"));
        string body = JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings);

        Assert.Equal(answer.Message, SubmissionClient.RefusalMessage(body, "The server refused the submission (426)."));
    }

    // The answer the server writes for a CRT it no longer serves, read as the Submit dialog shows it.
    [Fact]
    public void An_update_CRT_answer_is_shown_in_the_servers_words()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.Outdated(CrtVersion.Parse("3.0.0"), CrtVersion.Parse("3.2.0"));
        string body = JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings);

        Assert.Equal(answer.Message, SubmissionClient.RefusalMessage(body, "The server refused the submission (426)."));
    }

    // A refusal a newer server invents carries a sentence; it must not be swapped for a status code.
    [Theory]
    [InlineData("""{"message":"This system is closed to contributions."}""", "This system is closed to contributions.")]
    [InlineData("""{"error":"Something the server explains."}""", "Something the server explains.")]
    public void Any_refusal_with_a_sentence_is_shown_in_the_servers_words(string body, string expected)
    {
        Assert.Equal(expected, SubmissionClient.RefusalMessage(body, "fallback"));
    }

    // Findings are shown as their own list beside the message, so a body of findings alone keeps
    // the client's own sentence rather than repeating them.
    [Theory]
    [InlineData("""{"errors":[{"code":"file.missing","subject":"U8","message":"Missing."}]}""")]
    [InlineData("")]
    [InlineData("<html>Bad gateway</html>")]
    public void A_refusal_without_a_sentence_keeps_the_clients_own(string body)
    {
        Assert.Equal("The server refused the submission (502).", SubmissionClient.RefusalMessage(body, "The server refused the submission (502)."));
    }

    // ###########################################################################################
    // *** THE STATUS CHECK READS "UPDATE CRT" AS ITSELF (code review, 2026-10-04). *** It read every
    // non-success answer as "unreachable", so a 426 left the receipts frozen with nothing said. The
    // server's own sentence is taken whenever the answer is one; a 426 with no words gets CRT's.
    // ###########################################################################################
    [Fact]
    public void An_update_CRT_answer_to_the_status_check_is_read_with_the_servers_words()
    {
        ClientOutdatedAnswer answer = ClientVersionContract.Outdated(CrtVersion.Parse("3.0.0"), CrtVersion.Parse("3.2.0"));
        string body = JsonSerializer.Serialize(answer, ReviewApiContract.WireSettings);

        Assert.Equal(answer.Message, SubmissionClient.OutdatedMessage(426, body));

        // The code alone says it, whatever the status line.
        Assert.Equal(answer.Message, SubmissionClient.OutdatedMessage(400, body));
    }

    [Theory]
    [InlineData("")]
    [InlineData("<html>Upgrade Required</html>")]
    public void A_426_without_the_servers_words_gets_CRTs_own(string body)
    {
        Assert.Equal(ClientOutdatedException.FallbackMessage, SubmissionClient.OutdatedMessage(426, body));
    }

    // Anything else is not "update CRT" - the status check keeps treating it as no answer.
    [Theory]
    [InlineData(500, "")]
    [InlineData(429, """{"error":"Too many requests."}""")]
    [InlineData(401, """{"message":"The token is not valid."}""")]
    public void Any_other_refusal_is_not_an_update_CRT_answer(int status, string body)
    {
        Assert.Null(SubmissionClient.OutdatedMessage(status, body));
    }

    // The resume query's answer is CRT.Data's BlobUploadAnswer as the server writes it.
    [Fact]
    public void The_resume_offset_is_read_from_the_servers_blob_answer()
    {
        string body = JsonSerializer.Serialize(new BlobUploadAnswer(123_456, false), ReviewApiContract.WireSettings);

        Assert.Equal(123_456, SubmissionClient.ReadUploaded(body));
        Assert.Equal(0, SubmissionClient.ReadUploaded("{}"));
    }
}
