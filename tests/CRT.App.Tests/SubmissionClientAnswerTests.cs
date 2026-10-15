using System.Net;
using System.Net.Http;
using System.Text;
using Handlers.DataHandling;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// How the Submit dialog reads a SUCCESSFUL answer (code review, 2026-10-13).
//
// A network that asks you to sign in first - a hotel's or a cafe's "captive portal" - answers
// every request 200 with an HTML page. ReadFromJsonAsync threw a JsonException out of the Submit
// dialog, which caught only network failures, and App's dispatcher handler let the crash take CRT
// down. An answer that is not the JSON it should be is now "unreadable", and every caller already
// turns that into a sentence. No network: responses are built in memory (test rule 6).
// ###########################################################################################
public sealed class SubmissionClientAnswerTests
{
    private static HttpResponseMessage Answer(string body, string mediaType) =>
        new(HttpStatusCode.OK) { Content = new StringContent(body, Encoding.UTF8, mediaType) };

    [Fact]
    public async Task A_sign_in_page_answering_200_is_unreadable_not_an_exception()
    {
        using HttpResponseMessage portal = Answer(
            "<!doctype html><html><body>Please sign in to the hotel wifi</body></html>", "text/html");

        Assert.Null(await SubmissionClient.TryReadAnswerAsync<SubmissionResult>(portal));
    }

    [Fact]
    public async Task An_empty_or_truncated_answer_is_unreadable()
    {
        using HttpResponseMessage empty = Answer(string.Empty, "application/json");
        using HttpResponseMessage truncated = Answer("{\"submissionId\": 12, \"isAcc", "application/json");

        Assert.Null(await SubmissionClient.TryReadAnswerAsync<SubmissionResult>(empty));
        Assert.Null(await SubmissionClient.TryReadAnswerAsync<HashNegotiationResponse>(truncated));
    }

    [Fact]
    public async Task A_real_answer_is_read_with_the_web_defaults()
    {
        // camelCase, as the server writes it - the defaults ReadFromJsonAsync used.
        using HttpResponseMessage answer = Answer("{\"submissionId\": 42, \"isAccepted\": true}", "application/json");

        SubmissionResult? result = await SubmissionClient.TryReadAnswerAsync<SubmissionResult>(answer);

        Assert.NotNull(result);
        Assert.Equal(42, result!.SubmissionId);
        Assert.True(result.IsAccepted);
    }

    [Fact]
    public void The_unreadable_reply_sentence_tells_the_contributor_what_to_do()
    {
        Assert.Contains("could not be read", SubmissionClient.UnreadableReply);
        Assert.Contains("nothing has been lost", SubmissionClient.UnreadableReply);
    }
}
