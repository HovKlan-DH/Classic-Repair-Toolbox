using System.Net;
using System.Text;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// A server that answers every request from a function - no socket, no network (test rule 6).
// Handed to ReviewApiClient's HttpClient, so the Maintainer tab's real request path runs to the point of
// sending and gets a chosen answer back.
//
// The one-argument form runs on the CALLING thread, synchronously: HttpClient reaches the handler
// before its first await, and it never awaits. So a test on the UI thread can read controls from
// inside it - which is how it sees what the window looked like WHILE the request was in flight.
// The two-argument form may answer later, or never (NeverAsync) - a server busy past the limit.
// ###########################################################################################
internal sealed class AnsweringHttpHandler : HttpMessageHandler
{
    private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> thisAnswer;

    public AnsweringHttpHandler(Func<HttpRequestMessage, HttpResponseMessage> answer) =>
        this.thisAnswer = (request, _) => Task.FromResult(answer(request));

    public AnsweringHttpHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> answer) =>
        this.thisAnswer = answer;

    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) =>
        this.thisAnswer(request, cancellationToken);

    // A refusal the client reports as a failure, with no further request after it.
    internal static HttpResponseMessage Refused() =>
        new(HttpStatusCode.BadRequest)
        {
            Content = new StringContent("{\"error\":\"refused for the test\"}", Encoding.UTF8, "application/json")
        };

    internal static HttpResponseMessage Json(string json) =>
        new(HttpStatusCode.OK)
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };

    // No answer at all until the request is given up - the server still working when the window's
    // two minutes pass.
    internal static async Task<HttpResponseMessage> NeverAsync(CancellationToken token)
    {
        await Task.Delay(Timeout.Infinite, token);
        throw new InvalidOperationException("Unreachable: the delay only ends by being cancelled.");
    }
}
