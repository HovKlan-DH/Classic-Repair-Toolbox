using CRT;
using Handlers.MaintainerHandling;
using Handlers.OnlineHandling;

namespace ClassicRepairToolbox.Tests.Maintainer;

// ###########################################################################################
// ReviewApiClient's own behaviour, as far as it can be seen without a network: the requests go to
// AnsweringHttpHandler, which answers from a function (test rule 6).
// ###########################################################################################
public sealed class ReviewApiClientTests
{
    private static readonly ReviewSession Session =
        new("token", DateTimeOffset.UtcNow.AddDays(1), 1, "m@example.com", "Maintainer");

    // ###########################################################################################
    // *** EVERY REQUEST NAMES THE CRT THAT SENT IT (2026-09-29). *** "CRT <version>", the
    // User-Agent CRT's own check-in and board views send - so the server's log and its session row
    // say which build a maintainer was on. The separate application sent none.
    // ###########################################################################################
    [Fact]
    public async Task Every_request_carries_CRTs_own_user_agent()
    {
        var seen = new List<string>();

        var handler = new AnsweringHttpHandler(request =>
        {
            seen.Add(request.Headers.UserAgent.ToString());
            return AnsweringHttpHandler.Refused();
        });

        using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress, new HttpClient(handler));

        await client.LoginAsync("m@example.com", "secret");
        await client.GetQueueAsync(ReviewApiClientTests.Session);

        Assert.Equal(2, seen.Count);
        Assert.All(seen, userAgent => Assert.Equal(OnlineServices.VersionForServer, userAgent));
    }

    // A client handed in that already names itself keeps its own name.
    [Fact]
    public async Task A_supplied_client_that_names_itself_keeps_its_name()
    {
        string? seen = null;

        var handler = new AnsweringHttpHandler(request =>
        {
            seen = request.Headers.UserAgent.ToString();
            return AnsweringHttpHandler.Refused();
        });

        var http = new HttpClient(handler);
        http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Something/1.0");

        using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress, http);

        await client.GetQueueAsync(ReviewApiClientTests.Session);

        Assert.Equal("Something/1.0", seen);
    }

    // The review client names the build exactly as the check-in and the board views do - the SAME
    // string, not an equal-looking copy (code review, 2026-09-29).
    [Fact]
    public void The_review_clients_user_agent_is_CRTs_one_definition()
    {
        Assert.Equal(OnlineServices.VersionForServer, ReviewApiClient.UserAgent);
        Assert.StartsWith(AppConfig.AppShortName + " ", ReviewApiClient.UserAgent, StringComparison.Ordinal);
    }
}
