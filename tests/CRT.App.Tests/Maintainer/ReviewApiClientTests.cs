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

    // ###########################################################################################
    // *** AND THE API REVISION IT WAS BUILT FOR (2026-10-04). *** The server tells a CRT built for an
    // older revision to update - which it can only do if every request says which one. Sent on a
    // supplied client too, even one that names itself: the revision is this build's.
    // ###########################################################################################
    [Fact]
    public async Task Every_request_carries_the_api_revision_CRT_was_built_for()
    {
        var seen = new List<string?>();

        var handler = new AnsweringHttpHandler(request =>
        {
            seen.Add(request.Headers.TryGetValues(Handlers.DataHandling.ClientVersionContract.ApiRevisionHeader, out var values) ? values.Single() : null);
            return AnsweringHttpHandler.Refused();
        });

        var http = new HttpClient(handler);
        http.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "Something/1.0");

        using var client = new ReviewApiClient(ReviewApiRoutes.DefaultBaseAddress, http);

        await client.LoginAsync("m@example.com", "secret");
        await client.GetQueueAsync(ReviewApiClientTests.Session);
        await client.GetServerVersionAsync();

        Assert.Equal(3, seen.Count);
        Assert.All(seen, revision => Assert.Equal(
            Handlers.DataHandling.ClientVersionContract.ApiRevision,
            Handlers.DataHandling.ClientVersionContract.ApiRevisionFrom(revision)));
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

    // ###########################################################################################
    // The "Your account" window's changes (2026-10-03): each goes to its own route with the
    // session's token, and every refusal comes back as the sentence the server wrote - a code that
    // expired, every password rule broken at once - not a status code's generic words.
    // A 429 is "too many attempts" and a 401 is "sign in", as everywhere else.
    // ###########################################################################################
    [Fact]
    public async Task An_account_change_goes_to_its_route_with_the_sessions_token()
    {
        var seen = new List<(string Path, string? Token)>();

        var handler = new AnsweringHttpHandler(request =>
        {
            seen.Add((request.RequestUri!.AbsolutePath, request.Headers.Authorization?.Parameter));
            return AnsweringHttpHandler.Json("""{"message":"Done.","account":{"id":1,"email":"m@example.com","displayName":"M"}}""");
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        Assert.True((await client.ChangeNameAsync(ReviewApiClientTests.Session, "M")).IsOk);
        Assert.True((await client.RequestEmailChangeAsync(ReviewApiClientTests.Session, "n@example.com")).IsOk);
        Assert.True((await client.ConfirmEmailChangeAsync(ReviewApiClientTests.Session, "code")).IsOk);
        Assert.True((await client.ChangePasswordAsync(ReviewApiClientTests.Session, "new one here")).IsOk);

        Assert.Equal(
            [
                ("/api/accounts/me/name", "token"),
                ("/api/accounts/me/email", "token"),
                ("/api/accounts/me/email/confirm", "token"),
                ("/api/accounts/me/password", "token")
            ],
            seen);
    }

    [Theory]
    [InlineData(400, """{"message":"That code has expired. Ask for a new one."}""", ReviewApiFailure.Refused, "That code has expired. Ask for a new one.")]
    [InlineData(400, """{"errors":["At least 12 characters.","Not a common password."]}""", ReviewApiFailure.Refused, "At least 12 characters. Not a common password.")]
    [InlineData(429, "", ReviewApiFailure.RateLimited, "Too many attempts. Wait a moment and try again.")]
    [InlineData(401, "", ReviewApiFailure.NotSignedIn, "Sign in to continue.")]
    public async Task A_refused_account_change_says_what_the_server_said(int status, string body, ReviewApiFailure failure, string message)
    {
        var handler = new AnsweringHttpHandler(_ => new HttpResponseMessage((System.Net.HttpStatusCode)status)
        {
            Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json")
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        ReviewApiResult<Handlers.DataHandling.AccountChangeAnswer> result =
            await client.ChangePasswordAsync(ReviewApiClientTests.Session, "new one here");

        Assert.Equal(failure, result.Failure);
        Assert.Equal(message, result.Message);
    }

    // ###########################################################################################
    // *** "UPDATE CRT" (2026-10-04). *** A server that no longer serves this version answers 426 with
    // CRT.Data's ClientOutdatedAnswer. It reaches the maintainer as ClientOutdated, in the server's
    // words - on the JSON path and the bytes path alike, since both go through StatusFailureAsync.
    // ###########################################################################################
    [Fact]
    public async Task An_update_CRT_answer_comes_back_as_ClientOutdated_in_the_servers_words()
    {
        Handlers.DataHandling.ClientOutdatedAnswer answer = Handlers.DataHandling.ClientVersionContract.Outdated(
            Handlers.DataHandling.CrtVersion.Parse("3.0.0"), Handlers.DataHandling.CrtVersion.Parse("3.4.0"));

        string body = System.Text.Json.JsonSerializer.Serialize(answer, Handlers.DataHandling.ReviewApiContract.WireSettings);

        var handler = new AnsweringHttpHandler(_ => new HttpResponseMessage((System.Net.HttpStatusCode)426)
        {
            Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json")
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        ReviewApiResult<ReviewQueueResponse> queue = await client.GetQueueAsync(ReviewApiClientTests.Session);
        ReviewApiResult<byte[]> asset = await client.GetSubmittedAssetAsync(ReviewApiClientTests.Session, 4, new string('a', 64));

        Assert.Equal(ReviewApiFailure.ClientOutdated, queue.Failure);
        Assert.Equal(answer.Message, queue.Message);
        Assert.Equal(ReviewApiFailure.ClientOutdated, asset.Failure);
        Assert.Equal(answer.Message, asset.Message);
    }

    // ###########################################################################################
    // *** AND THE REST OF CRT IS TOLD (owner request, 2026-10-09). *** Every answer of this client
    // passes StatusFailureAsync, which raises ApiOutdatedSignal for the Maintainer tab on an "update
    // CRT" - Main then covers that tab until CRT is updated. Once per refused request, on both send
    // paths; an ordinary refusal raises nothing.
    // ###########################################################################################
    [Fact]
    public async Task An_update_CRT_answer_tells_the_rest_of_CRT_the_Maintainer_tab_is_turned_away()
    {
        Handlers.DataHandling.ClientOutdatedAnswer answer = Handlers.DataHandling.ClientVersionContract.OutdatedApi(
            Handlers.DataHandling.CrtVersion.Parse("3.0.0"));

        string body = System.Text.Json.JsonSerializer.Serialize(answer, Handlers.DataHandling.ReviewApiContract.WireSettings);
        bool outdated = true;

        var handler = new AnsweringHttpHandler(_ => outdated
            ? new HttpResponseMessage((System.Net.HttpStatusCode)426)
            {
                Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json")
            }
            : AnsweringHttpHandler.Refused());

        var raised = new List<(Handlers.Online.AppUpdateArea Area, string Words)>();

        void Listen(Handlers.Online.AppUpdateArea area, string words) => raised.Add((area, words));

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        Handlers.Online.ApiOutdatedSignal.Raised += Listen;

        try
        {
            await client.GetQueueAsync(ReviewApiClientTests.Session);
            await client.GetSubmittedAssetAsync(ReviewApiClientTests.Session, 4, new string('a', 64));

            outdated = false;
            await client.GetQueueAsync(ReviewApiClientTests.Session);
        }
        finally
        {
            Handlers.Online.ApiOutdatedSignal.Raised -= Listen;
        }

        Assert.Equal(
            [(Handlers.Online.AppUpdateArea.Maintainer, answer.Message), (Handlers.Online.AppUpdateArea.Maintainer, answer.Message)],
            raised);
    }

    // A 426 can mean nothing but "update CRT" - read so with no words of the server's too, as
    // SubmissionClient.OutdatedMessage reads one (2026-10-09). It was "The server answered 426."
    [Theory]
    [InlineData("")]
    [InlineData("<html>Upgrade Required</html>")]
    public async Task A_426_without_the_servers_words_is_still_ClientOutdated(string body)
    {
        var handler = new AnsweringHttpHandler(_ => new HttpResponseMessage((System.Net.HttpStatusCode)426)
        {
            Content = new StringContent(body)
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        ReviewApiResult<ReviewQueueResponse> queue = await client.GetQueueAsync(ReviewApiClientTests.Session);

        Assert.Equal(ReviewApiFailure.ClientOutdated, queue.Failure);
        Assert.Equal("This version of CRT is too old for the server - please update CRT.", queue.Message);
    }

    // ###########################################################################################
    // A refusal a newer server invents - a status this version has no case for - still reaches the
    // maintainer in the server's words, with the failure kind its status gives. A 401 stays "Sign in
    // to continue.": that kind is what sends the tab to its sign-in screen. A body with no sentence
    // keeps the client's own words.
    // ###########################################################################################
    [Theory]
    [InlineData(422, """{"message":"Something a newer server explains."}""", ReviewApiFailure.ServerError, "Something a newer server explains.")]
    [InlineData(403, """{"error":"This account does not maintain Commodore/C64/250407."}""", ReviewApiFailure.NotPermitted, "This account does not maintain Commodore/C64/250407.")]
    [InlineData(410, "", ReviewApiFailure.ServerError, "The server answered 410.")]
    [InlineData(500, """{"type":"about:blank","title":"Internal Server Error","status":500}""", ReviewApiFailure.ServerError, "The server answered 500.")]
    [InlineData(401, """{"message":"Your session has expired."}""", ReviewApiFailure.NotSignedIn, "Sign in to continue.")]
    public async Task A_refusal_says_what_the_server_said_whatever_its_status(int status, string body, ReviewApiFailure failure, string message)
    {
        var handler = new AnsweringHttpHandler(_ => new HttpResponseMessage((System.Net.HttpStatusCode)status)
        {
            Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json")
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        ReviewApiResult<ReviewQueueResponse> result = await client.GetQueueAsync(ReviewApiClientTests.Session);

        Assert.Equal(failure, result.Failure);
        Assert.Equal(message, result.Message);
    }

    // ###########################################################################################
    // The deployed server's version (2026-10-04): GET /api/health, and WITHOUT the bearer token -
    // the route is public, and the token is sent nowhere it is not needed.
    // ###########################################################################################
    [Fact]
    public async Task The_server_version_is_asked_of_the_health_route_with_no_token()
    {
        HttpRequestMessage? seen = null;

        var handler = new AnsweringHttpHandler(request =>
        {
            seen = request;
            return AnsweringHttpHandler.Json("{\"status\":\"ok\",\"version\":\"4.3.2\",\"utc\":\"2026-10-04T12:00:00+00:00\"}");
        });

        using var client = new ReviewApiClient("https://review.invalid", new HttpClient(handler));

        ReviewApiResult<Handlers.DataHandling.HealthStatus> result = await client.GetServerVersionAsync();

        Assert.True(result.IsOk, result.Message);
        Assert.Equal("4.3.2", result.Value!.Version);
        Assert.Null(result.Value.ApiRevision);
        Assert.Equal(HttpMethod.Get, seen!.Method);
        Assert.Equal("https://review.invalid/api/health", seen.RequestUri!.ToString());
        Assert.Null(seen.Headers.Authorization);
    }
}
