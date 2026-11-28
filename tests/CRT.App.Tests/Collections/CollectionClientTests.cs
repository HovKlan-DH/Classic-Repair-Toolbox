using System.Net;
using System.Text;
using System.Text.Json;
using ClassicRepairToolbox.Tests.Maintainer;
using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// Sharing a collection with the server (CollectionClient), on a fake server answering in CollectionContract's records as the
// server writes them (ReviewApiContract.WireSettings) - so a field renamed on one side fails here. Every request carries the
// collection's KEY, and nothing else says who shares (owner decision, 2026-11-28: "Anybody, no account").
// ###########################################################################################
public sealed class CollectionClientTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private const string Root = "https://classic-repair-toolbox.dk/crt-test";

    private static readonly string ShaA = new('a', 64);

    private static readonly string ShaB = new('b', 64);

    private static readonly string Key = new('c', CollectionRules.OwnerKeyLength);

    private static LocalCollection Collection(string? serverId = null) =>
        new()
        {
            LocalId = new string('0', 32),
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Title = "No picture",
            CreatedUtc = Now,
            OwnerKey = Key,
            ServerId = serverId,
            DeleteAtUtc = serverId is null ? null : Now.AddYears(1),
            Captures =
            [
                new LocalCapture { Sha256 = ShaA, ComponentLabel = "U7", Pin = "12", CapturedUtc = Now },
                new LocalCapture { Sha256 = ShaB, ComponentLabel = "U8", Pin = "1", CapturedUtc = Now }
            ]
        };

    private static HttpResponseMessage Answer(HttpStatusCode status, object body) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(body, ReviewApiContract.WireSettings), Encoding.UTF8, "application/json") };

    private static CollectionAnswer Shared(string id = "k7mq3x9vbn2p") =>
        new(id, $"https://classic-repair-toolbox.dk/crt-test/collections/{id}", Now, Now.AddYears(1), 2);

    private sealed class Server
    {
        public List<(HttpMethod Method, string Path, string? Key, bool Bearer, byte[] Body, string? ContentType)> Requests { get; } = [];

        public Func<HttpRequestMessage, int, HttpResponseMessage> Answer { get; set; } = (_, _) => new HttpResponseMessage(HttpStatusCode.InternalServerError);

        public CollectionClient Client() => new(Root, new HttpClient(new AnsweringHttpHandler(async (request, token) =>
        {
            byte[] body = request.Content is null ? [] : await request.Content.ReadAsByteArrayAsync(token);
            string? key = request.Headers.TryGetValues(CollectionRoutes.OwnerKeyHeader, out IEnumerable<string>? values) ? values.Single() : null;

            this.Requests.Add((request.Method, request.RequestUri!.AbsolutePath, key, request.Headers.Authorization is not null, body, request.Content?.Headers.ContentType?.MediaType));

            return this.Answer(request, this.Requests.Count);
        })));
    }

    private static Func<string, byte[]?> Captures(Dictionary<string, byte[]> bytes) => sha => bytes.TryGetValue(sha, out byte[]? png) ? png : null;

    [Fact]
    public async Task A_first_share_sends_the_whole_collection_under_its_key_and_keeps_what_it_sent()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Created, Shared()) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Shared, result.Kind);
        Assert.Equal("k7mq3x9vbn2p", result.Answer!.Id);

        (HttpMethod method, string path, string? key, bool bearer, byte[] body, _) = Assert.Single(server.Requests);

        Assert.Equal((HttpMethod.Post, "/crt-test/api/collections", Key, false), (method, path, key, bearer));

        // What the server binds - the very record, by the wire's names.
        CollectionContentRequest sent = JsonSerializer.Deserialize<CollectionContentRequest>(body, ReviewApiContract.WireSettings)!;

        Assert.Equal(["U7", "U8"], sent.Images.Select(image => image.ComponentLabel));
        Assert.Equal(Now.AddYears(1), sent.DeleteAtUtc);
        Assert.Equal(CollectionSharing.ContentHash(sent), result.SharedContentHash);

        // The date SENT is what the collection keeps - never the server's echo, whose database keeps milliseconds only.
        Assert.Equal(sent.DeleteAtUtc, result.SentDeleteAtUtc);
    }

    [Fact]
    public async Task The_captures_the_server_asks_for_are_sent_under_the_key_and_then_the_collection_again()
    {
        var server = new Server
        {
            Answer = (request, count) => count switch
            {
                1 => Answer(HttpStatusCode.Conflict, new CollectionRefusalAnswer(["The server does not hold 1 of the captures yet"], [ShaB])),
                2 => Answer(HttpStatusCode.OK, new CollectionImageUploadAnswer(ShaB, 800, 480)),
                _ => Answer(HttpStatusCode.Created, Shared())
            }
        };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), Captures(new() { [ShaB] = [7, 7, 7] }), Now);

        Assert.Equal(CollectionClient.Outcome.Shared, result.Kind);
        Assert.Equal(3, server.Requests.Count);

        var upload = server.Requests[1];

        Assert.Equal((HttpMethod.Put, $"/crt-test/api/collections/images/{ShaB}", Key, "image/png"), (upload.Method, upload.Path, upload.Key, upload.ContentType));
        Assert.Equal([7, 7, 7], upload.Body);
        Assert.Equal(HttpMethod.Post, server.Requests[2].Method);
    }

    // A collection with no key of the right shape - one shared with an account before 2026-11-28 - is the window's to give one
    // (CollectionWindow.WithOwnerKey); the client sends nothing without.
    [Theory]
    [InlineData(null)]
    [InlineData("not-a-key")]
    public async Task A_collection_with_no_key_sends_nothing(string? key)
    {
        var server = new Server();

        CollectionClient.Result result = await server.Client().ShareAsync(Collection() with { OwnerKey = key }, _ => null, Now);

        Assert.Equal((CollectionClient.Outcome.Refused, CollectionClient.NoKey), (result.Kind, result.Message));
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task A_capture_no_longer_on_this_computer_is_said_and_nothing_more_is_sent()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Conflict, new CollectionRefusalAnswer(["missing"], [ShaA])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Contains("no longer on this computer", result.Message, StringComparison.Ordinal);
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task A_collection_shared_before_is_sent_to_its_page_and_one_the_server_deleted_is_said()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2027-November-26."])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection("k7mq3x9vbn2p"), _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.GoneFromServer, result.Kind);
        Assert.Equal("That collection was deleted on 2027-November-26.", result.Message);
        Assert.Equal((HttpMethod.Put, "/crt-test/api/collections/k7mq3x9vbn2p", Key), (server.Requests[0].Method, server.Requests[0].Path, server.Requests[0].Key));
    }

    // Under this key the server knows no such page any more (404) - forgotten as one gone, so "Share" makes a new one.
    [Fact]
    public async Task A_page_the_server_does_not_know_under_this_key_is_gone()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.NotFound, new CollectionRefusalAnswer(["There is no collection with that id that this computer shared."])) };

        Assert.Equal(CollectionClient.Outcome.GoneFromServer, (await server.Client().ShareAsync(Collection("k7mq3x9vbn2p"), _ => null, Now)).Kind);
    }

    [Fact]
    public async Task A_refusal_is_said_in_the_server_s_words()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.TooManyRequests, new CollectionRefusalAnswer(["An internet connection shares at most 20 new collections a day."])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Equal("An internet connection shares at most 20 new collections a day.", result.Message);
    }

    [Fact]
    public async Task A_server_out_of_reach_is_said_in_words()
    {
        var client = new CollectionClient(Root, new HttpClient(new AnsweringHttpHandler((_, _) => throw new HttpRequestException("No such host"))));

        CollectionClient.Result result = await client.ShareAsync(Collection(), _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Failed, result.Kind);
        Assert.Equal(CollectionClient.NotReachable, result.Message);
    }

    // The rules both ends keep are asked before anything is sent.
    [Fact]
    public async Task A_collection_breaking_a_rule_is_refused_without_a_request()
    {
        var server = new Server();

        CollectionClient.Result result = await server.Client().ShareAsync(Collection() with { Captures = [] }, _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Contains("A collection holds at least one capture.", result.Errors);
        Assert.Empty(server.Requests);
    }

    // ---- The page's other actions, each under the key ------------------------------------------------------------------

    [Theory]
    [InlineData(HttpStatusCode.NoContent, CollectionClient.Outcome.Deleted)]
    [InlineData(HttpStatusCode.NotFound, CollectionClient.Outcome.Deleted)]
    [InlineData(HttpStatusCode.Gone, CollectionClient.Outcome.Deleted)]
    public async Task Deleting_a_page_is_done_also_when_the_server_has_none(HttpStatusCode status, CollectionClient.Outcome expected)
    {
        var server = new Server { Answer = (_, _) => new HttpResponseMessage(status) };

        CollectionClient.Result result = await server.Client().DeleteAsync("k7mq3x9vbn2p", Key);

        Assert.Equal(expected, result.Kind);
        Assert.Equal((HttpMethod.Delete, "/crt-test/api/collections/k7mq3x9vbn2p", Key), (server.Requests[0].Method, server.Requests[0].Path, server.Requests[0].Key));
    }

    [Fact]
    public async Task Stopping_sharing_and_deleting_an_image_go_to_their_routes_under_the_key()
    {
        var server = new Server { Answer = (_, _) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent("{}") } };
        CollectionClient client = server.Client();

        Assert.Equal(CollectionClient.Outcome.Done, (await client.SetSharingAsync("k7mq3x9vbn2p", false, Key)).Kind);
        Assert.Equal(CollectionClient.Outcome.Done, (await client.RemoveImageAsync("k7mq3x9vbn2p", ShaA, Key)).Kind);

        Assert.Equal(
            [
                (HttpMethod.Put, "/crt-test/api/collections/k7mq3x9vbn2p/sharing", Key),
                (HttpMethod.Delete, $"/crt-test/api/collections/k7mq3x9vbn2p/images/{ShaA}", Key)
            ],
            server.Requests.Select(request => (request.Method, request.Path, request.Key)));

        Assert.False(JsonSerializer.Deserialize<CollectionSharingRequest>(server.Requests[0].Body, ReviewApiContract.WireSettings)!.Shared);
    }

    // A page not under this key any more (404) or gone (410) is forgotten by the window, whatever the action.
    [Theory]
    [InlineData(HttpStatusCode.NotFound)]
    [InlineData(HttpStatusCode.Gone)]
    public async Task A_page_gone_is_said_so_whatever_the_action(HttpStatusCode status)
    {
        var server = new Server { Answer = (_, _) => Answer(status, new CollectionRefusalAnswer(["That collection was deleted on 2027-November-26."])) };
        CollectionClient client = server.Client();

        Assert.Equal(CollectionClient.Outcome.GoneFromServer, (await client.SetSharingAsync("k7mq3x9vbn2p", false, Key)).Kind);
        Assert.Equal(CollectionClient.Outcome.GoneFromServer, (await client.RemoveImageAsync("k7mq3x9vbn2p", ShaA, Key)).Kind);
    }
}
