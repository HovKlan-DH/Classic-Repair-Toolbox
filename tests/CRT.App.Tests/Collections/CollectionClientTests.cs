using System.Net;
using System.Text;
using System.Text.Json;
using ClassicRepairToolbox.Tests.Maintainer;
using Handlers.Collections;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Collections;

// ###########################################################################################
// Sharing a collection with the server (CollectionClient), on a fake server answering in CollectionContract's records as the
// server writes them (ReviewApiContract.WireSettings) - so a field renamed on one side fails here.
// ###########################################################################################
public sealed class CollectionClientTests
{
    private static readonly DateTimeOffset Now = new(2026, 11, 26, 12, 0, 0, TimeSpan.Zero);

    private const string Root = "https://classic-repair-toolbox.dk/crt-test";

    private static readonly string ShaA = new('a', 64);

    private static readonly string ShaB = new('b', 64);

    private static LocalCollection Collection(string? serverId = null) =>
        new()
        {
            LocalId = new string('0', 32),
            HardwareName = "Commodore 64",
            BoardName = "250407",
            Title = "No picture",
            CreatedUtc = Now,
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
        public List<(HttpMethod Method, string Path, string? Bearer, byte[] Body, string? ContentType)> Requests { get; } = [];

        public Func<HttpRequestMessage, int, HttpResponseMessage> Answer { get; set; } = (_, _) => new HttpResponseMessage(HttpStatusCode.InternalServerError);

        public CollectionClient Client() => new(Root, new HttpClient(new AnsweringHttpHandler(async (request, token) =>
        {
            byte[] body = request.Content is null ? [] : await request.Content.ReadAsByteArrayAsync(token);

            this.Requests.Add((request.Method, request.RequestUri!.AbsolutePath, request.Headers.Authorization?.Parameter, body, request.Content?.Headers.ContentType?.MediaType));

            return this.Answer(request, this.Requests.Count);
        })));
    }

    private static Func<string, byte[]?> Captures(Dictionary<string, byte[]> bytes) => sha => bytes.TryGetValue(sha, out byte[]? png) ? png : null;

    [Fact]
    public async Task A_first_share_sends_the_whole_collection_signed_in_and_keeps_what_it_sent()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Created, Shared()) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Shared, result.Kind);
        Assert.Equal("k7mq3x9vbn2p", result.Answer!.Id);

        (HttpMethod method, string path, string? bearer, byte[] body, _) = Assert.Single(server.Requests);

        Assert.Equal((HttpMethod.Post, "/crt-test/api/collections", "token-1"), (method, path, bearer));

        // What the server binds - the very record, by the wire's names.
        CollectionContentRequest sent = JsonSerializer.Deserialize<CollectionContentRequest>(body, ReviewApiContract.WireSettings)!;

        Assert.Equal(["U7", "U8"], sent.Images.Select(image => image.ComponentLabel));
        Assert.Equal(Now.AddYears(1), sent.DeleteAtUtc);
        Assert.Equal(CollectionSharing.ContentHash(sent), result.SharedContentHash);

        // The date SENT is what the collection keeps - never the server's echo, whose database keeps milliseconds only.
        Assert.Equal(sent.DeleteAtUtc, result.SentDeleteAtUtc);
    }

    [Fact]
    public async Task The_captures_the_server_asks_for_are_sent_and_then_the_collection_again()
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

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), "token-1", Captures(new() { [ShaB] = [7, 7, 7] }), Now);

        Assert.Equal(CollectionClient.Outcome.Shared, result.Kind);
        Assert.Equal(3, server.Requests.Count);

        var upload = server.Requests[1];

        Assert.Equal((HttpMethod.Put, $"/crt-test/api/collections/images/{ShaB}", "token-1", "image/png"), (upload.Method, upload.Path, upload.Bearer, upload.ContentType));
        Assert.Equal([7, 7, 7], upload.Body);
        Assert.Equal(HttpMethod.Post, server.Requests[2].Method);
    }

    [Fact]
    public async Task A_capture_no_longer_on_this_computer_is_said_and_nothing_more_is_sent()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Conflict, new CollectionRefusalAnswer(["missing"], [ShaA])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Contains("no longer on this computer", result.Message, StringComparison.Ordinal);
        Assert.Single(server.Requests);
    }

    [Fact]
    public async Task A_collection_shared_before_is_sent_to_its_page_and_one_the_server_deleted_is_said()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.Gone, new CollectionRefusalAnswer(["That collection was deleted on 2027-November-26."])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection("k7mq3x9vbn2p"), "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.GoneFromServer, result.Kind);
        Assert.Equal("That collection was deleted on 2027-November-26.", result.Message);
        Assert.Equal((HttpMethod.Put, "/crt-test/api/collections/k7mq3x9vbn2p"), (server.Requests[0].Method, server.Requests[0].Path));
    }

    [Fact]
    public async Task A_session_the_server_no_longer_knows_asks_to_sign_in()
    {
        var server = new Server { Answer = (_, _) => new HttpResponseMessage(HttpStatusCode.Unauthorized) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), "old-token", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.SignInNeeded, result.Kind);
        Assert.Equal(CollectionClient.SignInFirst, result.Message);
    }

    [Fact]
    public async Task A_refusal_is_said_in_the_server_s_words()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.TooManyRequests, new CollectionRefusalAnswer(["An account shares at most 20 new collections a day."])) };

        CollectionClient.Result result = await server.Client().ShareAsync(Collection(), "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Equal("An account shares at most 20 new collections a day.", result.Message);
    }

    [Fact]
    public async Task A_server_out_of_reach_is_said_in_words()
    {
        var client = new CollectionClient(Root, new HttpClient(new AnsweringHttpHandler((_, _) => throw new HttpRequestException("No such host"))));

        CollectionClient.Result result = await client.ShareAsync(Collection(), "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Failed, result.Kind);
        Assert.Equal(CollectionClient.NotReachable, result.Message);
    }

    // The rules both ends keep are asked before anything is sent.
    [Fact]
    public async Task A_collection_breaking_a_rule_is_refused_without_a_request()
    {
        var server = new Server();

        CollectionClient.Result result = await server.Client().ShareAsync(Collection() with { Captures = [] }, "token-1", _ => null, Now);

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Contains("A collection holds at least one capture.", result.Errors);
        Assert.Empty(server.Requests);
    }

    // ---- Comments (step 3) ---------------------------------------------------------------------

    private static CollectionCommentAnswer Comment(long id, bool mine = false, bool hidden = false) =>
        new(id, ShaA, mine ? "Anna" : "Bo", ByOwner: mine, ByAdministrator: false, Mine: mine, "Check R7.", Now, null, hidden);

    [Fact]
    public async Task The_owner_reads_the_comments_as_the_server_writes_them()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.OK, new CollectionCommentsAnswer([Comment(5), Comment(6, mine: true, hidden: true)])) };

        CollectionClient.CommentsResult result = await server.Client().ListCommentsAsync("k7mq3x9vbn2p", "token-1");

        Assert.Equal(CollectionClient.Outcome.Done, result.Kind);
        Assert.Equal([(5L, "Bo", false), (6L, "Anna", true)], result.Comments.Select(comment => (comment.Id, comment.AuthorName, comment.Hidden)));
        Assert.Equal((HttpMethod.Get, "/crt-test/api/collections/k7mq3x9vbn2p/comments", "token-1"), (server.Requests[0].Method, server.Requests[0].Path, server.Requests[0].Bearer));
    }

    [Fact]
    public async Task An_answer_hiding_and_deleting_each_go_to_their_route_with_their_record()
    {
        var server = new Server { Answer = (_, _) => new HttpResponseMessage(HttpStatusCode.NoContent) };
        CollectionClient client = server.Client();

        Assert.Equal(CollectionClient.Outcome.Done, (await client.AddCommentAsync("k7mq3x9vbn2p", ShaA, "Thanks - R7 it was.", "token-1")).Kind);
        Assert.Equal(CollectionClient.Outcome.Done, (await client.SetCommentHiddenAsync("k7mq3x9vbn2p", 5, true, "token-1")).Kind);
        Assert.Equal(CollectionClient.Outcome.Done, (await client.DeleteCommentAsync("k7mq3x9vbn2p", 6, "token-1")).Kind);

        Assert.Equal(
            [
                (HttpMethod.Post, "/crt-test/api/collections/k7mq3x9vbn2p/comments"),
                (HttpMethod.Put, "/crt-test/api/collections/k7mq3x9vbn2p/comments/5/hidden"),
                (HttpMethod.Delete, "/crt-test/api/collections/k7mq3x9vbn2p/comments/6")
            ],
            server.Requests.Select(request => (request.Method, request.Path)));

        CollectionCommentRequest answer = JsonSerializer.Deserialize<CollectionCommentRequest>(server.Requests[0].Body, ReviewApiContract.WireSettings)!;

        Assert.Equal((ShaA, "Thanks - R7 it was."), (answer.ImageSha256, answer.Body));
        Assert.True(JsonSerializer.Deserialize<CollectionCommentHideRequest>(server.Requests[1].Body, ReviewApiContract.WireSettings)!.Hidden);
    }

    [Fact]
    public async Task A_refused_comment_is_said_in_the_server_s_words()
    {
        var server = new Server { Answer = (_, _) => Answer(HttpStatusCode.BadRequest, new CollectionRefusalAnswer(["Write the comment first."])) };

        CollectionClient.Result result = await server.Client().AddCommentAsync("k7mq3x9vbn2p", ShaA, " ", "token-1");

        Assert.Equal(CollectionClient.Outcome.Refused, result.Kind);
        Assert.Equal("Write the comment first.", result.Message);
    }

    [Fact]
    public async Task The_owner_s_collections_say_their_newest_comment()
    {
        var server = new Server
        {
            Answer = (_, _) => Answer(HttpStatusCode.OK, new MyCollectionsAnswer(
                [new CollectionSummaryAnswer("k7mq3x9vbn2p", "https://x/collections/k7mq3x9vbn2p", "No picture", "Commodore 64", "250407", null, 3, Now, Now, Now.AddYears(1), 2, 17)]))
        };

        IReadOnlyList<CollectionSummaryAnswer>? mine = await server.Client().ListMineAsync("token-1");

        Assert.Equal((2, 17L), (mine![0].CommentCount, mine[0].LatestCommentId));
        Assert.Equal("/crt-test/api/collections/mine", server.Requests[0].Path);

        var refusing = new Server { Answer = (_, _) => new HttpResponseMessage(HttpStatusCode.Unauthorized) };
        Assert.Null(await refusing.Client().ListMineAsync("old-token"));
    }

    [Theory]
    [InlineData(HttpStatusCode.NoContent, CollectionClient.Outcome.Deleted)]
    [InlineData(HttpStatusCode.NotFound, CollectionClient.Outcome.Deleted)]
    [InlineData(HttpStatusCode.Unauthorized, CollectionClient.Outcome.SignInNeeded)]
    public async Task Deleting_a_page_is_done_also_when_the_server_has_none(HttpStatusCode status, CollectionClient.Outcome expected)
    {
        var server = new Server { Answer = (_, _) => new HttpResponseMessage(status) };

        CollectionClient.Result result = await server.Client().DeleteAsync("k7mq3x9vbn2p", "token-1");

        Assert.Equal(expected, result.Kind);
        Assert.Equal((HttpMethod.Delete, "/crt-test/api/collections/k7mq3x9vbn2p", "token-1"), (server.Requests[0].Method, server.Requests[0].Path, server.Requests[0].Bearer));
    }
}
