using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // THE WIRE BETWEEN CRT AND THE SERVER FOR BOARD VIEWS - CLAUDE.md's "a shared change needs a test
    // that would fail if only one side moved".
    //
    // CRT sends BoardViewContract.ToJson(report) (CRT.App's BoardViewReporter calls nothing else) to
    // "<api>/" + PathUnderApi; the server binds BoardViewReport with its configured JSON settings at
    // "/api/" + PathUnderApi. These tests put the client's exact body through the server's settings,
    // and read the route off the server's REAL route table (built, never started - rule 6), so a
    // renamed field, a changed naming policy or a moved route fails here rather than in the field,
    // where CRT would silently lose every view.
    // ###########################################################################################
    public sealed class BoardViewWireTests
    {
        private static readonly BoardViewReport Sent = new(
            "0f8fad5bd9cb469fa16570867728950e",
            "CRT 2026.10.0",
            "Windows",
            "Microsoft Windows 10.0.19045",
            "64-bit",
            [
                new BoardView("Commodore/C64/250407", new DateTimeOffset(2026, 9, 27, 11, 58, 3, TimeSpan.Zero), false),
                new BoardView("Commodore/VIC-20/250403", new DateTimeOffset(2026, 9, 27, 11, 59, 44, TimeSpan.Zero), true)
            ]);

        // The server's own settings: ASP.NET's web defaults with Program's wire settings on top.
        private static JsonSerializerOptions ServerSettings()
        {
            var options = new JsonSerializerOptions(JsonSerializerDefaults.Web);
            ReviewApiContract.ApplyWireSettings(options);
            return options;
        }

        [Fact]
        public void The_body_crt_sends_is_read_back_whole_by_the_server()
        {
            BoardViewReport? read = JsonSerializer.Deserialize<BoardViewReport>(
                BoardViewContract.ToJson(BoardViewWireTests.Sent), BoardViewWireTests.ServerSettings());

            Assert.NotNull(read);
            Assert.Equal(
                (BoardViewWireTests.Sent.BatchId, BoardViewWireTests.Sent.Version, BoardViewWireTests.Sent.OsHighlevel, BoardViewWireTests.Sent.OsVersion, BoardViewWireTests.Sent.Cpu),
                (read!.BatchId, read.Version, read.OsHighlevel, read.OsVersion, read.Cpu));
            Assert.Equal(BoardViewWireTests.Sent.Views!, read.Views!);
        }

        // The names on the wire, pinned - anything else reading a captured body (a log, a proxy, a
        // later client) sees these.
        [Fact]
        public void The_body_uses_the_agreed_field_names()
        {
            using JsonDocument body = JsonDocument.Parse(BoardViewContract.ToJson(BoardViewWireTests.Sent));
            JsonElement root = body.RootElement;

            Assert.Equal(
                ["batchId", "version", "osHighlevel", "osVersion", "cpu", "views"],
                root.EnumerateObject().Select(property => property.Name));

            Assert.Equal(
                ["boardId", "viewedUtc", "fromBeta"],
                root.GetProperty("views")[0].EnumerateObject().Select(property => property.Name));
        }

        // ###########################################################################################
        // *** A VIEW FROM A CRT BUILT BEFORE THE RENAME STILL COUNTS (owner decision, 2026-10-09). ***
        // Until "system" became "board", CRT named a view's board "systemId" - the 3.0.0 pre-releases
        // already installed still do, and the board-view route turns no CRT away. The server reads
        // the board from whichever name came (BoardView.IdOf).
        // ###########################################################################################
        [Fact]
        public void A_view_an_older_CRT_sends_as_systemId_is_read_with_its_board()
        {
            BoardViewReport? read = JsonSerializer.Deserialize<BoardViewReport>(
                """{"batchId":"0f8fad5bd9cb469fa16570867728950e","views":[{"systemId":"Commodore/C64/250407","viewedUtc":"2026-09-27T11:58:03+00:00","fromBeta":false}]}""",
                BoardViewWireTests.ServerSettings());

            BoardView view = Assert.Single(read!.Views!);
            Assert.Null(view.BoardId);
            Assert.Equal("Commodore/C64/250407", BoardView.IdOf(view));
        }

        // A new CRT never sends the old name.
        [Fact]
        public void The_old_systemId_name_is_never_sent()
        {
            Assert.DoesNotContain("systemId", BoardViewContract.ToJson(BoardViewWireTests.Sent), StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The fullest report CRT can send - the most views, each naming the longest id a column
        // holds, with the longest machine facts - fits the route's body limit. Over it, the server
        // answers 413 and CRT drops the report: views lost with nothing failing loudly.
        // ###########################################################################################
        [Fact]
        public void The_fullest_report_fits_the_routes_body_limit()
        {
            var view = new BoardView(new string('B', BoardViewRules.BoardIdLength), DateTimeOffset.UtcNow, true);

            var fullest = new BoardViewReport(
                Guid.NewGuid().ToString("N"),
                new string('v', BoardViewRules.VersionLength),
                new string('o', BoardViewRules.OsHighlevelLength),
                new string('s', BoardViewRules.OsVersionLength),
                new string('c', BoardViewRules.CpuLength),
                Enumerable.Repeat(view, BoardViewRules.MaxViewsPerReport).ToList());

            int bytes = System.Text.Encoding.UTF8.GetByteCount(BoardViewContract.ToJson(fullest));

            Assert.True(bytes < RequestBodyLimits.DefaultBytes, $"The fullest report is {bytes} bytes.");
        }

        // ###########################################################################################
        // The route CRT posts to exists on the server, takes the body, and is rate limited by the
        // policy the limiter registers.
        // ###########################################################################################
        [Fact]
        public void The_route_crt_posts_to_exists_and_is_rate_limited()
        {
            WebApplicationBuilder builder = WebApplication.CreateBuilder();
            Program.AddServerServices(builder.Services, new ServerOptions());

            using WebApplication app = builder.Build();
            Program.MapServerEndpoints(app);

            RouteEndpoint route = Assert.Single(
                ((IEndpointRouteBuilder)app).DataSources.SelectMany(source => source.Endpoints).OfType<RouteEndpoint>(),
                endpoint => endpoint.RoutePattern.RawText == "/api/" + BoardViewContract.PathUnderApi);

            Assert.Equal("POST", Assert.Single(route.Metadata.GetMetadata<IHttpMethodMetadata>()!.HttpMethods));
            Assert.Equal(BoardViewEndpoints.RateLimitPolicy, route.Metadata.GetMetadata<EnableRateLimitingAttribute>()?.PolicyName);
            Assert.NotNull(route.Metadata.GetMetadata<IAcceptsMetadata>());
        }
    }
}
