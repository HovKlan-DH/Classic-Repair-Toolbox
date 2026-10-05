using System.Text;
using System.Text.Json;
using CRT.Server.Configuration;
using CRT.Server.Handlers;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers RequestBodyLimits - how large a body each route may send (security review,
    // 2026-09-25). Every route used to accept Kestrel's default of about 30 MB, including the
    // anonymous submission endpoint, which stores its body in the database.
    //
    // *** ASSERTED ON THE SERVER'S REAL ROUTE TABLE (code review, 2026-09-25). *** Each limit is
    // written on its route where the route is mapped, so these tests build that table - through
    // Program's own AddServerServices and MapServerEndpoints - and read the limits off it. The
    // limits used to be a second, hand-kept copy of the routes matched by URL segment, which a new
    // route or a renamed segment silently fell through to 64 KB.
    //
    // Rule 6 holds: the app is BUILT and never started, so nothing listens and no request is
    // served. Building it constructs no store either - registering a service creates nothing until
    // something resolves it, and nothing here does.
    // ###########################################################################################
    public sealed class RequestBodyLimitsTests
    {
        private static readonly Lazy<IReadOnlyList<RouteEndpoint>> Routes = new(RequestBodyLimitsTests.BuildRoutes);

        private static IReadOnlyList<RouteEndpoint> BuildRoutes()
        {
            WebApplicationBuilder builder = WebApplication.CreateBuilder();
            Program.AddServerServices(builder.Services, new ServerOptions());

            using WebApplication app = builder.Build();
            Program.MapServerEndpoints(app);

            return ((IEndpointRouteBuilder)app).DataSources
                .SelectMany(source => source.Endpoints)
                .OfType<RouteEndpoint>()
                .ToList();
        }

        private static string MethodOf(RouteEndpoint endpoint) =>
            endpoint.Metadata.GetMetadata<IHttpMethodMetadata>()?.HttpMethods.SingleOrDefault() ?? "*";

        private static string PatternOf(RouteEndpoint endpoint) =>
            (endpoint.RoutePattern.RawText ?? string.Empty).TrimEnd('/');

        private static string NameOf(RouteEndpoint endpoint) =>
            $"{RequestBodyLimitsTests.MethodOf(endpoint)} {RequestBodyLimitsTests.PatternOf(endpoint)}";

        private static RouteEndpoint Route(string method, string pattern) =>
            Assert.Single(
                RequestBodyLimitsTests.Routes.Value,
                endpoint => RequestBodyLimitsTests.NameOf(endpoint) == $"{method} {pattern.TrimEnd('/')}");

        // ###########################################################################################
        // Every route that reads a body, and the limit decided for it. A route that takes more than
        // the default says so where it is mapped; the rest are listed at the default ON PURPOSE, so
        // that a route added later is a failure here until somebody decides.
        // ###########################################################################################
        public static TheoryData<string, string, long> DecidedBodyRoutes() => new()
        {
            { "POST", "/api/submissions", RequestBodyLimits.ManifestBytes },
            { "PUT", "/api/submissions/{submissionId:long}/blobs/{hash}", RequestBodyLimits.BlobChunkBytes },
            { "POST", "/api/review/submissions/{submissionId:long}/amend", RequestBodyLimits.ManifestBytes },

            // A system's table edit (2026-10-03): a whole board's rows, as an amendment carries - and
            // its check, which carries the same rows before the reason is asked for.
            { "POST", "/api/review/systems/edit", RequestBodyLimits.ManifestBytes },
            { "POST", "/api/review/systems/edit/check", RequestBodyLimits.ManifestBytes },
            { "POST", "/api/review/submissions/{submissionId:long}/approve", RequestBodyLimits.PathListBytes },
            { "POST", "/api/review/production/publish", RequestBodyLimits.PathListBytes },
            { "POST", "/api/admin/unused-files/remove", RequestBodyLimits.PathListBytes },
            { "POST", "/api/admin/systems/order", RequestBodyLimits.PathListBytes },

            // Small bodies, deliberately at the default: a sign-in, a token, an address and a
            // password, a maintainer change, a plan request, a review comment of at most 4,000
            // characters.
            { "POST", "/api/accounts/register", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/login", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/refresh", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/logout", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/forgot-password", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/reset-password", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/accept-invitation", RequestBodyLimits.DefaultBytes },

            // A maintainer's own account (2026-10-03): a name, an address and a password, a code,
            // or two passwords.
            { "POST", "/api/accounts/me/name", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/me/email", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/me/email/confirm", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/accounts/me/password", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/admin/maintainers", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/admin/maintainers/remove", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/admin/maintainers/invite", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/admin/maintainers/invitations/withdraw", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/production/plan", RequestBodyLimits.DefaultBytes },

            // A rollback: a system id, and the reason the contributor is told. Both small - the
            // comment is the same shape as a rejection's, which is at the default too.
            { "POST", "/api/review/production/rollback/plan", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/production/rollback", RequestBodyLimits.DefaultBytes },

            // Deleting a system (2026-10-03): a system id, a fingerprint and the reason the
            // contributors of its open submissions are told - a rollback's shape.
            { "POST", "/api/admin/systems/delete/plan", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/admin/systems/delete", RequestBodyLimits.DefaultBytes },

            // Resetting the contribution data (2026-10-04): the fingerprint of the counts shown.
            { "POST", "/api/admin/reset", RequestBodyLimits.DefaultBytes },

            // One system's facts for the "Systems" screen: a system id and nothing else.
            { "POST", "/api/review/systems/detail", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/systems/listing", RequestBodyLimits.DefaultBytes },

            // A system's Board data and Files views (2026-10-03): a system id and nothing else.
            { "POST", "/api/review/systems/table", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/systems/files", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/submissions/{submissionId:long}/reject", RequestBodyLimits.DefaultBytes },
            { "POST", "/api/review/submissions/{submissionId:long}/request-changes", RequestBodyLimits.DefaultBytes },

            // Board views from CRT (2026-09-27): at most BoardViewRules.MaxViewsPerReport short views.
            { "POST", "/api/usage/board-views", RequestBodyLimits.DefaultBytes },

            // CRT's launch check-in (2026-10-03): four short form fields - at the path CRT posts to
            // (AppConfig.CheckInUrl, pinned on CRT's side by LaunchCheckInTests) and Apache forwards
            // the old check-in address to.
            { "POST", "/api/usage/check-in", RequestBodyLimits.DefaultBytes },

            // Feedback from CRT (2026-10-03): the text and a zip of the attached files, up to the
            // limit CRT checks before sending (FeedbackContract) - at the path CRT posts to.
            { "POST", "/api/feedback", RequestBodyLimits.FeedbackBytes },

            // CRT 2.x's contribution upload (2026-10-04): a zip read to its end only so the "please
            // update" answer reaches the sender - see LegacyContributionEndpoints.
            { "POST", "/api/legacy/contribution", RequestBodyLimits.LegacyContributionBytes },
        };

        [Theory]
        [MemberData(nameof(DecidedBodyRoutes))]
        public void Each_route_that_reads_a_body_carries_the_limit_decided_for_it(string method, string pattern, long expected)
        {
            Assert.Equal(expected, RequestBodyLimits.For(RequestBodyLimitsTests.Route(method, pattern)));
        }

        // ###########################################################################################
        // The route CRT's "Rebuild both manifests" calls is really mapped, under the path the client
        // builds (ReviewApiRoutes.AdminRebuildManifests, pinned on CRT's side by ReviewApiRoutesTests).
        // It reads no body, so the guard below never sees it - and a renamed segment on either side
        // compiles and answers 404 (code review, 2026-10-01). Read off this file's real route table.
        // ###########################################################################################
        [Fact]
        public void The_manifest_rebuild_route_is_mapped_where_the_Maintainer_tab_calls_it()
        {
            RequestBodyLimitsTests.Route("POST", "/api/admin/manifest/rebuild");
        }

        // ###########################################################################################
        // *** THE GUARD THIS FILE EXISTS FOR. *** A route that reads a JSON body (routing marks it
        // with IAcceptsMetadata) or carries a limit of its own must be listed above. A new one fails
        // here with its name, instead of shipping at 64 KB and answering 413 to the first real board.
        // ###########################################################################################
        [Fact]
        public void No_route_reads_a_body_without_a_decided_limit()
        {
            var decided = new HashSet<string>(
                RequestBodyLimitsTests.DecidedBodyRoutes().Select(row => $"{row.Data.Item1} {row.Data.Item2.TrimEnd('/')}"),
                StringComparer.Ordinal);

            List<string> undecided = RequestBodyLimitsTests.Routes.Value
                .Where(endpoint =>
                    endpoint.Metadata.GetMetadata<IAcceptsMetadata>() is not null ||
                    endpoint.Metadata.GetMetadata<BodyLimit>() is not null)
                .Select(RequestBodyLimitsTests.NameOf)
                .Where(name => !decided.Contains(name))
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();

            Assert.True(
                undecided.Count == 0,
                "These routes read a request body but have no decided limit - give each .WithBodyLimit(...) " +
                "where it is mapped if it needs more than 64 KB, and list it in DecidedBodyRoutes:\n" +
                string.Join("\n", undecided));
        }

        // ###########################################################################################
        // The contributor's "I discarded my draft" notice (2026-09-28) is mapped at the path CRT
        // builds from CRT.Data's DraftDiscardContract - pinned literally here and there - and reads no
        // body, so it needs no decided limit.
        // ###########################################################################################
        [Fact]
        public void The_draft_discarded_notice_is_mapped_where_CRT_sends_it_and_reads_no_body()
        {
            RouteEndpoint route = RequestBodyLimitsTests.Route("POST", "/api/submissions/{submissionId:long}/draft-discarded");

            Assert.Null(route.Metadata.GetMetadata<IAcceptsMetadata>());
        }

        // DENY BY DEFAULT: an endpoint without a limit of its own, and a request matching no route at
        // all, get the small limit.
        [Fact]
        public void A_route_without_a_limit_of_its_own_and_no_route_at_all_get_the_small_default()
        {
            Assert.Equal(RequestBodyLimits.DefaultBytes, RequestBodyLimits.For(RequestBodyLimitsTests.Route("POST", "/api/accounts/login")));
            Assert.Equal(RequestBodyLimits.DefaultBytes, RequestBodyLimits.For(RequestBodyLimitsTests.Route("GET", "/api/review/queue")));
            Assert.Equal(RequestBodyLimits.DefaultBytes, RequestBodyLimits.For(null));
        }

        // The limit is ALSO ASP.NET Core's own request-size metadata, so the framework's routing
        // middleware - which applies that where it finds it - can never set a different number.
        [Fact]
        public void A_routes_limit_is_the_frameworks_request_size_metadata_too()
        {
            RouteEndpoint amend = RequestBodyLimitsTests.Route("POST", "/api/review/submissions/{submissionId:long}/amend");

            Assert.Equal(RequestBodyLimits.ManifestBytes, amend.Metadata.GetMetadata<IRequestSizeLimitMetadata>()?.MaxRequestBodySize);
        }

        [Fact]
        public void The_manifest_gets_room_for_the_largest_real_board_several_times_over()
        {
            // The largest shipped board's manifest measured about 1.2 MB of JSON.
            Assert.True(RequestBodyLimits.ManifestBytes >= 4 * 1_300_000);

            // And well under the default it replaces.
            Assert.True(RequestBodyLimits.ManifestBytes < 30_000_000);
        }

        // ###########################################################################################
        // *** THE CHUNK LIMIT MUST FIT WHAT CRT SENDS. *** A client chunk larger than the server's
        // cap would compile, ship, and have every upload refused with 413. Both sides now read the
        // one constant in the shared contract; this asserts the server's cap covers it.
        // ###########################################################################################
        [Fact]
        public void An_upload_chunk_gets_at_least_what_the_client_sends()
        {
            Assert.True(RequestBodyLimits.BlobChunkBytes >= SubmissionFormat.UploadChunkBytes);
        }

        // ###########################################################################################
        // *** SAVING THE MAINTAINER'S TABLE SENDS A WHOLE BOARD (code review, 2026-09-25). *** The
        // amend route fell to the 64 KB default, so every save of a real board was refused with 413
        // before the endpoint ran. Measured on the largest shipped board, serialised exactly as the
        // Maintainer tab sends it.
        // ###########################################################################################
        [Fact]
        public async Task Saving_the_maintainers_table_for_the_largest_shipped_board_fits_its_limit()
        {
            string workbook = RequestBodyLimitsTests.LargestShippedWorkbook();
            string cacheKey = "body-limits:" + Guid.NewGuid().ToString("N");
            BoardData? board;

            try
            {
                board = await BoardDataReader.LoadAsync(workbook, cacheKey);
            }
            finally
            {
                BoardDataReader.ClearCache(cacheKey);
            }

            Assert.NotNull(board);

            string body = JsonSerializer.Serialize(
                new AmendRequest(1, SubmissionRowsBoard.FromBoard(board!)), ReviewApiContract.WireSettings);
            long bytes = Encoding.UTF8.GetByteCount(body);

            // Why the default could not do: this is the failure the route used to have.
            Assert.True(bytes > RequestBodyLimits.DefaultBytes, $"The largest board's table is only {bytes} bytes.");

            // And room to grow several times over.
            long limit = RequestBodyLimits.For(RequestBodyLimitsTests.Route("POST", "/api/review/submissions/{submissionId:long}/amend"));
            Assert.True(bytes * 4 < limit);
        }

        // Approving, publishing to production and removing unused files send back a list of paths
        // the maintainer was shown. Even the whole shipped tree listed that way fits with room to spare.
        [Theory]
        [InlineData("/api/review/submissions/{submissionId:long}/approve")]
        [InlineData("/api/review/production/publish")]
        [InlineData("/api/admin/unused-files/remove")]
        public void A_list_of_every_file_in_the_tree_fits_the_routes_that_send_one(string pattern)
        {
            string dataRoot = RequestBodyLimitsTests.DataRoot();

            List<string> everyFile = Directory
                .EnumerateFiles(dataRoot, "*", SearchOption.AllDirectories)
                .Select(file => Path.GetRelativePath(dataRoot, file).Replace(Path.DirectorySeparatorChar, '/'))
                .ToList();

            Assert.True(everyFile.Count > 1000, "The shipped tree was not found.");

            string body = JsonSerializer.Serialize(new ReviewDecisionRequest(null, everyFile), ReviewApiContract.WireSettings);
            long limit = RequestBodyLimits.For(RequestBodyLimitsTests.Route("POST", pattern));

            Assert.Equal(RequestBodyLimits.PathListBytes, limit);
            Assert.True(Encoding.UTF8.GetByteCount(body) * 2 < limit);
        }

        private static string DataRoot()
        {
            string? folder = AppContext.BaseDirectory;

            while (folder is not null && !File.Exists(Path.Combine(folder, "Classic-Repair-Toolbox.slnx")))
                folder = Path.GetDirectoryName(folder);

            Assert.NotNull(folder);

            return Path.Combine(folder!, "Assets", "Data");
        }

        // The biggest board workbook in the shipped tree by bytes - the board with the most rows.
        private static string LargestShippedWorkbook() =>
            Directory
                .EnumerateFiles(RequestBodyLimitsTests.DataRoot(), "Data *" + DataGenerationRules.WorkbookExtension, SearchOption.AllDirectories)
                .OrderByDescending(file => new FileInfo(file).Length)
                .First();
    }
}
