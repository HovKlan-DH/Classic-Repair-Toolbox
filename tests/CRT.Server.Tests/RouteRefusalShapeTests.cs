using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.DependencyInjection;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // The code review of 2026-10-13 on the routes themselves:
    //
    //   - every contributor route (they need no account) is limited per address, generously enough
    //     for a real upload of a large board;
    //   - a 404 carries words - CRT reads a refusal's `error` (ApiRefusal.Read), and a bodiless 404
    //     showed nothing at all.
    //
    // The route table is the server's real one (ServerRouteTable); results are EXECUTED against a
    // bare context, as ReviewEndpointsRefusalTests does - no HTTP pipeline is started.
    // ###########################################################################################
    public sealed class RouteRefusalShapeTests
    {
        [Fact]
        public void Every_contributor_route_is_limited_per_address()
        {
            List<RouteEndpoint> routes = ServerRouteTable.Routes
                .Where(route => ServerRouteTable.PatternOf(route).StartsWith("/api/submissions", StringComparison.Ordinal))
                .ToList();

            Assert.NotEmpty(routes);

            foreach (RouteEndpoint route in routes)
            {
                string? policy = route.Metadata.GetMetadata<EnableRateLimitingAttribute>()?.PolicyName;

                Assert.True(
                    policy is SubmissionEndpoints.RateLimitPolicy or SubmissionEndpoints.StatusRateLimitPolicy,
                    $"{ServerRouteTable.NameOf(route)} is not limited per address.");
            }
        }

        // ###########################################################################################
        // *** ASKING HAS AN ALLOWANCE OF ITS OWN (code review, 2026-10-14). *** Every CRT behind one
        // address asks about its open submissions each minute; in the uploads' bucket that steady
        // asking refused status checks and failed uploads. The two questions CRT asks repeatedly go
        // to the status allowance, and everything that SENDS stays in the other.
        // ###########################################################################################
        [Fact]
        public void A_submissions_status_and_discard_notice_are_counted_apart_from_sending_it()
        {
            List<(string Pattern, string Method, string? Policy)> routes = ServerRouteTable.Routes
                .Where(route => ServerRouteTable.PatternOf(route).StartsWith("/api/submissions", StringComparison.Ordinal))
                .Select(route => (
                    ServerRouteTable.PatternOf(route),
                    ServerRouteTable.MethodOf(route),
                    route.Metadata.GetMetadata<EnableRateLimitingAttribute>()?.PolicyName))
                .ToList();

            List<(string Pattern, string Method, string? Policy)> asking = routes
                .Where(route => route.Policy == SubmissionEndpoints.StatusRateLimitPolicy)
                .ToList();

            Assert.Equal(2, asking.Count);
            Assert.Contains(asking, route => route.Method == "GET" && route.Pattern.EndsWith("{submissionId:long}", StringComparison.Ordinal));
            Assert.Contains(asking, route => route.Method == "POST" && route.Pattern.EndsWith(DraftDiscardContract.RouteSegment, StringComparison.Ordinal));

            Assert.Equal(4, routes.Count(route => route.Policy == SubmissionEndpoints.RateLimitPolicy));
        }

        // A 1 GB board in upload-sized parts, a status question for each - all inside one hour's
        // sending allowance.
        [Fact]
        public void The_contributor_allowance_fits_a_large_upload_with_room_to_spare()
        {
            long parts = (1024L * 1024 * 1024) / SubmissionFormat.UploadChunkBytes;
            long needed = (parts * 2) + 10;

            Assert.True(SubmissionEndpoints.RateLimitPerHour >= needed, $"{SubmissionEndpoints.RateLimitPerHour} < {needed}");
        }

        // Thirty CRTs behind one shared address, each asking about ten open submissions every minute
        // for an hour, fit the asking allowance.
        [Fact]
        public void The_asking_allowance_fits_a_shared_address_full_of_CRTs()
        {
            const int crtsBehindOneAddress = 30;
            const int openSubmissionsEach = 10;

            int needed = crtsBehindOneAddress * openSubmissionsEach * 60;

            Assert.True(SubmissionEndpoints.StatusRateLimitPerHour >= needed, $"{SubmissionEndpoints.StatusRateLimitPerHour} < {needed}");
        }

        [Fact]
        public async Task A_review_routes_404_says_what_is_missing()
        {
            string body = await RouteRefusalShapeTests.ExecuteAsync(ReviewEndpoints.NoSuchSubmission(), StatusCodes.Status404NotFound);

            Assert.Equal("No such submission.", ApiRefusal.Read(body).Message);
        }

        // The contributor routes' 404 has words too - the SAME words for "no such submission" and
        // "not yours", so the two stay indistinguishable.
        [Fact]
        public async Task A_contributor_routes_404_has_words_and_says_nothing_about_whose_it_is()
        {
            string body = await RouteRefusalShapeTests.ExecuteAsync(SubmissionEndpoints.NoSuchSubmission(), StatusCodes.Status404NotFound);

            Assert.Equal("No such submission.", ApiRefusal.Read(body).Message);
        }

        private static async Task<string> ExecuteAsync(IResult result, int expectedStatus)
        {
            var context = new DefaultHttpContext
            {
                RequestServices = new ServiceCollection().AddLogging().BuildServiceProvider(),
            };

            context.Response.Body = new MemoryStream();

            await result.ExecuteAsync(context);

            Assert.Equal(expectedStatus, context.Response.StatusCode);

            context.Response.Body.Position = 0;
            return await new StreamReader(context.Response.Body).ReadToEndAsync();
        }
    }
}
