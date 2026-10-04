using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Routing;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ApiUsageRules - what a request counts as, and how the counts become Account > "API usage"
    // (owner request, 2026-10-04: "how about tracking the API end-points, to see if it is possible to
    // retire any").
    //
    // The route tests read the server's REAL route table (ServerRouteTable), so a route mapped in a
    // way the counting cannot name fails here rather than silently going uncounted.
    // ###########################################################################################
    public sealed class ApiUsageRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 4, 12, 0, 0, TimeSpan.Zero);

        // Every route counts as its method and pattern - never a real path - in the form the body
        // limits name routes, and short enough for crt_api_calls.route.
        [Fact]
        public void Every_mapped_route_counts_as_its_method_and_pattern()
        {
            Assert.NotEmpty(ServerRouteTable.Routes);

            foreach (RouteEndpoint endpoint in ServerRouteTable.Routes)
            {
                (string Method, string Route)? counted = ApiUsageRules.RouteOf(endpoint);

                Assert.NotNull(counted);
                Assert.Equal(ServerRouteTable.MethodOf(endpoint), counted!.Value.Method);
                Assert.Equal(ServerRouteTable.PatternOf(endpoint), counted.Value.Route);
                Assert.StartsWith("/api/", counted.Value.Route, StringComparison.Ordinal);
                Assert.True(counted.Value.Route.Length <= ApiUsageRules.MaximumRouteLength);
            }
        }

        // ###########################################################################################
        // A route with an id in it is ONE row however many submissions it is called for - its
        // pattern is what is counted, so no id ever reaches the table.
        // ###########################################################################################
        [Fact]
        public void A_call_for_one_submission_counts_as_the_route_not_the_submission()
        {
            RouteEndpoint status = Assert.Single(
                ServerRouteTable.Routes,
                endpoint => ServerRouteTable.NameOf(endpoint) == "GET /api/submissions/{submissionId:long}");

            var counter = new ApiUsageCounter();
            ApiUsageRules.Count(counter, status, "CRT 3.0.0", ApiUsageRulesTests.Now);
            ApiUsageRules.Count(counter, status, "CRT 3.0.0", ApiUsageRulesTests.Now);

            ApiUsageTally tally = Assert.Single(counter.TakeAll());
            Assert.Equal(("GET", "/api/submissions/{submissionId:long}", "3.0.0", 2L), (tally.Method, tally.Route, tally.Version, tally.Calls));
        }

        // A request matching no route - a scanner, a typo - is not counted: its path is the sender's.
        [Fact]
        public void A_request_matching_no_route_is_not_counted()
        {
            var counter = new ApiUsageCounter();

            ApiUsageRules.Count(counter, endpoint: null, "CRT 3.0.0", ApiUsageRulesTests.Now);
            ApiUsageRules.Count(counter, new Endpoint(_ => Task.CompletedTask, EndpointMetadataCollection.Empty, "not a route"), "CRT 3.0.0", ApiUsageRulesTests.Now);

            Assert.Empty(counter.TakeAll());
        }

        [Theory]
        [InlineData("CRT 3.0.0-beta.2", "3.0.0-beta.2")]
        [InlineData("CRT/3.0.0", "3.0.0")]
        [InlineData("crt 2.5.0", "2.5.0")]
        [InlineData("CRT 3.0.0+abc123", "3.0.0")]
        [InlineData("Mozilla/5.0 (Windows NT 10.0)", ApiUsageVersion.NotCrt)]
        [InlineData("curl/8.4.0", ApiUsageVersion.NotCrt)]
        [InlineData("", ApiUsageVersion.NotCrt)]
        [InlineData(null, ApiUsageVersion.NotCrt)]
        [InlineData("CRT not-a-version", ApiUsageVersion.NotCrt)]
        public void A_requests_version_is_the_CRT_version_its_User_Agent_names(string? userAgent, string expected)
        {
            Assert.Equal(expected, ApiUsageRules.VersionOf(userAgent));
        }

        // A made-up version too long to be CRT's never reaches the table as written.
        [Fact]
        public void A_version_too_long_to_be_CRTs_counts_as_other()
        {
            string made = "CRT 3.0.0-" + string.Join(".", Enumerable.Repeat("abcdefgh", 10));

            Assert.Equal(ApiUsageVersion.Other, ApiUsageRules.VersionOf(made));
        }

        [Theory]
        [InlineData(null, 90)]
        [InlineData(30, 30)]
        [InlineData(0, 1)]
        [InlineData(-5, 1)]
        [InlineData(5000, ApiUsageRules.MaximumDays)]
        public void The_days_shown_are_ninety_unless_asked_and_never_beyond_a_year(int? asked, int expected)
        {
            Assert.Equal(expected, ApiUsageRules.DaysToShow(asked));
        }

        // ###########################################################################################
        // *** A ROUTE NOBODY CALLED IS LISTED TOO *** - it is the answer most worth seeing - and a
        // route the table still holds that is no longer mapped is listed as well.
        // ###########################################################################################
        [Fact]
        public void Every_route_is_listed_called_or_not_with_its_versions_newest_first()
        {
            ApiUsageAnswer answer = ApiUsageRules.Build(
                90,
                [("GET", "/api/review/queue"), ("POST", "/api/review/systems/edit")],
                [
                    new ApiUsageRow("GET", "/api/review/queue", "3.0.0", 40, ApiUsageRulesTests.Now.AddDays(-3)),
                    new ApiUsageRow("GET", "/api/review/queue", ApiUsageVersion.NotCrt, 2, ApiUsageRulesTests.Now.AddDays(-9)),
                    new ApiUsageRow("GET", "/api/review/queue", "3.0.0-beta.7", 5, ApiUsageRulesTests.Now.AddDays(-30)),
                    new ApiUsageRow("GET", "/api/review/queue", "3.1.0", 7, ApiUsageRulesTests.Now),
                    new ApiUsageRow("POST", "/api/review/old-route", "3.0.0", 1, ApiUsageRulesTests.Now.AddDays(-60)),
                ],
                []);

            Assert.Equal(90, answer.Days);
            Assert.Equal(
                ["POST /api/review/old-route", "GET /api/review/queue", "POST /api/review/systems/edit"],
                answer.Routes.Select(route => $"{route.Method} {route.Route}"));

            ApiUsageRoute queue = answer.Routes.Single(route => route.Route == "/api/review/queue");
            Assert.Equal(["3.1.0", "3.0.0", "3.0.0-beta.7", ApiUsageVersion.NotCrt], queue.Versions.Select(version => version.Version));
            Assert.Equal(54, queue.Calls);
            Assert.Equal(ApiUsageRulesTests.Now, queue.LastUtc);

            ApiUsageRoute unused = answer.Routes.Single(route => route.Route == "/api/review/systems/edit");
            Assert.Equal(0, unused.Calls);
            Assert.Null(unused.LastUtc);
            Assert.Empty(unused.Versions);
        }

        // ###########################################################################################
        // *** NEVER RETIRED (owner decision, 2026-10-04). *** What every CRT ever released sends -
        // the check-in, feedback, board views, the health check, the 2.x contribution address - is
        // marked so on the real route table, and nothing in the Submit or Maintainer areas is.
        // ###########################################################################################
        [Fact]
        public void The_routes_every_CRT_ever_released_sends_are_marked_never_retired_and_no_others()
        {
            ApiUsageAnswer answer = ApiUsageRules.Build(
                90,
                ServerRouteTable.Routes.Select(ApiUsageRules.RouteOf).OfType<(string Method, string Route)>(),
                [],
                []);

            HashSet<string> never = answer.Routes
                .Where(route => route.NeverRetired)
                .Select(route => $"{route.Method} {route.Route}")
                .ToHashSet(StringComparer.Ordinal);

            Assert.Contains("POST /api/usage/check-in", never);
            Assert.Contains("POST /api/usage/board-views", never);
            Assert.Contains("POST /api/feedback", never);
            Assert.Contains("GET /api/health", never);
            Assert.Contains("POST /api/legacy/contribution", never);

            Assert.All(answer.Routes.Where(route => route.NeverRetired), route => Assert.Equal("Forever", route.Area));
            Assert.All(answer.Routes.Where(route => !route.NeverRetired), route => Assert.Contains(route.Area, new[] { "Submissions", "Maintainer" }));
            Assert.Contains(answer.Routes, route => route.Route == "/api/submissions" && route.Area == "Submissions");
            Assert.Contains(answer.Routes, route => route.Route == "/api/review/queue" && route.Area == "Maintainer");
        }

        // ###########################################################################################
        // The launches per version, from the check-ins' User-Agent text: one version written two ways
        // is one version, newest first, and a text that names none is kept as it is.
        // ###########################################################################################
        [Fact]
        public void Launches_are_per_version_newest_first_whichever_way_the_version_was_written()
        {
            ApiUsageAnswer answer = ApiUsageRules.Build(
                90,
                [],
                [],
                [
                    new CheckInLaunchRow("CRT 2.5.0", 120, 900),
                    new CheckInLaunchRow("CRT 3.0.0", 40, 300),
                    new CheckInLaunchRow("CRT/3.0.0", 2, 5),
                    new CheckInLaunchRow("Something odd", 1, 1),
                ]);

            Assert.Equal(["3.0.0", "2.5.0", "Something odd"], answer.Installations.Select(installation => installation.Version));
            Assert.Equal((42, 305), (answer.Installations[0].Installations, answer.Installations[0].Launches));
        }
    }
}
