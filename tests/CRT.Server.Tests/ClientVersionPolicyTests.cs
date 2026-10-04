using System.Text;
using CRT.Server.Handlers.Compat;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ClientVersionPolicy and ClientVersionGate - which installed CRTs each part of the API
    // serves, and the "update CRT" answer the rest get (owner request, 2026-10-04: "It is important
    // that all older versions will continue to work, including the checkin and feedback").
    //
    // The gate is driven with a DefaultHttpContext: nothing listens, no request is served (rule 6).
    // Its answer is read back with CRT.Data's ApiRefusal - the code CRT itself reads it with - so a
    // renamed field on either side fails here.
    // ###########################################################################################
    public sealed class ClientVersionPolicyTests
    {
        private static readonly ClientVersionPolicy Strict = new(
            submissions: CrtVersion.Parse("3.2.0"),
            maintainer: CrtVersion.Parse("3.4.0"));

        // -----------------------------------------------------------------------------------
        // The minimums in force
        // -----------------------------------------------------------------------------------

        // *** NO MINIMUM VERSION IS SET. *** Raising one turns installed CRTs away, so it must be a
        // decision written down (VERSION.md's MAJOR row, CLAUDE.md) - and changing this expectation
        // is the moment that decision is made, not a test to be updated in passing.
        //
        // What IS in force is the API revision (2026-10-04): this server's own, so a CRT built from
        // the same source - and any request naming no revision - is served.
        [Fact]
        public void No_minimum_version_is_set_and_the_servers_own_api_revision_is_served()
        {
            string current = ClientVersionContract.ApiRevision.ToString(System.Globalization.CultureInfo.InvariantCulture);

            Assert.Null(ClientVersionPolicy.Current.Submissions);
            Assert.Null(ClientVersionPolicy.Current.Maintainer);
            Assert.Equal(ClientVersionContract.ApiRevision, ClientVersionPolicy.Current.ApiRevision);
            Assert.Null(ClientVersionPolicy.Current.Refuse("/api/submissions", "CRT 0.0.1"));
            Assert.Null(ClientVersionPolicy.Current.Refuse("/api/review/queue", "CRT 0.0.1"));
            Assert.Null(ClientVersionPolicy.Current.Refuse("/api/submissions", "CRT 3.0.0-alpha.2", current));
            Assert.Null(ClientVersionPolicy.Current.Refuse("/api/review/queue", "CRT 3.0.0-alpha.2", current));
        }

        // -----------------------------------------------------------------------------------
        // Areas
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("/api/submissions", ClientArea.Submissions)]
        [InlineData("/api/submissions/42/finalise", ClientArea.Submissions)]
        [InlineData("/API/Submissions/42", ClientArea.Submissions)]
        [InlineData("/api/review/queue", ClientArea.Maintainer)]
        [InlineData("/api/review/systems/edit", ClientArea.Maintainer)]
        [InlineData("/api/admin/manifest/rebuild", ClientArea.Maintainer)]
        [InlineData("/api/accounts/login", ClientArea.Maintainer)]
        [InlineData("/api/accounts/me/password", ClientArea.Maintainer)]
        [InlineData("/api/health", ClientArea.Forever)]
        [InlineData("/api/feedback", ClientArea.Forever)]
        [InlineData("/api/usage/check-in", ClientArea.Forever)]
        [InlineData("/api/usage/board-views", ClientArea.Forever)]
        [InlineData("/api/legacy/contribution", ClientArea.Forever)]
        [InlineData("/api/reviewer", ClientArea.Forever)]
        [InlineData("/api/submissionsX", ClientArea.Forever)]
        [InlineData("", ClientArea.Forever)]
        [InlineData(null, ClientArea.Forever)]
        public void Each_path_belongs_to_one_area(string? path, ClientArea expected)
        {
            Assert.Equal(expected, ClientVersionPolicy.AreaOf(path));
        }

        // ###########################################################################################
        // *** EVERY ROUTE THE SERVER MAPS, CLASSIFIED. *** The never-refused set is listed here on
        // purpose: a new top-level route falls into it by default, and this fails until somebody
        // decides whether old CRTs may ever be turned away from it.
        // ###########################################################################################
        [Fact]
        public void Only_the_routes_every_CRT_ever_released_uses_are_never_refused()
        {
            string[] forever = ServerRouteTable.Routes
                .Where(route => ClientVersionPolicy.AreaOf(ServerRouteTable.PatternOf(route)) == ClientArea.Forever)
                .Select(ServerRouteTable.NameOf)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToArray();

            Assert.Equal(
                [
                    "GET /api/health",
                    "POST /api/feedback",
                    "POST /api/legacy/contribution",
                    "POST /api/usage/board-views",
                    "POST /api/usage/check-in"
                ],
                forever);
        }

        // -----------------------------------------------------------------------------------
        // Refuse
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData("/api/submissions", "CRT 3.1.9", "3.2.0")]
        [InlineData("/api/submissions/7/blobs/abc", "CRT 3.2.0-beta.4", "3.2.0")]
        [InlineData("/api/review/queue", "CRT 3.3.0", "3.4.0")]
        [InlineData("/api/accounts/login", "CRT 2.5.0", "3.4.0")]
        public void A_CRT_older_than_its_areas_minimum_is_told_to_update(string path, string userAgent, string minimum)
        {
            ClientOutdatedAnswer? answer = ClientVersionPolicyTests.Strict.Refuse(path, userAgent);

            Assert.NotNull(answer);
            Assert.Equal(ClientVersionContract.OutdatedCode, answer!.Code);
            Assert.Equal(minimum, answer.MinimumVersion);
            Assert.Contains($"[{CrtVersion.FromUserAgent(userAgent)}]", answer.Message);
        }

        [Theory]
        [InlineData("/api/submissions", "CRT 3.2.0")]
        [InlineData("/api/submissions", "CRT 4.0.0-alpha.1")]
        [InlineData("/api/review/queue", "CRT 3.4.0")]
        public void A_CRT_at_or_above_the_minimum_is_served(string path, string userAgent)
        {
            Assert.Null(ClientVersionPolicyTests.Strict.Refuse(path, userAgent));
        }

        // The promise: the check-in, feedback, board views and the 2.x contribution address answer
        // every CRT, however old, whatever the minimums are set to.
        [Theory]
        [InlineData("/api/usage/check-in")]
        [InlineData("/api/feedback")]
        [InlineData("/api/usage/board-views")]
        [InlineData("/api/legacy/contribution")]
        [InlineData("/api/health")]
        public void The_routes_every_CRT_uses_are_never_refused(string path)
        {
            var everythingRaised = new ClientVersionPolicy(CrtVersion.Parse("99.0.0"), CrtVersion.Parse("99.0.0"));

            Assert.Null(everythingRaised.Refuse(path, "CRT 1.0.0"));
        }

        // FAIL OPEN, the old PHP page's rule: a request naming no CRT version is never refused - a
        // browser opening a mailed link to /api/accounts/verify must still work.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("Mozilla/5.0 (Windows NT 10.0; Win64; x64)")]
        [InlineData("CRT dev-build")]
        public void A_request_naming_no_CRT_version_is_never_refused(string? userAgent)
        {
            Assert.Null(ClientVersionPolicyTests.Strict.Refuse("/api/accounts/verify", userAgent));
            Assert.Null(ClientVersionPolicyTests.Strict.Refuse("/api/submissions", userAgent));
        }

        // -----------------------------------------------------------------------------------
        // The API revision
        // -----------------------------------------------------------------------------------

        // A server at revision 3, as after two breaking changes.
        private static readonly ClientVersionPolicy AtRevision3 = new(submissions: null, maintainer: null, apiRevision: 3);

        // ###########################################################################################
        // *** A CRT BUILT FOR AN OLDER API REVISION IS TOLD TO UPDATE (2026-10-04). *** Nobody picks a
        // version: the published alpha carries the revision of the source it was built from, and the
        // server knows its own. The answer names no minimum version - the server knows none that has
        // its revision - only "the newest".
        // ###########################################################################################
        [Theory]
        [InlineData("/api/submissions", "2")]
        [InlineData("/api/submissions/7/finalise", "1")]
        [InlineData("/api/review/queue", "2")]
        [InlineData("/api/admin/systems", "2")]
        [InlineData("/api/accounts/login", "2")]
        public void A_CRT_built_for_an_older_api_revision_is_told_to_update(string path, string revision)
        {
            ClientOutdatedAnswer? answer = ClientVersionPolicyTests.AtRevision3.Refuse(path, "CRT 3.0.0-alpha.2", revision);

            Assert.NotNull(answer);
            Assert.Equal(ClientVersionContract.OutdatedCode, answer!.Code);
            Assert.Null(answer.MinimumVersion);
            Assert.Equal(ClientVersionContract.OutdatedApi(CrtVersion.Parse("3.0.0-alpha.2")).Message, answer.Message);
        }

        // The same revision is served, and so is a NEWER one: the server is deployed before the CRT
        // that needs it, and a newer CRT against an older server is the one being tried out.
        [Theory]
        [InlineData("3")]
        [InlineData("4")]
        public void A_CRT_of_the_servers_revision_or_newer_is_served(string revision)
        {
            Assert.Null(ClientVersionPolicyTests.AtRevision3.Refuse("/api/submissions", "CRT 3.0.0-alpha.2", revision));
            Assert.Null(ClientVersionPolicyTests.AtRevision3.Refuse("/api/review/queue", "CRT 3.0.0-alpha.2", revision));
        }

        // FAIL OPEN, as for a request naming no version: no revision, or one that does not read, is
        // not refused for it - a CRT 2.x, a browser, a stray proxy.
        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("zero")]
        [InlineData("0")]
        public void A_request_naming_no_readable_api_revision_is_not_refused_for_it(string? revision)
        {
            Assert.Null(ClientVersionPolicyTests.AtRevision3.Refuse("/api/submissions", "CRT 3.0.0-alpha.2", revision));
        }

        // The check-in, feedback, board views, health and the 2.x contribution address answer an old
        // revision too - the promise holds whatever the revision is.
        [Theory]
        [InlineData("/api/usage/check-in")]
        [InlineData("/api/feedback")]
        [InlineData("/api/usage/board-views")]
        [InlineData("/api/legacy/contribution")]
        [InlineData("/api/health")]
        public void The_routes_every_CRT_uses_are_never_refused_for_the_api_revision(string path)
        {
            Assert.Null(ClientVersionPolicyTests.AtRevision3.Refuse(path, "CRT 3.0.0-alpha.2", "1"));
        }

        // Both levers at once: an older revision is refused for the revision even when the version
        // would pass, and a revision that passes still meets the area's minimum version.
        [Fact]
        public void The_api_revision_and_a_minimum_version_each_refuse_on_their_own()
        {
            var both = new ClientVersionPolicy(submissions: CrtVersion.Parse("3.2.0"), maintainer: null, apiRevision: 3);

            Assert.Null(both.Refuse("/api/submissions", "CRT 3.4.0", "2")!.MinimumVersion);
            Assert.Equal("3.2.0", both.Refuse("/api/submissions", "CRT 3.1.0", "3")!.MinimumVersion);
            Assert.Null(both.Refuse("/api/submissions", "CRT 3.2.0", "3"));
        }

        // -----------------------------------------------------------------------------------
        // The gate in the pipeline
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_refused_request_gets_426_with_the_answer_CRT_reads_and_never_reaches_the_endpoint()
        {
            DefaultHttpContext context = ClientVersionPolicyTests.Request("/api/submissions", "CRT 3.0.0", "the manifest");
            bool reached = false;

            await ClientVersionGate.HandleAsync(
                context, _ => { reached = true; return Task.CompletedTask; }, ClientVersionPolicyTests.Strict, NullLogger.Instance);

            Assert.False(reached);
            Assert.Equal(426, context.Response.StatusCode);

            ApiRefusal refusal = ApiRefusal.Read(ClientVersionPolicyTests.ResponseText(context));

            Assert.True(refusal.IsClientOutdated);
            Assert.Equal("3.2.0", refusal.MinimumVersion);
            Assert.Contains("[3.0.0]", refusal.Message);
            Assert.Contains("[3.2.0]", refusal.Message);
        }

        // The upload is read to its end before the answer - Apache turns an answer that arrives
        // while the client is still sending into a 502, and the words would never arrive.
        [Fact]
        public async Task A_refused_request_has_its_body_read_to_the_end_before_the_answer()
        {
            DefaultHttpContext context = ClientVersionPolicyTests.Request("/api/submissions/7/blobs/abc", "CRT 3.0.0", new string('x', 100_000));

            await ClientVersionGate.HandleAsync(context, _ => Task.CompletedTask, ClientVersionPolicyTests.Strict, NullLogger.Instance);

            Assert.Equal(context.Request.Body.Length, context.Request.Body.Position);
        }

        // The gate reads the revision from its header - CRT.Data's name for it, which CRT sends.
        [Fact]
        public async Task A_request_naming_an_older_api_revision_in_its_header_gets_426()
        {
            DefaultHttpContext context = ClientVersionPolicyTests.Request("/api/review/queue", "CRT 3.0.0-alpha.2", string.Empty);
            context.Request.Headers[ClientVersionContract.ApiRevisionHeader] = "2";
            bool reached = false;

            await ClientVersionGate.HandleAsync(
                context, _ => { reached = true; return Task.CompletedTask; }, ClientVersionPolicyTests.AtRevision3, NullLogger.Instance);

            Assert.False(reached);
            Assert.Equal(426, context.Response.StatusCode);

            ApiRefusal refusal = ApiRefusal.Read(ClientVersionPolicyTests.ResponseText(context));

            Assert.True(refusal.IsClientOutdated);
            Assert.Null(refusal.MinimumVersion);
            Assert.Contains("[3.0.0-alpha.2]", refusal.Message);
        }

        [Fact]
        public async Task A_served_request_goes_on_to_the_endpoint_untouched()
        {
            DefaultHttpContext context = ClientVersionPolicyTests.Request("/api/submissions", "CRT 3.2.0", "the manifest");
            bool reached = false;

            await ClientVersionGate.HandleAsync(
                context, _ => { reached = true; return Task.CompletedTask; }, ClientVersionPolicyTests.Strict, NullLogger.Instance);

            Assert.True(reached);
            Assert.Equal(200, context.Response.StatusCode);
            Assert.Equal(0, context.Request.Body.Position);
        }

        private static DefaultHttpContext Request(string path, string userAgent, string body)
        {
            var context = new DefaultHttpContext();
            context.Request.Method = "POST";
            context.Request.Path = path;
            context.Request.Headers.UserAgent = userAgent;
            context.Request.Body = new MemoryStream(Encoding.UTF8.GetBytes(body));
            context.Response.Body = new MemoryStream();
            return context;
        }

        private static string ResponseText(DefaultHttpContext context)
        {
            context.Response.Body.Position = 0;
            return new StreamReader(context.Response.Body).ReadToEnd();
        }
    }
}
