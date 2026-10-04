using CRT.Server.Handlers.Compat;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // The pure half of counting which CRT versions call which route (owner request, 2026-10-04: "how
    // about tracking the API end-points, to see if it is possible to retire any, if almost no versions
    // uses it any more"). What a request counts as, and how the counts become Account > "API usage".
    //
    // *** A ROUTE IS ITS PATTERN, never the path asked for. *** "/api/review/submissions/
    // {submissionId:long}" is one row however many submissions it is called for, so no id lands in
    // the table, and a request matching no route (a scanner, a typo) is not counted at all - its path
    // is the sender's to write and would grow the table without end.
    //
    // *** A VERSION IS WHAT THE USER-AGENT NAMES *** (CrtVersion.FromUserAgent, the version policy's own
    // reading), as CrtVersion writes it - "CRT 3.0.0" and "CRT/3.0.0" are one version. A request
    // naming none is ApiUsageVersion.NotCrt; the text of a User-Agent is never stored.
    // ###########################################################################################
    public static class ApiUsageRules
    {
        // The days shown when none are asked for, and the most that may be.
        public const int DefaultDays = 90;

        public const int MaximumDays = 366;

        // A version longer than this is somebody's made-up User-Agent, not a CRT: counted as
        // ApiUsageVersion.Other. CRT's own are 5 to 20 characters ("3.0.0-alpha.2").
        public const int MaximumVersionLength = 40;

        // crt_api_calls.route's width; every route the server maps is far shorter.
        public const int MaximumRouteLength = 200;

        // Counts one request - from Program's pipeline, after routing chose its endpoint.
        public static void Count(ApiUsageCounter counter, Endpoint? endpoint, string? userAgent, DateTimeOffset nowUtc)
        {
            ArgumentNullException.ThrowIfNull(counter);

            if (ApiUsageRules.RouteOf(endpoint) is (string method, string route))
                counter.Count(method, route, ApiUsageRules.VersionOf(userAgent), nowUtc);
        }

        // ###########################################################################################
        // The method and pattern of a route, as RequestBodyLimitsTests names them ("POST",
        // "/api/submissions" - no trailing slash), or null for anything that is not a route.
        // ###########################################################################################
        public static (string Method, string Route)? RouteOf(Endpoint? endpoint)
        {
            if (endpoint is not RouteEndpoint route)
                return null;

            string pattern = (route.RoutePattern.RawText ?? string.Empty).TrimEnd('/');

            if (pattern.Length == 0 || pattern.Length > ApiUsageRules.MaximumRouteLength)
                return null;

            if (!pattern.StartsWith('/'))
                pattern = "/" + pattern;

            string method = route.Metadata.GetMetadata<IHttpMethodMetadata>()?.HttpMethods.FirstOrDefault() ?? "*";

            return (method.ToUpperInvariant(), pattern);
        }

        public static string VersionOf(string? userAgent)
        {
            if (CrtVersion.FromUserAgent(userAgent) is not CrtVersion version)
                return ApiUsageVersion.NotCrt;

            string text = version.ToString();

            return text.Length <= ApiUsageRules.MaximumVersionLength ? text : ApiUsageVersion.Other;
        }

        public static int DaysToShow(int? asked) =>
            asked is int days ? Math.Clamp(days, 1, ApiUsageRules.MaximumDays) : ApiUsageRules.DefaultDays;

        // ###########################################################################################
        // The answer: EVERY route - the ones mapped now, called or not, and any the table still holds
        // that no longer is - each with its versions newest first, then the routes by pattern. And the
        // launches per version, newest first.
        //
        // Area and NeverRetired are ClientVersionPolicy's: the Forever area is what every CRT ever
        // released sends (the check-in, feedback, board views, the health check, the 2.x contribution
        // address), which is never retired (owner decision, 2026-10-04).
        //
        // Launches by the same version written two ways ("CRT 3.0.0", "CRT/3.0.0") are added up, so
        // an installation that used both counts twice - CRT has only ever sent the first.
        // ###########################################################################################
        public static ApiUsageAnswer Build(
            int days,
            IEnumerable<(string Method, string Route)> mapped,
            IReadOnlyList<ApiUsageRow> rows,
            IReadOnlyList<CheckInLaunchRow> launches)
        {
            ArgumentNullException.ThrowIfNull(mapped);
            ArgumentNullException.ThrowIfNull(rows);
            ArgumentNullException.ThrowIfNull(launches);

            var routes = new SortedSet<(string Method, string Route)>(
                mapped.Concat(rows.Select(row => (row.Method, row.Route))),
                Comparer<(string Method, string Route)>.Create((left, right) =>
                {
                    int byRoute = string.CompareOrdinal(left.Route, right.Route);
                    return byRoute != 0 ? byRoute : string.CompareOrdinal(left.Method, right.Method);
                }));

            var answer = new List<ApiUsageRoute>();

            foreach ((string method, string route) in routes)
            {
                List<ApiUsageVersion> versions = rows
                    .Where(row => row.Method == method && row.Route == route)
                    .GroupBy(row => row.Version, StringComparer.Ordinal)
                    .Select(group => new ApiUsageVersion(group.Key, group.Sum(row => row.Calls), group.Max(row => row.LastUtc)))
                    .OrderBy(version => version.Version, ApiUsageRules.NewestFirst)
                    .ToList();

                ClientArea area = ClientVersionPolicy.AreaOf(route);

                answer.Add(new ApiUsageRoute(
                    method,
                    route,
                    area.ToString(),
                    area == ClientArea.Forever,
                    versions.Sum(version => version.Calls),
                    versions.Count == 0 ? null : versions.Max(version => version.LastUtc),
                    versions));
            }

            List<ApiUsageInstallations> installations = launches
                .GroupBy(launch => ApiUsageRules.VersionOf(launch.VersionText) is string version && version != ApiUsageVersion.NotCrt
                    ? version
                    : launch.VersionText.Trim(), StringComparer.Ordinal)
                .Select(group => new ApiUsageInstallations(group.Key, group.Sum(launch => launch.Installations), group.Sum(launch => launch.Launches)))
                .OrderBy(installation => installation.Version, ApiUsageRules.NewestFirst)
                .ToList();

            return new ApiUsageAnswer(days, answer, installations);
        }

        // Versions newest first; anything that is not a version ("(not CRT)", "(other)") after them.
        public static IComparer<string> NewestFirst { get; } = Comparer<string>.Create((left, right) =>
        {
            bool leftIs = CrtVersion.TryParse(left, out CrtVersion? leftVersion);
            bool rightIs = CrtVersion.TryParse(right, out CrtVersion? rightVersion);

            if (leftIs && rightIs)
                return rightVersion!.CompareTo(leftVersion);

            if (leftIs != rightIs)
                return leftIs ? -1 : 1;

            return string.CompareOrdinal(left, right);
        });
    }

    // One row of crt_api_calls added up over the days asked for.
    public sealed record ApiUsageRow(string Method, string Route, string Version, long Calls, DateTimeOffset LastUtc);

    // ###########################################################################################
    // Launches of one version text in the launch check-ins (crt_update): how many installations -
    // distinct addresses - and how many launches. VersionText is the column as stored, the User-Agent
    // ("CRT 3.0.0").
    // ###########################################################################################
    public sealed record CheckInLaunchRow(string VersionText, int Installations, int Launches);
}
