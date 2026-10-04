using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Handlers.Usage
{
    // ###########################################################################################
    // ADMIN > "API USAGE" (owner request, 2026-10-04: "how about tracking the API end-points, to see if
    // it is possible to retire any, if almost no versions uses it any more ... ideally we can remove
    // dead end-points along the way and then make sure to show proper 'please update' where
    // required"). GET /api/admin/api-usage?days=90 - administrator only.
    //
    // It ANSWERS the question and decides nothing: which versions called each route in the last N
    // days, and how many installations launched each version (the check-ins). Retiring a route is
    // the project owner's decision per route, never automatic, and never for the routes every CRT
    // ever released sends (NeverRetired - owner decision, 2026-10-04). How a route is retired - kept
    // in place, answering "please update" - is CLAUDE.md's "Installed CRTs keep working".
    // ###########################################################################################
    public static class ApiUsageFlow
    {
        public static async Task<ApiUsageAnswer?> ReadAsync(
            ReviewAccess? access,
            int? days,
            IEnumerable<(string Method, string Route)> mapped,
            IApiUsageStore store,
            ICheckInStore checkIns,
            DateTimeOffset nowUtc,
            ILogger logger,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(mapped);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(checkIns);
            ArgumentNullException.ThrowIfNull(logger);

            if (!ReviewAuthority.CanAdminister(access))
                return null;

            int shown = ApiUsageRules.DaysToShow(days);

            // Today counts as one of the days: 1 is today alone.
            DateOnly firstDay = DateOnly.FromDateTime(nowUtc.UtcDateTime).AddDays(1 - shown);

            IReadOnlyList<ApiUsageRow> rows = await store.ReadSinceAsync(firstDay, cancellationToken);

            // The launches are a second opinion; a table that cannot be read leaves them out, never
            // the routes.
            IReadOnlyList<CheckInLaunchRow> launches;

            try
            {
                launches = await checkIns.CountLaunchesAsync(firstDay, cancellationToken);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                logger.LogWarning(ex, "The launch check-ins could not be counted for the API usage screen.");
                launches = [];
            }

            return ApiUsageRules.Build(shown, mapped, rows, launches);
        }
    }

    // ###########################################################################################
    // Every route the server maps, read when asked - so Account > "API usage" lists the routes nobody
    // called too, which are the ones most worth seeing. Handed the route table by Program once every
    // route is mapped.
    // ###########################################################################################
    public sealed class ApiRouteList
    {
        private IEnumerable<EndpointDataSource> thisSources = [];

        public void Use(IEnumerable<EndpointDataSource> sources)
        {
            ArgumentNullException.ThrowIfNull(sources);
            this.thisSources = sources;
        }

        public IReadOnlyList<(string Method, string Route)> All() =>
            this.thisSources
                .SelectMany(source => source.Endpoints)
                .Select(ApiUsageRules.RouteOf)
                .OfType<(string Method, string Route)>()
                .Distinct()
                .ToList();
    }
}
