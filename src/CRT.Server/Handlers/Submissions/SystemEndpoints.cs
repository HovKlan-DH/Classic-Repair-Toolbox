using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE "SYSTEMS" SCREEN, over HTTP (owner request, 2026-09-27). A rim over SystemOverviewFlow
    // and nothing else.
    //
    //   GET  /api/review/systems                     - every system (SystemOverviewAnswer).
    //   POST /api/review/systems/detail  {systemId}  - one system (SystemDetailAnswer). A POST
    //                                                  because a system id carries slashes.
    //
    // Read-only, and for ANY maintainer (the flow's header says why). THE STATUS CODES: 401 no
    // usable credentials, 403 an account that may not review anything, 404 no such system.
    // ###########################################################################################
    public static class SystemEndpoints
    {
        public static void MapSystemEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder systems = app.MapGroup("/api/review/systems");

            systems.MapGet("/", SystemEndpoints.ListAsync);
            systems.MapPost("/detail", SystemEndpoints.DetailAsync);

            // Where a new system goes in the drop-down lists (owner request, 2026-09-27) - see
            // SystemListingFlow. A placement is two names, notes and a path: the default limit.
            systems.MapGet("/listing", SystemEndpoints.ListingAsync);
            systems.MapPost("/listing", SystemEndpoints.SetPlacementAsync);
        }

        private static async Task<IResult> ListingAsync(
            HttpContext context,
            IAccountStore accounts,
            SystemListingFlow listing,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SystemListingOutcome outcome = await listing.ListAsync(
                access,
                options.DataTreeRoot,
                PublishedSystemLister.List(options.DataTreeRoot),
                cancellationToken);

            return outcome.IsForbidden
                ? ReviewEndpoints.NotAMaintainer()
                : Results.Ok(outcome.Answer);
        }

        private static async Task<IResult> SetPlacementAsync(
            SetPlacementRequest request,
            HttpContext context,
            IAccountStore accounts,
            SystemListingFlow listing,
            ServerOptions options,
            ILogger<SystemListingFlow> logger,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SetPlacementOutcome outcome = await listing.SetAsync(
                access, request, options.DataTreeRoot, DateTimeOffset.UtcNow, cancellationToken);

            if (outcome.Answer is SetPlacementAnswer answer)
            {
                // Written into BETA's main Excel data file: the sync manifest must say so, or no
                // client downloads the new list - the same rebuild a publish ends with, and like
                // there it never fails the request (the file is already written).
                if (answer.ListedInBeta)
                {
                    int written = DataChecksumManifest.Write(
                        options.DataTreeRoot ?? string.Empty,
                        options.PublicDataBaseUrl ?? string.Empty,
                        options.ManifestPath ?? string.Empty);

                    if (written < 0)
                    {
                        logger.LogWarning(
                            "A system was added to BETA's drop-down lists but the checksum manifest at [{ManifestPath}] " +
                            "could not be regenerated - clients will not see it until it is rebuilt.",
                            options.ManifestPath);
                    }
                }

                return Results.Ok(answer);
            }

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Error });

            return outcome.IsConflict
                ? Results.Conflict(new { error = outcome.Error })
                : Results.BadRequest(new { error = outcome.Error });
        }

        private static async Task<IResult> ListAsync(
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            IBoardViewStore boardViews,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SystemOverviewOutcome outcome = await SystemOverviewFlow.ListAsync(
                access,
                PublishedSystemLister.List(options.DataTreeRoot),
                SystemEndpoints.ProductionSystems(options),
                submissions,
                accounts,
                cancellationToken,
                boardViews);

            if (outcome.IsForbidden)
                return ReviewEndpoints.NotAMaintainer();

            return Results.Ok(new SystemOverviewAnswer(outcome.Systems!));
        }

        private static async Task<IResult> DetailAsync(
            SystemDetailRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            IBoardViewStore boardViews,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SystemOverviewOutcome outcome = await SystemOverviewFlow.DetailAsync(
                access,
                request?.SystemId,
                PublishedSystemLister.List(options.DataTreeRoot),
                SystemEndpoints.ProductionSystems(options),
                submissions,
                accounts,
                cancellationToken,
                boardViews: boardViews);

            if (outcome.IsForbidden)
                return ReviewEndpoints.NotAMaintainer();

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = "No such system." });

            return Results.Ok(outcome.Detail);
        }

        // ###########################################################################################
        // The boards production holds, or null when this server has no production tree to look in -
        // so the screen never says "not in production" about a tree nobody asked. The promotion's
        // own root first; the older ProductionTreeRoot (named so BETA can be kept away from it)
        // otherwise.
        // ###########################################################################################
        private static IReadOnlyList<PublishedSystemLister.KnownSystem>? ProductionSystems(ServerOptions options)
        {
            string? root = !string.IsNullOrWhiteSpace(options.ProductionDataTreeRoot)
                ? options.ProductionDataTreeRoot
                : options.ProductionTreeRoot;

            return string.IsNullOrWhiteSpace(root) || !Directory.Exists(root)
                ? null
                : PublishedSystemLister.List(root);
        }
    }
}
