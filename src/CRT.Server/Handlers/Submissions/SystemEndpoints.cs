using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE "SYSTEMS" SCREEN, over HTTP (owner request, 2026-09-27). A rim over the flows named
    // beside each route and nothing else.
    //
    //   GET  /api/review/systems                     - every system (SystemOverviewAnswer).
    //   POST /api/review/systems/detail  {systemId}  - one system (SystemDetailAnswer). A POST
    //                                                  because a system id carries slashes.
    //   POST /api/review/systems/table   {systemId}  - BETA's board (SystemEditFlow, 2026-10-03).
    //   POST /api/review/systems/edit/check          - what a table edit would remove from BETA.
    //   POST /api/review/systems/edit                - a table edit, PUBLISHED to BETA.
    //   POST /api/review/systems/files   {systemId}  - every file it uses (SystemFilesFlow).
    //
    // For ANY maintainer to read (the overview flow's header says why); an edit only for the
    // system's own maintainers and the administrator. THE STATUS CODES: 401 no usable credentials,
    // 403 an account that may not review anything (or not change this system), 404 no such system.
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

            // A system's Board data and Files views (owner request, 2026-10-03) - see SystemEditFlow
            // and SystemFilesFlow. An edit carries a whole board's rows, as an amendment does - and
            // so does its check.
            systems.MapPost("/table", SystemEndpoints.TableAsync);
            systems.MapPost("/edit/check", SystemEndpoints.CheckEditAsync)
                .WithBodyLimit(RequestBodyLimits.ManifestBytes);
            systems.MapPost("/edit", SystemEndpoints.EditAsync)
                .WithBodyLimit(RequestBodyLimits.ManifestBytes);
            systems.MapPost("/files", SystemEndpoints.FilesAsync);
        }

        // ###########################################################################################
        // POST /api/review/systems/table {systemId} - BETA's board as the table opens on it. 404 with
        // the reason when BETA holds nothing of the system.
        // ###########################################################################################
        private static async Task<IResult> TableAsync(
            SystemDetailRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            PublishedBoardReader boards,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            // The stable source's board, read-only (2026-10-04) - or BETA's, as before. The stable
            // root is the one the overview judged "in the stable source" by.
            SystemTableOutcome outcome = DataTreeNames.IsProduction(request?.Tree)
                ? await SystemEditFlow.ReadStableTableAsync(
                    access,
                    request?.SystemId,
                    options.StableSourceRoot,
                    boards,
                    options.ProductionPublicDataBaseUrl,
                    cancellationToken)
                : await SystemEditFlow.ReadTableAsync(
                    access,
                    request?.SystemId,
                    options.DataTreeRoot,
                    submissions,
                    boards,
                    options.PublicDataBaseUrl,
                    options.IsProductionPublishingConfigured,
                    cancellationToken);

            if (outcome.IsForbidden)
                return ReviewEndpoints.NotAMaintainer();

            return outcome.IsNotFound
                ? Results.NotFound(new { error = outcome.Error })
                : Results.Ok(outcome.Answer);
        }

        // ###########################################################################################
        // POST /api/review/systems/edit/check - what publishing the table edit would remove from
        // BETA, asked before the maintainer is asked for a reason; nothing is made. The status codes
        // are the edit's own, so a refusal here is the one the edit would meet.
        // ###########################################################################################
        private static async Task<IResult> CheckEditAsync(
            SystemEditRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            PublishedBoardReader boards,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SystemEditOutcome outcome = await SystemEditFlow.CheckAsync(
                access,
                request,
                options.DataTreeRoot,
                submissions,
                boards,
                options.IsProductionPublishingConfigured,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return outcome.IsAccepted
                ? Results.Ok(new SystemEditCheckAnswer(outcome.Removals))
                : SystemEndpoints.EditRefusal(outcome);
        }

        // ###########################################################################################
        // POST /api/review/systems/edit - the maintainer's table edit, PUBLISHED TO BETA as a
        // submission from their account that they approve at once (owner decision, 2026-10-03: "it
        // should go directly to the next queue, 'BETA > Stable', so it can directly be tested in
        // BETA"). The status codes match an amendment's: 403 not this account's system (or one waiting
        // in BETA), 404 not in BETA, 409 BETA changed since the table was opened or the files it
        // would remove did, 400 a refusal with its findings.
        //
        // *** PUBLISHED: BETA'S CHECKSUM MANIFEST IS REWRITTEN, as after any approval, and nobody is
        // mailed. *** The approval's own mail goes to the submission's sender - here the maintainer
        // who has just published it - and tells them nothing. Neither may fail the request: the board
        // is already in BETA (ReviewEndpoints.ApproveAsync explains why that matters).
        //
        // *** MADE BUT NOT PUBLISHED: whoever must approve is told, as for a contributor's upload -
        // except the sender. *** The submission waits under Contributor Submissions; a mail problem
        // never fails the request.
        // ###########################################################################################
        private static async Task<IResult> EditAsync(
            SystemEditRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            BlobStore blobs,
            PublishedBoardReader boards,
            ApprovePublishFlow approvals,
            SubmissionNotifier notifier,
            ServerOptions options,
            ILogger<SubmissionNotifier> logger,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            SystemEditOutcome outcome = await SystemEditFlow.PublishAsync(
                access,
                request,
                options.DataTreeRoot,
                submissions,
                blobs,
                boards,
                approvals,
                options.IsProductionPublishingConfigured,
                SubmissionEndpoints.ClientAddress(context),
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (!outcome.IsAccepted)
                return SystemEndpoints.EditRefusal(outcome);

            if (outcome.IsPublished)
            {
                try
                {
                    ReviewEndpoints.RewriteBetaManifest(outcome.SubmissionId, options, logger);
                }
                catch (Exception ex)
                {
                    logger.LogError(
                        ex,
                        "Submission {SubmissionId} from the Systems screen WAS published, but BETA's checksum manifest " +
                        "could not be rewritten after it. The data tree is correct.",
                        outcome.SubmissionId);
                }
            }
            else
            {
                try
                {
                    SubmissionRecord? record = await submissions.FindAsync(outcome.SubmissionId, cancellationToken);

                    if (record is not null)
                    {
                        IReadOnlyList<MailRecipient> recipients = (await SubmissionRouting.RecipientsForAsync(record, accounts, cancellationToken))
                            .Where(recipient => !string.Equals(recipient.Email, access!.Account.Email, StringComparison.OrdinalIgnoreCase))
                            .ToList();

                        if (recipients.Count > 0)
                        {
                            await notifier.NotifyMaintainersAsync(
                                recipients, record.SystemId, record.Id, record.Summary, isNewSystem: false, cancellationToken);
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "Submission {SubmissionId} was made from the Systems screen but its maintainers could not be told.", outcome.SubmissionId);
                }
            }

            return Results.Ok(new SystemEditAnswer(
                outcome.SubmissionId,
                outcome.Findings,
                outcome.IsPublished,
                outcome.Revision,
                outcome.IsPublished ? outcome.Removals : null,
                outcome.NotPublishedReason));
        }

        // An edit or its check refused, as the status code the flow's flags name.
        private static IResult EditRefusal(SystemEditOutcome outcome)
        {
            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Error });

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsConflict)
                return Results.Conflict(new { error = outcome.Error });

            return Results.BadRequest(new FindingsRefusalAnswer(outcome.FullError, outcome.Findings));
        }

        // ###########################################################################################
        // POST /api/review/systems/files {systemId} - every file the system uses, as BETA holds it.
        // For any maintainer, like the rest of the screen.
        // ###########################################################################################
        private static async Task<IResult> FilesAsync(
            SystemDetailRequest request,
            HttpContext context,
            IAccountStore accounts,
            PublishedBoardReader boards,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            if (!ReviewAuthority.CanReviewAnything(access))
                return ReviewEndpoints.NotAMaintainer();

            string? systemId = request?.SystemId?.Trim();

            // The stable source's files (2026-10-04) - or BETA's, as before. The stable root is the
            // one the overview judged "in the stable source" by.
            if (DataTreeNames.IsProduction(request?.Tree))
            {
                if (options.StableSourceRoot is not string stableRoot)
                    return Results.NotFound(new { error = SystemEditFlow.NoStableSourceMessage });

                IReadOnlyList<SystemFileEntry>? stable = SystemDescriptorRules.IsValidSystemId(systemId)
                    ? await SystemFilesFlow.BuildAsync(stableRoot, systemId!, boards, cancellationToken, SystemFileSource.Production)
                    : null;

                return stable is null
                    ? Results.NotFound(new { error = SystemEditFlow.NotInStableMessage })
                    : Results.Ok(new SystemFilesAnswer(systemId!, stable, BetaDataUrl: null, ProductionDataUrl: options.ProductionPublicDataBaseUrl));
            }

            IReadOnlyList<SystemFileEntry>? files =
                SystemDescriptorRules.IsValidSystemId(systemId) && !string.IsNullOrWhiteSpace(options.DataTreeRoot)
                    ? await SystemFilesFlow.BuildAsync(options.DataTreeRoot, systemId!, boards, cancellationToken)
                    : null;

            return files is null
                ? Results.NotFound(new { error = SystemEditFlow.NotInBetaMessage })
                : Results.Ok(new SystemFilesAnswer(systemId!, files, options.PublicDataBaseUrl));
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
                boardViews,
                listings: SystemEndpoints.Listings(options));

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
                boardViews: boardViews,
                listings: SystemEndpoints.Listings(options));

            if (outcome.IsForbidden)
                return ReviewEndpoints.NotAMaintainer();

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = "No such system." });

            return Results.Ok(outcome.Detail);
        }

        // ###########################################################################################
        // The boards production holds, or null when this server has no production tree to look in -
        // so the screen never says "not in production" about a tree nobody asked. The tree is
        // ServerOptions.StableSourceRoot, which the stable table and files read too.
        // ###########################################################################################
        private static IReadOnlyList<PublishedSystemLister.KnownSystem>? ProductionSystems(ServerOptions options)
        {
            string? root = options.StableSourceRoot;

            return root is null || !Directory.Exists(root)
                ? null
                : PublishedSystemLister.List(root);
        }

        // Which systems BETA's and the stable source's drop-down lists name (2026-10-04).
        private static SystemListings Listings(ServerOptions options) =>
            SystemOverviewFlow.ReadListings(options.DataTreeRoot, options.StableSourceRoot);
    }
}
