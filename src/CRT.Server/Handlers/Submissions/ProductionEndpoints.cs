using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // BETA TO PRODUCTION, over HTTP (maintainer request, 2026-09-25). A rim over
    // ProductionPromotionFlow and nothing else.
    //
    // Under "/api/review/production": the same people as the review queue - an administrator, or
    // a reviewer of the system in question - with the same authentication, and every route asks
    // about the system it names.
    //
    //   GET  /                 - the systems whose BETA state is ahead of production, that the
    //                            caller may publish.
    //   POST /plan    {systemId}                           - what publishing it would copy.
    //   POST /publish {systemId, expectedBetaContentHash}  - do it.
    //
    // POSTs with a body because a system id carries slashes.
    //
    // THE STATUS CODES: 401 / 403 as ReviewEndpoints; 404 no such system; 409 BETA changed since
    // it was checked, or production already has it; 400 the plan refuses; 503 publishing to
    // production is not switched on for this server - the one answer that is about the SERVER
    // rather than the request.
    // ###########################################################################################
    public static class ProductionEndpoints
    {
        public static void MapProductionEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder production = app.MapGroup("/api/review/production");

            production.MapGet("/", ProductionEndpoints.ListAsync);
            production.MapPost("/plan", ProductionEndpoints.PlanAsync);
            production.MapPost("/publish", ProductionEndpoints.PublishAsync)
                .WithBodyLimit(RequestBodyLimits.PathListBytes);
        }

        // The list, the plan and publish bodies, and their answers, are CRT.Data's ReviewApiContract
        // records (ProductionListAnswer, ProductionPlanRequest, ProductionPublishRequest,
        // ProductionPlanAnswer, ProductionPublishAnswer), shared with the review application.

        private static async Task<IResult> ListAsync(
            HttpContext context,
            IAccountStore accounts,
            ProductionPromotionFlow flow,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ProductionEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            IReadOnlyList<SystemRecord> systems = options.IsProductionPublishingConfigured
                ? await flow.ListAwaitingAsync(access, cancellationToken)
                : [];

            // CRT.Data's ProductionListAnswer, read by the review application as the same record.
            // Configured is told rather than inferred from an empty list: "nothing is waiting" and
            // "this server cannot do it" must not look the same.
            return Results.Ok(new ProductionListAnswer(
                options.IsProductionPublishingConfigured,
                systems.Select(system => new ProductionListEntry(
                    system.SystemId,
                    system.Manufacturer,
                    system.Hardware,
                    system.Board,
                    BetaRevision: system.CurrentRevision,
                    BetaContentHash: system.ContentHash,
                    system.ProductionRevision,
                    system.ProductionPublishedUtc)).ToList()));
        }

        private static async Task<IResult> PlanAsync(
            ProductionPlanRequest request,
            HttpContext context,
            IAccountStore accounts,
            ProductionPromotionFlow flow,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ProductionEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            PromotionPlanOutcome outcome = await flow.PlanAsync(access, request?.SystemId, options, cancellationToken);

            if (outcome.IsNotConfigured)
                return Results.Json(new { error = outcome.Refusal }, statusCode: StatusCodes.Status503ServiceUnavailable);

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Refusal }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Refusal });

            SystemRecord system = outcome.System!;
            ProductionPromotionResult plan = outcome.Plan!;

            return Results.Ok(new ProductionPlanAnswer(
                system.SystemId,
                system.CurrentRevision,

                // What the publish request must send back - see ProductionPromotionFlow step 4.
                system.ContentHash,
                ProductionPromotionRules.IsAwaitingProduction(system),
                plan.TouchesSharedFiles,

                // The one answer the button follows, so the app cannot enable something the
                // publish step then refuses: nothing refused, something waiting, and this account
                // still able to add its approval.
                CanPublish: outcome.Refusal is null &&
                    ProductionPromotionRules.IsAwaitingProduction(system) &&
                    outcome.Approval!.CanApprove,
                Refusal: outcome.Refusal,

                // Who must approve, who has, and whether this account's approval publishes -
                // CRT.Data's ApprovalStatus, the same record the submission detail carries.
                Approval: outcome.Approval,
                UnchangedCount: plan.UnchangedCount,

                // CRT.Data's PromotionFile, read by the review application as the same record.
                Files: plan.Files,
                Problems: plan.Problems,

                // What promoting would REMOVE from production - CRT.Data's FileRemovalPreview, shown
                // before anyone approves and sent back with the publish request (2026-09-25).
                Removals: outcome.Removals));
        }

        private static async Task<IResult> PublishAsync(
            ProductionPublishRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            ProductionPromotionFlow flow,
            SubmissionNotifier notifier,
            ServerOptions options,
            ILogger<ProductionPromotionFlow> logger,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ProductionEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            DateTimeOffset now = DateTimeOffset.UtcNow;

            PromotionOutcome outcome = await flow.PromoteAsync(
                access, request?.SystemId, request?.ExpectedBetaContentHash, options, now, request?.ExpectedRemovals, cancellationToken);

            // The first of two approvals: recorded, nothing copied, the other side told.
            if (outcome.IsAwaitingApproval)
            {
                try
                {
                    IReadOnlyList<string> recipients = await SubmissionRouting.RecipientsForRolesAsync(
                        outcome.WaitingFor, outcome.System!.SystemId, accounts, cancellationToken);

                    await notifier.NotifyApprovalNeededAsync(
                        recipients,
                        outcome.System.SystemId,
                        "publishing the board to the source (production)",
                        ApprovePublishFlow.Label(access!),
                        cancellationToken);
                }
                catch (Exception ex)
                {
                    logger.LogWarning(ex, "A production approval was recorded but the other approver could not be told.");
                }

                return Results.Ok(new ProductionPublishAnswer(
                    outcome.System!.SystemId,
                    "awaiting",
                    WaitingFor: outcome.WaitingFor));
            }

            if (!outcome.IsPublished)
            {
                if (outcome.IsNotConfigured)
                    return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status503ServiceUnavailable);

                if (outcome.IsForbidden)
                    return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

                if (outcome.IsNotFound)
                    return Results.NotFound(new { error = outcome.Error });

                return outcome.IsConflict
                    ? Results.Conflict(new { error = outcome.Error })
                    : Results.BadRequest(new { error = outcome.Error });
            }

            // ###########################################################################################
            // *** EVERYTHING FROM HERE IS AFTER THE FACT, AND NONE OF IT MAY FAIL THE REQUEST. ***
            // Production has been written - the same position ReviewEndpoints.ApproveAsync is in
            // after a publish, and for the same reason: an error here would be read as "it did not
            // happen" and retried.
            // ###########################################################################################
            try
            {
                await ProductionEndpoints.AfterPublishAsync(
                    access!, outcome, accounts, submissions, notifier, options, now, logger, cancellationToken);
            }
            catch (Exception ex)
            {
                logger.LogError(
                    ex,
                    "{SystemId} WAS published to production, but the follow-up work failed. The data is correct; " +
                    "the production checksum manifest or the notification mails may not be.",
                    outcome.System!.SystemId);
            }

            return Results.Ok(new ProductionPublishAnswer(
                outcome.System!.SystemId,
                "published",
                Revision: outcome.System.CurrentRevision,
                FilesCopied: outcome.FilesCopied,
                RemovedFiles: outcome.RemovedFiles));
        }

        // ###########################################################################################
        // The production manifest, then who is told.
        //
        // THE MANIFEST FIRST: until it advertises the new checksums, no client downloads anything
        // - the same lesson ReviewEndpoints.AfterPublishAsync records for BETA.
        //
        // THEN the contributors whose work went out with this promotion (merged since the previous
        // one), and the administrators when a reviewer did it.
        // ###########################################################################################
        private static async Task AfterPublishAsync(
            ReviewAccess access,
            PromotionOutcome outcome,
            IAccountStore accounts,
            ISubmissionStore submissions,
            SubmissionNotifier notifier,
            ServerOptions options,
            DateTimeOffset now,
            ILogger logger,
            CancellationToken cancellationToken)
        {
            SystemRecord system = outcome.System!;

            int written = DataChecksumManifest.Write(
                options.ProductionDataTreeRoot ?? string.Empty,
                options.ProductionPublicDataBaseUrl ?? string.Empty,
                options.ProductionManifestPath ?? string.Empty);

            if (written < 0)
            {
                logger.LogWarning(
                    "{SystemId} was published to production but the production checksum manifest at [{Path}] " +
                    "could not be regenerated - clients will not see it until it is rebuilt.",
                    system.SystemId, options.ProductionManifestPath);
            }

            IReadOnlyList<SubmissionRecord> carried = await submissions.GetMergedSubmissionsAsync(
                system.SystemId, outcome.PreviousProductionPublishedUtc, now, cancellationToken);

            foreach (SubmissionRecord submission in carried)
            {
                await notifier.NotifyDecisionAsync(
                    submission.ContactEmail,
                    submission.SystemId,
                    ProductionPromotionRules.PublishedState,
                    reviewerComment: null,
                    cancellationToken: cancellationToken);
            }

            if (!access.Account.IsAdministrator)
            {
                IReadOnlyList<AccountRecord> administrators = await accounts.GetAdministratorsAsync(cancellationToken);

                await notifier.NotifyProductionPublishAsync(
                    administrators.Where(admin => admin.IsVerified && !admin.IsLocked).Select(admin => admin.Email),
                    system.SystemId,
                    $"{access.Account.DisplayName} ({access.Account.Email})",
                    system.CurrentRevision,
                    outcome.FilesCopied,
                    cancellationToken);
            }
        }

        private static async Task<(ReviewAccess? Access, IResult? Refusal)> AuthoriseAsync(
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return (null, refusal);

            if (!ReviewAuthority.CanReviewAnything(access))
                return (null, ReviewEndpoints.NotAReviewer());

            return (access, null);
        }
    }
}
