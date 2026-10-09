using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Usage;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The ADMINISTRATOR's API (Phase 6 roles, 2026-09-25): which accounts exist, which boards
    // exist, and who reviews what - a rim over MaintainerAssignmentFlows - the "Unused files"
    // screen, a rim over UnusedFileFlows, deleting a board, a rim over BoardDeletionFlow,
    // resetting the contribution data (DataResetFlow) and the API usage counts (ApiUsageFlow).
    //
    // Its own group, "/api/admin", rather than routes under "/api/review": everything here is
    // administrator-only, and a group with one rule cannot have a route mapped into it that
    // quietly inherits the wrong one.
    //
    // THE STATUS CODES match ReviewEndpoints: 401 no credentials, 403 signed in but not an
    // administrator, 404 no such board or account, 400 a grant the rules refuse (unverified,
    // locked, already an administrator).
    // ###########################################################################################
    public static class AdminEndpoints
    {
        public static void MapAdminEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder admin = app.MapGroup("/api/admin");

            admin.MapGet("/boards", AdminEndpoints.GetBoardsAsync);
            admin.MapGet("/accounts", AdminEndpoints.GetAccountsAsync);

            // Both POST with a body. A board id carries slashes, so it cannot ride in the route
            // without a catch-all, and a catch-all cannot be followed by the account id.
            admin.MapPost("/maintainers", AdminEndpoints.AddMaintainerAsync);
            admin.MapPost("/maintainers/remove", AdminEndpoints.RemoveMaintainerAsync);

            // Inviting an address with no account yet, and taking an invitation back - see
            // MaintainerInvitationFlows. Both answer { message }.
            admin.MapPost("/maintainers/invite", AdminEndpoints.InviteMaintainerAsync);
            admin.MapPost("/maintainers/invitations/withdraw", AdminEndpoints.WithdrawInvitationAsync);

            // Files nothing uses, per tree, and removing the ones the administrator chose - see
            // UnusedFileFlows. The tree is "beta" or "production".
            admin.MapGet("/unused-files", AdminEndpoints.ListUnusedFilesAsync);
            admin.MapPost("/unused-files/remove", AdminEndpoints.RemoveUnusedFilesAsync)
                .WithBodyLimit(RequestBodyLimits.PathListBytes);

            // Rebuilding every tree's dataChecksums.json by hand, after the data was edited on the
            // box rather than through the service - see ManifestRebuildFlow. No body: the button
            // does both trees, so there is nothing to send and no body limit to set.
            admin.MapPost("/manifest/rebuild", AdminEndpoints.RebuildManifestsAsync);

            // Deleting a board completely (owner request, 2026-10-03) - see BoardDeletionFlow. The
            // plan first, shown in the confirmation; the delete is held to it by its fingerprint.
            // Both POST a body, because a board id carries slashes; both bodies are small.
            admin.MapPost("/boards/delete/plan", AdminEndpoints.PlanBoardDeletionAsync);
            admin.MapPost("/boards/delete", AdminEndpoints.DeleteBoardAsync);

            // The order of CRT's drop-down lists, in BETA and the stable source (owner request,
            // 2026-10-04) - see BoardOrderFlow. The body is every board id BETA lists, so it gets
            // the path-list limit rather than the 64 KB default.
            admin.MapPost("/boards/order", AdminEndpoints.SetBoardOrderAsync)
                .WithBodyLimit(RequestBodyLimits.PathListBytes);

            // Resetting the contribution data for going live (owner request, 2026-10-04) - see
            // DataResetFlow. The counts first; the reset is held to them by their fingerprint, and
            // happens only while the server's AllowDataReset setting is on. A fingerprint is small.
            admin.MapGet("/reset", AdminEndpoints.PlanDataResetAsync);
            admin.MapPost("/reset", AdminEndpoints.ResetDataAsync);

            // Which CRT versions call which route (owner request, 2026-10-04) - see ApiUsageFlow.
            admin.MapGet("/api-usage", AdminEndpoints.GetApiUsageAsync);
        }

        // ###########################################################################################
        // GET /api/admin/reset - what a reset would delete, and whether this server allows one. Read
        // only; answered while the reset is switched off too, so the screen can say how to switch it on.
        // ###########################################################################################
        private static async Task<IResult> PlanDataResetAsync(
            HttpContext context,
            IAccountStore accounts,
            IDataResetStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            DataResetOutcome outcome = await DataResetFlow.PlanAsync(access, options, store, cancellationToken);

            return AdminEndpoints.RefusalFor(outcome) ?? Results.Ok(outcome.Plan);
        }

        // ###########################################################################################
        // POST /api/admin/reset  { fingerprint } - deletes it. 503 switched off, 409 changed since the
        // counts were shown, 500 the database refused (nothing deleted), 200 with what went.
        //
        // After the rows: every deleted submission's partial uploads, and every stored file no
        // submission needs any more - the blob store's own collection, run now.
        // ###########################################################################################
        private static async Task<IResult> ResetDataAsync(
            DataResetRequest request,
            HttpContext context,
            IAccountStore accounts,
            IDataResetStore store,
            ISubmissionStore submissions,
            BlobStore blobs,
            ApiUsageCounter usage,
            ServerOptions options,
            PublishLock publishLock,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            DataResetOutcome outcome = await DataResetFlow.ResetAsync(
                access,
                request?.Fingerprint,
                options,
                store,
                publishLock,
                async (submissionIds, token) =>
                {
                    foreach (long submissionId in submissionIds)
                        blobs.ClearPartials(submissionId);

                    return await SubmissionFlows.CollectUnreferencedBlobsAsync(submissions, blobs, token);
                },
                usage,
                DateTimeOffset.UtcNow,
                loggerFactory.CreateLogger(typeof(DataResetFlow).FullName!),
                cancellationToken);

            return AdminEndpoints.RefusalFor(outcome) ?? Results.Ok(outcome.Answer);
        }

        private static IResult? RefusalFor(DataResetOutcome outcome)
        {
            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotEnabled)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status503ServiceUnavailable);

            if (outcome.IsConflict)
                return Results.Conflict(new { error = outcome.Error });

            return outcome.Error is not null
                ? Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status500InternalServerError)
                : null;
        }

        // ###########################################################################################
        // GET /api/admin/api-usage?days=90 - every route, the CRT versions that called it in the last
        // `days` days (90 when left out, at most a year), and the launches per version.
        // ###########################################################################################
        private static async Task<IResult> GetApiUsageAsync(
            int? days,
            HttpContext context,
            IAccountStore accounts,
            IApiUsageStore store,
            ICheckInStore checkIns,
            ApiRouteList routes,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            ApiUsageAnswer? answer = await ApiUsageFlow.ReadAsync(
                access,
                days,
                routes.All(),
                store,
                checkIns,
                DateTimeOffset.UtcNow,
                loggerFactory.CreateLogger(typeof(ApiUsageFlow).FullName!),
                cancellationToken);

            return Results.Ok(answer);
        }

        // ###########################################################################################
        // POST /api/admin/boards/order  { boardIds } - every board in BETA's drop-down lists, in
        // the order wanted. 400 for a list that is not one, 409 for one that no longer matches BETA's
        // (or a file that cannot be read or written), 200 with what was rewritten.
        // ###########################################################################################
        private static async Task<IResult> SetBoardOrderAsync(
            BoardOrderRequest request,
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            PublishLock publishLock,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            BoardOrderOutcome outcome = await BoardOrderFlow.SetAsync(
                access!,
                request,
                options,
                tree => DataChecksumManifest.Write(tree.Root, tree.PublicBaseUrl, tree.ManifestPath),
                publishLock,
                accounts,
                DateTimeOffset.UtcNow,
                cancellationToken,
                loggerFactory.CreateLogger(typeof(BoardOrderFlow).FullName!));

            if (outcome.Answer is not null)
                return Results.Ok(outcome.Answer);

            return outcome.IsConflict
                ? Results.Conflict(new { error = outcome.Error })
                : Results.BadRequest(new { error = outcome.Error });
        }

        // The two lists, and the bodies of changing a pool and removing unused files, are CRT.Data's
        // ReviewApiContract records (MaintainerBoardsAnswer, MaintainerAccountsAnswer,
        // MaintainerChangeRequest, UnusedFilesRemoveRequest, UnusedFilesRemoveAnswer), shared with the
        // Maintainer tab.

        // ###########################################################################################
        // GET /api/admin/boards - every board with its maintainers.
        // ###########################################################################################
        private static async Task<IResult> GetBoardsAsync(
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            IReadOnlyList<BoardWithMaintainers> boards = await MaintainerAssignmentFlows.ListBoardsAsync(
                PublishedBoardLister.List(options.DataTreeRoot), submissions, accounts, cancellationToken);

            return Results.Ok(new MaintainerBoardsAnswer(
                boards.Select(board => new MaintainerBoardEntry(
                    board.BoardId,
                    board.Manufacturer,
                    board.Hardware,
                    board.Board,
                    board.CurrentRevision,
                    board.IsAccepting,
                    board.Maintainers.Select(AdminEndpoints.ToMaintainer).ToList())).ToList()));
        }

        // ###########################################################################################
        // GET /api/admin/accounts - every account, to pick a maintainer from.
        //
        // *** NO PASSWORD HASH, no session, no token - only what the administrator needs to
        // recognise a person and see whether they can be granted anything. ***
        // ###########################################################################################
        private static async Task<IResult> GetAccountsAsync(
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            IReadOnlyList<AccountRecord> all =
                await accounts.ListAccountsAsync(MaintainerAssignmentFlows.AccountListLimit, cancellationToken);

            return Results.Ok(new MaintainerAccountsAnswer(
                all.Select(account => new MaintainerAccountEntry(
                    account.Id,
                    account.Email,
                    account.DisplayName,
                    account.IsAdministrator,
                    account.IsVerified,
                    account.IsLocked)).ToList()));
        }

        // POST /api/admin/maintainers  { boardId, accountId }
        private static async Task<IResult> AddMaintainerAsync(
            MaintainerChangeRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            MaintainerAssignmentOutcome outcome = await MaintainerAssignmentFlows.AddAsync(
                access,
                request?.BoardId,
                request?.AccountId ?? 0,
                PublishedBoardLister.List(options.DataTreeRoot),
                accounts,
                submissions,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return AdminEndpoints.ToResult(outcome, request);
        }

        // POST /api/admin/maintainers/remove  { boardId, accountId }
        private static async Task<IResult> RemoveMaintainerAsync(
            MaintainerChangeRequest request,
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            MaintainerAssignmentOutcome outcome = await MaintainerAssignmentFlows.RemoveAsync(
                access, request?.BoardId, request?.AccountId ?? 0, accounts, DateTimeOffset.UtcNow, cancellationToken);

            return AdminEndpoints.ToResult(outcome, request);
        }

        // POST /api/admin/maintainers/invite  { boardId, email }
        private static async Task<IResult> InviteMaintainerAsync(
            MaintainerInviteRequest request,
            HttpContext context,
            IAccountStore accounts,
            ISubmissionStore submissions,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlows.InviteAsync(
                access,
                request?.BoardId,
                request?.Email,
                PublishedBoardLister.List(options.DataTreeRoot),
                accounts,
                submissions,
                mailer,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return AdminEndpoints.ToResult(outcome);
        }

        // POST /api/admin/maintainers/invitations/withdraw  { invitationId }
        private static async Task<IResult> WithdrawInvitationAsync(
            InvitationWithdrawRequest request,
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlows.WithdrawAsync(
                access, request?.InvitationId ?? 0, accounts, DateTimeOffset.UtcNow, cancellationToken);

            return AdminEndpoints.ToResult(outcome);
        }

        private static IResult ToResult(MaintainerInvitationOutcome outcome)
        {
            if (outcome.IsDone)
                return Results.Ok(new { message = outcome.Message });

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Message }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Message });

            return Results.BadRequest(new { error = outcome.Message });
        }

        // ###########################################################################################
        // GET /api/admin/unused-files?tree=beta - every file that tree does not use. Reads every
        // workbook in the tree, so it takes a few seconds; nothing is written.
        // ###########################################################################################
        private static async Task<IResult> ListUnusedFilesAsync(
            string? tree,
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            if (!UnusedFileFlows.TryResolve(tree, options, out UnusedFileFlows.DataTree? target, out string error))
                return Results.BadRequest(new { error });

            return Results.Ok(await Task.Run(() => UnusedFileFlows.List(target!), cancellationToken));
        }

        // POST /api/admin/unused-files/remove  { tree, files }
        private static async Task<IResult> RemoveUnusedFilesAsync(
            UnusedFilesRemoveRequest request,
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            PublishLock publishLock,
            ILogger<UnusedFileRemoval> logger,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            if (!UnusedFileFlows.TryResolve(request?.Tree, options, out UnusedFileFlows.DataTree? target, out string error))
                return Results.BadRequest(new { error });

            UnusedFileRemoval removal = await UnusedFileFlows.RemoveAsync(
                access!,
                target!,
                (request?.Files ?? []).ToList(),
                publishLock,
                accounts,
                logger,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return Results.Ok(new UnusedFilesRemoveAnswer(
                target!.Name,
                removal.Removed,
                removal.Kept,
                removal.NotDoneBecause));
        }

        // ###########################################################################################
        // POST /api/admin/manifest/rebuild - rebuild dataChecksums.json for every configured tree.
        //
        // The scan hashes every file in the tree, so it runs on a pool thread rather than holding
        // the request thread; ManifestRebuildFlow takes the publish lock around it.
        // ###########################################################################################
        private static async Task<IResult> RebuildManifestsAsync(
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            PublishLock publishLock,
            ILoggerFactory loggerFactory,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            IReadOnlyList<ManifestRebuildFlow.TreeOutcome> outcomes = await ManifestRebuildFlow.RebuildAsync(
                access!,
                options,
                tree => DataChecksumManifest.Write(tree.Root, tree.PublicBaseUrl, tree.ManifestPath),
                publishLock,
                accounts,
                DateTimeOffset.UtcNow,
                cancellationToken,
                loggerFactory.CreateLogger(typeof(ManifestRebuildFlow).FullName!));

            return Results.Ok(new ManifestRebuildAnswer(
                ManifestRebuildFlow.Headline(outcomes),
                [.. outcomes.Select(outcome => new ManifestRebuildEntry(
                    outcome.Tree, outcome.Skipped, outcome.Entries, outcome.Message))]));
        }

        // ###########################################################################################
        // POST /api/admin/boards/delete/plan  { boardId } - what deleting it would remove. Reads
        // every workbook in both trees when the board has files, so it takes seconds; writes nothing.
        // A plan that cannot go ahead is still 200, carrying blockedBecause.
        // ###########################################################################################
        private static async Task<IResult> PlanBoardDeletionAsync(
            BoardDetailRequest request,
            HttpContext context,
            IAccountStore accounts,
            BoardDeletionFlow flow,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            BoardDeletionOutcome outcome = await flow.PlanAsync(
                access, request?.BoardId, options, DateTimeOffset.UtcNow, cancellationToken);

            return AdminEndpoints.RefusalFor(outcome) ?? Results.Ok(outcome.Plan!.ToAnswer());
        }

        // POST /api/admin/boards/delete  { boardId, fingerprint, reason }
        private static async Task<IResult> DeleteBoardAsync(
            BoardDeleteRequest request,
            HttpContext context,
            IAccountStore accounts,
            BoardDeletionFlow flow,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            BoardDeletionOutcome outcome = await flow.DeleteAsync(
                access, request?.BoardId, request?.Fingerprint, request?.Reason, options, DateTimeOffset.UtcNow, cancellationToken);

            if (AdminEndpoints.RefusalFor(outcome) is IResult problem)
                return problem;

            return Results.Ok(new BoardDeleteAnswer(
                outcome.Plan!.BoardId,
                outcome.BetaFilesRemoved,
                outcome.ProductionFilesRemoved,
                outcome.Plan.Submissions.Count,
                outcome.ContributorsMailed));
        }

        // The same codes BETA's push-back answers: 503 not switched on, 403, 404, 409 blocked or
        // changed since it was shown, 400 anything else refused.
        private static IResult? RefusalFor(BoardDeletionOutcome outcome)
        {
            if (outcome.IsNotConfigured)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status503ServiceUnavailable);

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Error });

            if (outcome.IsConflict)
                return Results.Conflict(new { error = outcome.Error });

            return outcome.Error is not null ? Results.BadRequest(new { error = outcome.Error }) : null;
        }

        private static IResult ToResult(MaintainerAssignmentOutcome outcome, MaintainerChangeRequest? request)
        {
            if (outcome.IsDone)
                return Results.Ok(new MaintainerChangeAnswer(request?.BoardId, request?.AccountId));

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Error });

            return Results.BadRequest(new { error = outcome.Error });
        }

        private static PoolMaintainerEntry ToMaintainer(MaintainerRecord maintainer) =>
            new(maintainer.AccountId, maintainer.DisplayName, maintainer.Email);

        // ###########################################################################################
        // The same authentication ReviewEndpoints performs, with the administrator rule on top.
        // ###########################################################################################
        private static async Task<(ReviewAccess? Access, IResult? Refusal)> AuthoriseAsync(
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await ReviewEndpoints.AuthenticateAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return (null, refusal);

            if (!ReviewAuthority.CanAdminister(access))
            {
                return (null, Results.Json(
                    new { error = "Only an administrator can do this." },
                    statusCode: StatusCodes.Status403Forbidden));
            }

            return (access, null);
        }
    }
}
