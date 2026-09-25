using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The ADMINISTRATOR's API (Phase 6 roles, 2026-09-25): which accounts exist, which systems
    // exist, and who reviews what - a rim over ReviewerAssignmentFlows - and the "Unused files"
    // screen, a rim over UnusedFileFlows.
    //
    // Its own group, "/api/admin", rather than routes under "/api/review": everything here is
    // administrator-only, and a group with one rule cannot have a route mapped into it that
    // quietly inherits the wrong one.
    //
    // THE STATUS CODES match ReviewEndpoints: 401 no credentials, 403 signed in but not an
    // administrator, 404 no such system or account, 400 a grant the rules refuse (unverified,
    // locked, already an administrator).
    // ###########################################################################################
    public static class AdminEndpoints
    {
        public static void MapAdminEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder admin = app.MapGroup("/api/admin");

            admin.MapGet("/systems", AdminEndpoints.GetSystemsAsync);
            admin.MapGet("/accounts", AdminEndpoints.GetAccountsAsync);

            // Both POST with a body. A system id carries slashes, so it cannot ride in the route
            // without a catch-all, and a catch-all cannot be followed by the account id.
            admin.MapPost("/reviewers", AdminEndpoints.AddReviewerAsync);
            admin.MapPost("/reviewers/remove", AdminEndpoints.RemoveReviewerAsync);

            // Files nothing uses, per tree, and removing the ones the administrator chose - see
            // UnusedFileFlows. The tree is "beta" or "production".
            admin.MapGet("/unused-files", AdminEndpoints.ListUnusedFilesAsync);
            admin.MapPost("/unused-files/remove", AdminEndpoints.RemoveUnusedFilesAsync)
                .WithBodyLimit(RequestBodyLimits.PathListBytes);
        }

        // The two lists, and the bodies of changing a pool and removing unused files, are CRT.Data's
        // ReviewApiContract records (ReviewerSystemsAnswer, ReviewerAccountsAnswer,
        // ReviewerChangeRequest, UnusedFilesRemoveRequest, UnusedFilesRemoveAnswer), shared with the
        // review application.

        // ###########################################################################################
        // GET /api/admin/systems - every system with its reviewers.
        // ###########################################################################################
        private static async Task<IResult> GetSystemsAsync(
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

            IReadOnlyList<SystemWithReviewers> systems = await ReviewerAssignmentFlows.ListSystemsAsync(
                PublishedSystemLister.List(options.DataTreeRoot), submissions, accounts, cancellationToken);

            return Results.Ok(new ReviewerSystemsAnswer(
                systems.Select(system => new ReviewerSystemEntry(
                    system.SystemId,
                    system.Manufacturer,
                    system.Hardware,
                    system.Board,
                    system.CurrentRevision,
                    system.IsAccepting,
                    system.Reviewers.Select(AdminEndpoints.ToReviewer).ToList())).ToList()));
        }

        // ###########################################################################################
        // GET /api/admin/accounts - every account, to pick a reviewer from.
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
                await accounts.ListAccountsAsync(ReviewerAssignmentFlows.AccountListLimit, cancellationToken);

            return Results.Ok(new ReviewerAccountsAnswer(
                all.Select(account => new ReviewerAccountEntry(
                    account.Id,
                    account.Email,
                    account.DisplayName,
                    account.IsAdministrator,
                    account.IsVerified,
                    account.IsLocked)).ToList()));
        }

        // POST /api/admin/reviewers  { systemId, accountId }
        private static async Task<IResult> AddReviewerAsync(
            ReviewerChangeRequest request,
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

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.AddAsync(
                access,
                request?.SystemId,
                request?.AccountId ?? 0,
                PublishedSystemLister.List(options.DataTreeRoot),
                accounts,
                submissions,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return AdminEndpoints.ToResult(outcome, request);
        }

        // POST /api/admin/reviewers/remove  { systemId, accountId }
        private static async Task<IResult> RemoveReviewerAsync(
            ReviewerChangeRequest request,
            HttpContext context,
            IAccountStore accounts,
            CancellationToken cancellationToken)
        {
            (ReviewAccess? access, IResult? refusal) =
                await AdminEndpoints.AuthoriseAsync(context, accounts, cancellationToken);

            if (refusal is not null)
                return refusal;

            ReviewerAssignmentOutcome outcome = await ReviewerAssignmentFlows.RemoveAsync(
                access, request?.SystemId, request?.AccountId ?? 0, accounts, DateTimeOffset.UtcNow, cancellationToken);

            return AdminEndpoints.ToResult(outcome, request);
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

        private static IResult ToResult(ReviewerAssignmentOutcome outcome, ReviewerChangeRequest? request)
        {
            if (outcome.IsDone)
                return Results.Ok(new { systemId = request?.SystemId, accountId = request?.AccountId });

            if (outcome.IsForbidden)
                return Results.Json(new { error = outcome.Error }, statusCode: StatusCodes.Status403Forbidden);

            if (outcome.IsNotFound)
                return Results.NotFound(new { error = outcome.Error });

            return Results.BadRequest(new { error = outcome.Error });
        }

        private static PoolReviewerEntry ToReviewer(ReviewerRecord reviewer) =>
            new(reviewer.AccountId, reviewer.DisplayName, reviewer.Email);

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
