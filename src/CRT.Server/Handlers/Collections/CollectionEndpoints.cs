using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // THE COLLECTION ROUTES CRT SHARES THROUGH (2026-11-26) - a RIM over CollectionFlows: each reads the request into plain
    // values, calls one flow, and maps its verdict onto a status code. CollectionContract says how CRT uses them.
    //
    //   PUT    /api/collections/images/{sha256}  a capture's PNG bytes as the body      200 CollectionImageUploadAnswer
    //   POST   /api/collections                   CollectionContentRequest               201 CollectionAnswer
    //   PUT    /api/collections/{id}              CollectionContentRequest               200 CollectionAnswer
    //   DELETE /api/collections/{id}                                                     204
    //   PUT    /api/collections/{id}/sharing      CollectionSharingRequest               200 CollectionAnswer (15.1.0)
    //   GET    /api/collections/mine                                                     200 MyCollectionsAnswer
    //
    // Every refusal is CollectionRefusalAnswer ({ errors: [...] }, and missingImages with 409). Each needs a signed-in account
    // (401 without), under a per-address request limit of their own (RateLimitPolicy) - sharing a chip's 40 pins is some 42
    // requests. Not set up on this server (CollectionsRoot): 503, the same sentence for every route (the group's filter).
    // ###########################################################################################
    public static class CollectionEndpoints
    {
        public const string RateLimitPolicy = "collections";

        // Far above one person sharing - a board's worth of chips in an hour is a few hundred requests.
        public const int PermitsPerAddressPerHour = 2000;

        public static void AddCollectionRateLimit(this IServiceCollection services) =>
            services.AddPerAddressRateLimit(CollectionEndpoints.RateLimitPolicy, CollectionEndpoints.PermitsPerAddressPerHour);

        public static void MapCollectionEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder collections = app.MapGroup(CollectionRoutes.Collections).RequireRateLimiting(CollectionEndpoints.RateLimitPolicy);

            // Not set up (CollectionsRoot): EVERY route answers 503 with the same sentence, asked once for the group (code
            // review, 2026-11-27: only the upload, the share and its update asked - the comments, the delete and the list
            // answered from the database as if sharing were on).
            collections.AddEndpointFilter(async (invocation, next) =>
                invocation.HttpContext.RequestServices.GetRequiredService<ServerOptions>().IsCollectionSharingConfigured
                    ? await next(invocation)
                    : CollectionEndpoints.Result(CollectionOutcome.Refused(CollectionVerdict.Off, CollectionFlows.NotSetUp)));

            collections.MapPut("/images/{sha256}", CollectionEndpoints.UploadImageAsync).WithBodyLimit(RequestBodyLimits.CollectionImageBytes);
            collections.MapGet("/mine", CollectionEndpoints.ListMineAsync);
            collections.MapPost("", CollectionEndpoints.CreateAsync).WithBodyLimit(RequestBodyLimits.CollectionContentBytes);
            collections.MapPut("/{id}", CollectionEndpoints.UpdateAsync).WithBodyLimit(RequestBodyLimits.CollectionContentBytes);
            collections.MapDelete("/{id}", CollectionEndpoints.DeleteAsync);
            collections.MapPut("/{id}/sharing", CollectionEndpoints.SetSharingAsync);
            collections.MapDelete("/{id}/images/{sha256}", CollectionEndpoints.RemoveImageAsync);

            // Its comments, for the owner's CRT (step 3) - CollectionCommentFlows. Small bodies, at the default limit.
            collections.MapGet("/{id}/comments", CollectionEndpoints.ListCommentsAsync);
            collections.MapPost("/{id}/comments", CollectionEndpoints.AddCommentAsync);
            collections.MapPut("/{id}/comments/{commentId:long}", CollectionEndpoints.EditCommentAsync);
            collections.MapDelete("/{id}/comments/{commentId:long}", CollectionEndpoints.DeleteCommentAsync);
            collections.MapPut("/{id}/comments/{commentId:long}/hidden", CollectionEndpoints.HideCommentAsync);
        }

        private static async Task<IResult> ListCommentsAsync(
            string id,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionCommentFlows.ListForOwnerAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken), id, store, DateTimeOffset.UtcNow, cancellationToken));

        private static async Task<IResult> AddCommentAsync(
            string id,
            CollectionCommentRequest? request,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionCommentFlows.AddAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                request?.ImageSha256,
                request?.Body,
                store,
                accounts,
                mailer,
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));

        private static async Task<IResult> EditCommentAsync(
            string id,
            long commentId,
            CollectionCommentEditRequest? request,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionCommentFlows.EditAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                commentId,
                request?.Body,
                store,
                accounts,
                DateTimeOffset.UtcNow,
                cancellationToken));

        private static async Task<IResult> DeleteCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionCommentFlows.DeleteAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken), id, commentId, store, DateTimeOffset.UtcNow, cancellationToken));

        private static async Task<IResult> HideCommentAsync(
            string id,
            long commentId,
            CollectionCommentHideRequest? request,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionCommentFlows.SetHiddenAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                commentId,
                request?.Hidden ?? true,
                store,
                DateTimeOffset.UtcNow,
                cancellationToken));

        // A comment's verdict as HTTP: done is 204, or 200 with the comments asked for.
        internal static IResult Result(CommentOutcome outcome) => outcome.Verdict switch
        {
            CollectionVerdict.Done when outcome.Comments is not null => Results.Ok(outcome.Comments),
            CollectionVerdict.Done => Results.NoContent(),
            _ => Results.Json(new CollectionRefusalAnswer(outcome.Errors), statusCode: CollectionEndpoints.StatusOf(outcome.Verdict))
        };

        private static async Task<IResult> UploadImageAsync(
            string sha256,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            // Signed in first, the body read only then (code review, 2026-11-27, Security: anybody made the service read and
            // hold up to 2 MB a request before it answered 401). The flow refuses a missing account as it always did.
            AccountRecord? account = await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken);

            if (account is null)
                return CollectionEndpoints.Result(await CollectionFlows.UploadImageAsync(null, sha256, [], store, new CollectionFiles(options.CollectionsRoot!), CollectionEndpoints.StorageOf(options), DateTimeOffset.UtcNow, cancellationToken));

            byte[] bytes;

            using (var body = new MemoryStream())
            {
                await context.Request.Body.CopyToAsync(body, cancellationToken);
                bytes = body.ToArray();
            }

            return CollectionEndpoints.Result(await CollectionFlows.UploadImageAsync(
                account,
                sha256,
                bytes,
                store,
                new CollectionFiles(options.CollectionsRoot!),
                CollectionEndpoints.StorageOf(options),
                DateTimeOffset.UtcNow,
                cancellationToken));
        }

        private static async Task<IResult> CreateAsync(
            CollectionContentRequest? content,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            return CollectionEndpoints.Result(await CollectionFlows.CreateAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                content,
                context.Request.Headers[CollectionRoutes.ClientKeyHeader].ToString(),
                store,
                new CollectionFiles(options.CollectionsRoot!),
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));
        }

        private static async Task<IResult> UpdateAsync(
            string id,
            CollectionContentRequest? content,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            return CollectionEndpoints.Result(await CollectionFlows.UpdateAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                content,
                store,
                new CollectionFiles(options.CollectionsRoot!),
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));
        }

        private static async Task<IResult> DeleteAsync(
            string id,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.DeleteAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                store,
                DateTimeOffset.UtcNow,
                cancellationToken));

        // "Delete image" in CRT - the capture off the page at once, nothing else changed.
        private static async Task<IResult> RemoveImageAsync(
            string id,
            string sha256,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.RemoveImageAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                id,
                sha256,
                store,
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));

        // "Stop sharing" in CRT, or sharing again - a request with no body is no answer to either.
        private static async Task<IResult> SetSharingAsync(
            string id,
            CollectionSharingRequest? request,
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            AccountRecord? account = await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken);

            if (account is not null && request is null)
                return CollectionEndpoints.Result(CollectionOutcome.Refused(CollectionVerdict.BadRequest, "Say whether it is shared."));

            return CollectionEndpoints.Result(await CollectionFlows.SetSharedAsync(
                account,
                id,
                request?.Shared ?? false,
                store,
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));
        }

        private static async Task<IResult> ListMineAsync(
            HttpContext context,
            IAccountStore accounts,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.ListMineAsync(
                await CollectionEndpoints.AccountOfAsync(context, accounts, options, cancellationToken),
                store,
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));

        // The account the request was sent with - the session slid, as every signed-in route slides it.
        private static Task<AccountRecord?> AccountOfAsync(HttpContext context, IAccountStore accounts, ServerOptions options, CancellationToken cancellationToken) =>
            AccountFlows.AuthenticateAsync(AccountEndpoints.BearerToken(context), accounts, DateTimeOffset.UtcNow, cancellationToken, options);

        private static CollectionStorage StorageOf(ServerOptions options) =>
            new(options.CollectionsMaxStoredBytes, () => Program.FreeBytesAt(options.CollectionsRoot!), options.MinimumFreeDiskBytes);

        // ###########################################################################################
        // A verdict as HTTP - one table, so every route answers alike.
        // ###########################################################################################
        internal static IResult Result(CollectionOutcome outcome) => outcome.Verdict switch
        {
            CollectionVerdict.Created => Results.Created(outcome.Answer!.Url, outcome.Answer),
            CollectionVerdict.Removed => Results.NoContent(),
            CollectionVerdict.Done when outcome.Upload is not null => Results.Ok(outcome.Upload),
            CollectionVerdict.Done when outcome.Mine is not null => Results.Ok(outcome.Mine),
            CollectionVerdict.Done => Results.Ok(outcome.Answer),
            _ => Results.Json(new CollectionRefusalAnswer(outcome.Errors, outcome.Missing), statusCode: CollectionEndpoints.StatusOf(outcome.Verdict))
        };

        internal static int StatusOf(CollectionVerdict verdict) => verdict switch
        {
            CollectionVerdict.Unauthorized => StatusCodes.Status401Unauthorized,
            CollectionVerdict.Forbidden => StatusCodes.Status403Forbidden,
            CollectionVerdict.NotFound => StatusCodes.Status404NotFound,
            CollectionVerdict.Gone => StatusCodes.Status410Gone,
            CollectionVerdict.MissingImages => StatusCodes.Status409Conflict,
            CollectionVerdict.TooMany => StatusCodes.Status429TooManyRequests,
            CollectionVerdict.NoRoom or CollectionVerdict.Off => StatusCodes.Status503ServiceUnavailable,
            _ => StatusCodes.Status400BadRequest
        };
    }
}
