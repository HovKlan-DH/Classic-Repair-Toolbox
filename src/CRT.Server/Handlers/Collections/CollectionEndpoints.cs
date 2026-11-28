using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
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
    //   DELETE /api/collections/{id}/images/{sha256}                                     200 CollectionAnswer (16.1.0)
    //
    // Every refusal is CollectionRefusalAnswer ({ errors: [...] }, and missingImages with 409). Each carries the collection's
    // KEY (CollectionRoutes.OwnerKeyHeader; 400 without - no account since server 17.0.0) and is counted by the internet
    // connection it came from, under a per-address request limit of their own (RateLimitPolicy) - sharing a chip's 40 pins is
    // some 42 requests. Not set up on this server (CollectionsRoot): 503, the same sentence for every route (the group's filter).
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
            // review, 2026-11-27: only the upload, the share and its update asked - the delete answered from the database as if
            // sharing were on).
            collections.AddEndpointFilter(async (invocation, next) =>
                invocation.HttpContext.RequestServices.GetRequiredService<ServerOptions>().IsCollectionSharingConfigured
                    ? await next(invocation)
                    : CollectionEndpoints.Result(CollectionOutcome.Refused(CollectionVerdict.Off, CollectionFlows.NotSetUp)));

            collections.MapPut("/images/{sha256}", CollectionEndpoints.UploadImageAsync).WithBodyLimit(RequestBodyLimits.CollectionImageBytes);
            collections.MapPost("", CollectionEndpoints.CreateAsync).WithBodyLimit(RequestBodyLimits.CollectionContentBytes);
            collections.MapPut("/{id}", CollectionEndpoints.UpdateAsync).WithBodyLimit(RequestBodyLimits.CollectionContentBytes);
            collections.MapDelete("/{id}", CollectionEndpoints.DeleteAsync);
            collections.MapPut("/{id}/sharing", CollectionEndpoints.SetSharingAsync);
            collections.MapDelete("/{id}/images/{sha256}", CollectionEndpoints.RemoveImageAsync);
        }

        private static async Task<IResult> UploadImageAsync(
            string sha256,
            HttpContext context,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string? key = CollectionEndpoints.KeyOf(context);

            // The key's shape looked at first, the body read only then (code review, 2026-11-27, Security: anybody made the service
            // read and hold up to 2 MB a request before it answered). The flow refuses a missing key as it always does.
            byte[] bytes = [];

            if (CollectionRules.IsOwnerKey(key))
            {
                using var body = new MemoryStream();
                await context.Request.Body.CopyToAsync(body, cancellationToken);
                bytes = body.ToArray();
            }

            return CollectionEndpoints.Result(await CollectionFlows.UploadImageAsync(
                key,
                CollectionEndpoints.ConnectionOf(context),
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
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.CreateAsync(
                CollectionEndpoints.KeyOf(context),
                CollectionEndpoints.ConnectionOf(context),
                content,
                store,
                new CollectionFiles(options.CollectionsRoot!),
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));

        private static async Task<IResult> UpdateAsync(
            string id,
            CollectionContentRequest? content,
            HttpContext context,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.UpdateAsync(
                CollectionEndpoints.KeyOf(context),
                id,
                content,
                store,
                new CollectionFiles(options.CollectionsRoot!),
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));

        private static async Task<IResult> DeleteAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.DeleteAsync(
                CollectionEndpoints.KeyOf(context),
                id,
                store,
                DateTimeOffset.UtcNow,
                cancellationToken));

        // "Delete image" in CRT - the capture off the page at once, nothing else changed.
        private static async Task<IResult> RemoveImageAsync(
            string id,
            string sha256,
            HttpContext context,
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionEndpoints.Result(await CollectionFlows.RemoveImageAsync(
                CollectionEndpoints.KeyOf(context),
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
            ICollectionStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string? key = CollectionEndpoints.KeyOf(context);

            if (CollectionRules.IsOwnerKey(key) && request is null)
                return CollectionEndpoints.Result(CollectionOutcome.Refused(CollectionVerdict.BadRequest, "Say whether it is shared."));

            return CollectionEndpoints.Result(await CollectionFlows.SetSharedAsync(
                key,
                id,
                request?.Shared ?? false,
                store,
                AccountPages.BaseUrl(options.PublicApiBaseUrl),
                DateTimeOffset.UtcNow,
                cancellationToken));
        }

        // The collection's key, as the request carries it - null for none.
        private static string? KeyOf(HttpContext context) =>
            context.Request.Headers[CollectionRoutes.OwnerKeyHeader].ToString() is { Length: > 0 } key ? key : null;

        // The internet connection the request came from, as the day's limits count it (OscilloscopeModelFlows.ConnectionOf -
        // an IPv4 address, an IPv6 /64).
        private static string? ConnectionOf(HttpContext context) =>
            OscilloscopeModelFlows.ConnectionOf(context.Connection.RemoteIpAddress?.ToString());

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
