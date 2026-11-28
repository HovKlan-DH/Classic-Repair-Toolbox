using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // A SHARED COLLECTION'S PAGE ON HTTP (2026-11-26) - a rim over the flows and CollectionPages; Apache forwards
    // "/collections/" to the service, as it forwards "/account/" (INSTALLING.md).
    //
    //   GET  /collections/                                 the administrator's list of every collection (CollectionListFlows);
    //                                                      signed out, the way in
    //   GET  /collections/{id}                             the page - 404 for no such collection, 410 for one deleted, saying
    //                                                      when. ?remove=1 the administrator's question before deleting it.
    //                                                      Its owner stopped sharing it: 404 saying it is not shared at the
    //                                                      moment - but to the administrator, who sees it, said so.
    //   GET  /collections/{id}/{sha256}.png                a capture - only one the collection shows, while it is shown (to
    //                                                      this reader)
    //   GET  /collections/{id}/download                    every capture in one ZIP ("Download all", 15.1.0) - as its pictures
    //   POST /collections/{id}/delete                      the administrator deletes the collection (step 4)
    //
    // Read by anybody with the link. The website's ONE signed-in reader is the administrator (WebSessions - the cookie, and the
    // token every form carries; owner decision, 2026-11-28, as the accounts for everybody went - the comments and the account
    // locks went with them), and a form from another website is refused before anything is read (AccountPages
    // .IsFromAnotherSite). Nothing here is counted as API usage or refused as an outdated CRT (ClientArea.Web). Every answer
    // asks search engines to keep nothing (X-Robots-Tag).
    // ###########################################################################################
    public static class CollectionPageEndpoints
    {
        public const string RobotsTag = "noindex, nofollow, noarchive";

        // ###########################################################################################
        // THE PAGES' LIMIT PER ADDRESS (code review, 2026-11-27): every picture of a page is a request that reads its collection,
        // and anybody with a link could ask for them without end. Far above a person reading - a 300-capture page is some 300
        // requests, so this is dozens of such pages an hour - and a wall to a script.
        // ###########################################################################################
        public const string RateLimitPolicy = "collection-pages";

        public const int PermitsPerAddressPerHour = 20_000;

        public static void AddCollectionPageRateLimit(this IServiceCollection services) =>
            services.AddPerAddressRateLimit(CollectionPageEndpoints.RateLimitPolicy, CollectionPageEndpoints.PermitsPerAddressPerHour);

        public static void MapCollectionPageEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder pages = app.MapGroup(CollectionRoutes.PagePath).RequireRateLimiting(CollectionPageEndpoints.RateLimitPolicy);

            pages.MapGet("/", CollectionPageEndpoints.ListAsync);
            pages.MapGet("/{id}", CollectionPageEndpoints.PageAsync);
            pages.MapPost("/{id}/delete", CollectionPageEndpoints.DeleteCollectionAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapGet("/{id}/download", CollectionPageEndpoints.DownloadAsync);
            pages.MapGet("/{id}/{file}", CollectionPageEndpoints.ImageAsync);
        }

        private static async Task<IResult> PageAsync(
            string id,
            string? remove,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            await CollectionPageEndpoints.RenderAsync(
                id, context, store, accounts, options, cancellationToken,
                view => view with { RemovingCollection = remove == "1" });

        // ###########################################################################################
        // GET /collections/ - the administrator's list of every collection, or, signed out, the way in.
        // ###########################################################################################
        private static async Task<IResult> ListAsync(
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);

            if (!options.IsCollectionSharingConfigured)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotSetUp(), signedIn: false, StatusCodes.Status503ServiceUnavailable);

            (AccountRecord? reader, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            if (reader is null || token is null ||
                await CollectionListFlows.ListForAsync(reader, store, DateTimeOffset.UtcNow, cancellationToken) is not { } all)
            {
                return CollectionPageEndpoints.Page(context, CollectionPages.ListSignedOut(siteBaseUrl), signedIn: reader is not null);
            }

            return CollectionPageEndpoints.Page(
                context,
                CollectionPages.List(siteBaseUrl, new CollectionViewer(reader.Id, reader.DisplayName, reader.IsAdministrator), WebSessions.Csrf(token), all),
                signedIn: true);
        }

        // ###########################################################################################
        // POST /collections/{id}/delete - the administrator takes a collection down (step 4; CollectionPages.RemoveCollection),
        // by CollectionFlows.DeleteFromWebAsync - the administrator's alone. Refused from another website, then the session and
        // the form's token; anybody else signed in is refused (403) with the page as it was. Done: on to the list.
        // ###########################################################################################
        private static async Task<IResult> DeleteCollectionAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);

            // Sharing not set up, no form acts - as the pages and every /api/collections route answer 503 then.
            if (!options.IsCollectionSharingConfigured)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotSetUp(), signedIn: false, StatusCodes.Status503ServiceUnavailable);

            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            (AccountRecord? account, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            if (account is null || token is null || !WebSessions.CsrfMatches(token, AccountPageEndpoints.Field(form, "csrf")))
                return AccountPageEndpoints.SeeOther(context, CollectionPageEndpoints.SignInBackTo(siteBaseUrl, CollectionPages.PagePath(siteBaseUrl, id)));

            CollectionOutcome outcome = await CollectionFlows.DeleteFromWebAsync(account, id, store, DateTimeOffset.UtcNow, cancellationToken);

            return outcome.Verdict == CollectionVerdict.Removed
                ? AccountPageEndpoints.SeeOther(context, CollectionPages.ListPath(siteBaseUrl))
                : await CollectionPageEndpoints.RenderAsync(
                    id, context, store, accounts, options, cancellationToken,
                    view => view with { PageErrors = outcome.Errors },
                    status: CollectionEndpoints.StatusOf(outcome.Verdict));
        }

        // The sign-in page, coming back to `path` (WebSessions.ReturnPath decides it may).
        private static string SignInBackTo(string siteBaseUrl, string path) =>
            AccountPages.Link(siteBaseUrl, "sign-in") + "?return=" + Uri.EscapeDataString(path);

        // ###########################################################################################
        // The page as this reader sees it - `adjust` adds what they are doing (a delete asked, a refusal).
        // ###########################################################################################
        private static async Task<IResult> RenderAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken,
            Func<CollectionPageView, CollectionPageView>? adjust = null,
            int status = StatusCodes.Status200OK)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);

            if (!options.IsCollectionSharingConfigured)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotSetUp(), signedIn: false, StatusCodes.Status503ServiceUnavailable);

            // The reader is read at most once - for the page below, or first when the collection is offline for everybody else.
            (AccountRecord? Account, string? Token)? reader = null;

            async Task<(AccountRecord? Account, string? Token)> ReadReaderAsync(CancellationToken token) =>
                reader ??= await CollectionPageEndpoints.ReaderAsync(context, accounts, options, token);

            ShownCollection? shown = await CollectionFlows.FindShownToAsync(
                id, async token => (await ReadReaderAsync(token)).Account, store, DateTimeOffset.UtcNow, cancellationToken, evenIfDeleted: true);

            if (shown is null)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotFound(), signedIn: false, StatusCodes.Status404NotFound);

            CollectionRecord collection = shown.Collection;

            if (collection.IsDeleted)
                return CollectionPageEndpoints.Page(context, CollectionPages.Deleted(collection), signedIn: false, StatusCodes.Status410Gone);

            // Its owner stopped sharing it, and this reader is not the administrator: only that it is not shared at the moment -
            // its link comes back once it is shared again (migration 0028). Never kept by a cache: it may be back any moment.
            if (shown.IsWithheld)
            {
                IResult withheld = CollectionPageEndpoints.Page(context, CollectionPages.SharingStopped(collection), signedIn: false, StatusCodes.Status404NotFound);
                context.Response.Headers.CacheControl = "no-store";
                return withheld;
            }

            (AccountRecord? viewer, string? csrfToken) = await ReadReaderAsync(cancellationToken);

            var view = new CollectionPageView
            {
                Viewer = viewer is null ? null : new CollectionViewer(viewer.Id, viewer.DisplayName, viewer.IsAdministrator),
                Csrf = csrfToken is null ? null : WebSessions.Csrf(csrfToken),
                Offline = shown.Offline
            };

            if (adjust is not null)
                view = adjust(view);

            return CollectionPageEndpoints.Page(context, CollectionPages.Page(siteBaseUrl, collection, view), viewer is not null, status);
        }

        private static async Task<IResult> ImageAsync(
            string id,
            string file,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            context.Response.Headers["X-Robots-Tag"] = CollectionPageEndpoints.RobotsTag;

            string sha256 = file.EndsWith(".png", StringComparison.Ordinal) ? file[..^4] : string.Empty;

            if (!options.IsCollectionSharingConfigured || !DataTreeLayout.IsSha256(sha256))
                return Results.NotFound();

            // The administrator sees an offline collection's pictures as its page (FindShownToAsync) - the session read only then.
            ShownCollection? shown = await CollectionFlows.FindShownImageToAsync(
                id,
                sha256,
                async token => (await CollectionPageEndpoints.ReaderAsync(context, accounts, options, token)).Account,
                store,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (shown is null)
                return Results.NotFound();

            Stream? bytes = new CollectionFiles(options.CollectionsRoot!).OpenRead(sha256);

            if (bytes is null)
                return Results.NotFound();

            context.Response.Headers.XContentTypeOptions = "nosniff";

            // One shown to the administrator alone is kept by nobody between them and the service. One shown to all may be
            // kept, but is ASKED FOR AGAIN at every view (code review, 2026-11-27: for an hour a picture of a collection taken
            // down was still served from caches) - its checksum the ETag, so a picture unchanged costs a 304, never its bytes.
            context.Response.Headers.CacheControl = shown.IsOffline ? "private, no-store" : "public, no-cache";

            return Results.Stream(bytes, CollectionPng.ContentType, entityTag: new Microsoft.Net.Http.Headers.EntityTagHeaderValue($"\"{sha256}\""));
        }

        // ###########################################################################################
        // "DOWNLOAD ALL" (owner request, 2026-11-27, after the owner's capture gallery): every capture of a collection shown to this
        // reader, in one ZIP, named as the viewer's "Download" names each (CollectionDownload says how it is written).
        // ###########################################################################################
        private static async Task<IResult> DownloadAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            context.Response.Headers["X-Robots-Tag"] = CollectionPageEndpoints.RobotsTag;

            if (!options.IsCollectionSharingConfigured)
                return Results.NotFound();

            ShownCollection? shown = await CollectionFlows.FindShownToAsync(
                id,
                async token => (await CollectionPageEndpoints.ReaderAsync(context, accounts, options, token)).Account,
                store,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (shown is null || shown.IsWithheld || shown.Collection.Images.Count == 0)
                return Results.NotFound();

            var files = new CollectionFiles(options.CollectionsRoot!);
            IReadOnlyList<CollectionDownload.Entry> entries = CollectionDownload.EntriesOf(shown.Collection);

            context.Response.Headers.XContentTypeOptions = "nosniff";
            context.Response.Headers.CacheControl = "private, no-store";

            return Results.Stream(
                body => CollectionDownload.WriteAsync(body, entries, files.OpenRead, cancellationToken),
                CollectionDownload.ContentType,
                CollectionDownload.FileNameOf(shown.Collection));
        }

        // The reader signed in on the website, by the cookie - the session slid as every signed-in request slides it, and the
        // cookie written again to last as long (code review, 2026-11-27: set at sign-in alone, the browser dropped it after 30
        // days however often it was used).
        private static async Task<(AccountRecord? Account, string? Token)> ReaderAsync(
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);
            string? token = WebSessions.TokenOf(context.Request, siteBaseUrl);

            if (token is null)
                return (null, null);

            DateTimeOffset now = DateTimeOffset.UtcNow;
            AccountRecord? account = await AccountFlows.AuthenticateAsync(token, accounts, now, cancellationToken, options);

            if (account is null)
                return (null, null);

            if (!context.Response.HasStarted)
                WebSessions.Keep(context.Response, siteBaseUrl, token, now.AddDays(options.RefreshTokenDays));

            return (account, token);
        }

        // A page signed in is the reader's own: never kept by anybody between them and the service.
        private static IResult Page(HttpContext context, string html, bool signedIn, int status = StatusCodes.Status200OK)
        {
            IHeaderDictionary headers = context.Response.Headers;

            headers.ContentSecurityPolicy = CollectionPages.ContentSecurityPolicy();
            headers.XContentTypeOptions = "nosniff";
            headers.XFrameOptions = "DENY";
            headers["Referrer-Policy"] = AccountPages.ReferrerPolicy;
            headers["X-Robots-Tag"] = CollectionPageEndpoints.RobotsTag;
            // A page that writes the session's cookie is the reader's own however it came out - a 404 for a signed-in reader too
            // (code review, 2026-11-27, Security: "no-cache" let a shared cache keep the token).
            headers.CacheControl = signedIn || headers.SetCookie.Count > 0 ? "private, no-store" : "no-cache";

            return Results.Content(html, "text/html; charset=utf-8", System.Text.Encoding.UTF8, status);
        }
    }
}
