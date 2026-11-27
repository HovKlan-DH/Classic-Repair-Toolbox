using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Collections
{
    // ###########################################################################################
    // A SHARED COLLECTION'S PAGE ON HTTP (2026-11-26) - a rim over the flows and CollectionPages; Apache forwards
    // "/collections/" to the service, as it forwards "/account/" (INSTALLING.md).
    //
    //   GET  /collections/                                 "Your collections" (step 4) - CollectionListFlows; signed out, the
    //                                                      way in. ?lock=<account> the administrator's question before
    //                                                      locking an account.
    //   POST /collections/accounts/{a}/lock|unlock         the administrator locks or unlocks an account (AccountLockFlows)
    //   GET  /collections/{id}                             the page - 404 for no such collection (or one whose owner's account
    //                                                      is locked or gone - but to the administrator, who is told it is
    //                                                      offline), 410 for one deleted, saying when. ?edit=<c> and
    //                                                      ?delete=<c> show a comment being edited or asked about, ?remove=1
    //                                                      the administrator's question before deleting the collection.
    //                                                      Its owner stopped sharing it: 404 saying it is not shared at the
    //                                                      moment - but to its owner and the administrator, who see it, said so.
    //   GET  /collections/{id}/{sha256}.png                a capture - only one the collection shows, while it is shown (to
    //                                                      this reader)
    //   GET  /collections/{id}/download                    every capture in one ZIP ("Download all", 15.1.0) - as its pictures
    //   POST /collections/{id}/delete                      the administrator deletes the collection (step 4)
    //   POST /collections/{id}/comments                    a comment (step 3) - CollectionCommentFlows.AddAsync
    //   POST /collections/{id}/comments/{c}/edit|delete|hide|show
    //
    // Read by anybody with the link; the comment forms need the reader signed in on the website (WebSessions - the cookie, and
    // the token every form carries), and a form from another website is refused before anything is read (AccountPages
    // .IsFromAnotherSite). A form done goes back to the page (303); one refused shows the page again with why and what was
    // typed. Nothing here is counted as API usage or refused as an outdated CRT (ClientArea.Web). Every answer asks search
    // engines to keep nothing (X-Robots-Tag).
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
            pages.MapPost("/accounts/{accountId:long}/lock", CollectionPageEndpoints.LockAccountAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/accounts/{accountId:long}/unlock", CollectionPageEndpoints.UnlockAccountAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapGet("/{id}", CollectionPageEndpoints.PageAsync);
            pages.MapPost("/{id}/delete", CollectionPageEndpoints.DeleteCollectionAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapGet("/{id}/download", CollectionPageEndpoints.DownloadAsync);
            pages.MapGet("/{id}/{file}", CollectionPageEndpoints.ImageAsync);
            pages.MapPost("/{id}/comments", CollectionPageEndpoints.AddCommentAsync).WithBodyLimit(RequestBodyLimits.CommentFormBytes);
            pages.MapPost("/{id}/comments/{commentId:long}/edit", CollectionPageEndpoints.EditCommentAsync).WithBodyLimit(RequestBodyLimits.CommentFormBytes);
            pages.MapPost("/{id}/comments/{commentId:long}/delete", CollectionPageEndpoints.DeleteCommentAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/{id}/comments/{commentId:long}/hide", CollectionPageEndpoints.HideCommentAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/{id}/comments/{commentId:long}/show", CollectionPageEndpoints.ShowCommentAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
        }

        private static async Task<IResult> PageAsync(
            string id,
            long? edit,
            long? delete,
            string? remove,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            await CollectionPageEndpoints.RenderAsync(
                id, context, store, accounts, options, cancellationToken,
                view => view with { EditingId = edit, DeletingId = delete, RemovingCollection = remove == "1" });

        // ###########################################################################################
        // GET /collections/ - "Your collections" (step 4): the reader's lists (CollectionListFlows), or, signed out, the way in.
        // ?lock=<account> asks the administrator before locking that account (AccountLockFlows).
        // ###########################################################################################
        private static Task<IResult> ListAsync(
            long? @lock,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.RenderListAsync(context, store, accounts, options, cancellationToken, @lock);

        // The list as this reader sees it - `locking` the account asked about, `problem` what a lock or unlock refused said.
        private static async Task<IResult> RenderListAsync(
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken,
            long? locking = null,
            string? problem = null,
            int status = StatusCodes.Status200OK)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);

            if (!options.IsCollectionSharingConfigured)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotSetUp(), signedIn: false, StatusCodes.Status503ServiceUnavailable);

            (AccountRecord? reader, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            if (reader is null || token is null)
                return CollectionPageEndpoints.Page(context, CollectionPages.ListSignedOut(siteBaseUrl), signedIn: false);

            CollectionListView list = await CollectionListFlows.ListForAsync(reader, store, accounts, DateTimeOffset.UtcNow, cancellationToken);
            var viewer = new CollectionViewer(reader.Id, reader.DisplayName, reader.IsAdministrator);

            // Only an account the administrator may lock is asked about (AccountLockFlows decides).
            LockedAccount? asked = await AccountLockFlows.ToAskAboutLockingAsync(reader, locking, accounts, cancellationToken) is { } account
                ? new LockedAccount(account.Id, account.DisplayName, account.Email)
                : null;

            return CollectionPageEndpoints.Page(
                context,
                CollectionPages.List(siteBaseUrl, viewer, WebSessions.Csrf(token), list, asked, problem),
                signedIn: true,
                status);
        }

        // ###########################################################################################
        // POST /collections/accounts/{id}/lock and .../unlock - the administrator locks or unlocks an account (owner decision,
        // 2026-11-27; AccountLockFlows decides). Every form's three steps first (SignedInFormAsync). Done: back to the list,
        // at "Locked accounts"; refused: the list again, saying why.
        // ###########################################################################################
        private static Task<IResult> LockAccountAsync(
            long accountId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            ILoggerFactory loggers,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.SetLockedAsync(
                accountId, locked: true, context, store, accounts, options, loggers.CreateLogger(typeof(AccountLockFlows).FullName!), cancellationToken);

        private static Task<IResult> UnlockAccountAsync(
            long accountId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            ILoggerFactory loggers,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.SetLockedAsync(
                accountId, locked: false, context, store, accounts, options, loggers.CreateLogger(typeof(AccountLockFlows).FullName!), cancellationToken);

        private static async Task<IResult> SetLockedAsync(
            long accountId,
            bool locked,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            ILogger logger,
            CancellationToken cancellationToken)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);
            SignedInForm signedIn = await CollectionPageEndpoints.SignedInFormAsync(context, accounts, options, CollectionPages.ListPath(siteBaseUrl), cancellationToken);

            if (signedIn.Instead is { } instead)
                return instead;

            AccountLockOutcome outcome = await AccountLockFlows.SetLockedAsync(signedIn.Account, accountId, locked, accounts, DateTimeOffset.UtcNow, cancellationToken, logger);

            if (outcome.IsDone)
                return AccountPageEndpoints.SeeOther(context, $"{CollectionPages.ListPath(siteBaseUrl)}#locked");

            int status = outcome.Verdict switch
            {
                AccountLockVerdict.NotFound => StatusCodes.Status404NotFound,
                AccountLockVerdict.Refused => StatusCodes.Status409Conflict,
                _ => StatusCodes.Status403Forbidden
            };

            return await CollectionPageEndpoints.RenderListAsync(context, store, accounts, options, cancellationToken, problem: outcome.Error, status: status);
        }

        // ###########################################################################################
        // POST /collections/{id}/delete - the administrator takes a collection down (step 4; CollectionPages.RemoveCollection),
        // by CollectionFlows.DeleteFromWebAsync - the administrator's alone, then the route CRT's own delete takes. Refused from
        // another website, then the session and the form's token, as every form here is; anybody else signed in is refused (403)
        // with the page as it was. Done: on to "Your collections".
        // ###########################################################################################
        private static async Task<IResult> DeleteCollectionAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);
            SignedInForm signedIn = await CollectionPageEndpoints.SignedInFormAsync(context, accounts, options, CollectionPages.PagePath(siteBaseUrl, id), cancellationToken);

            if (signedIn.Instead is { } instead)
                return instead;

            CollectionOutcome outcome = await CollectionFlows.DeleteFromWebAsync(signedIn.Account, id, store, DateTimeOffset.UtcNow, cancellationToken);

            return outcome.Verdict == CollectionVerdict.Removed
                ? AccountPageEndpoints.SeeOther(context, CollectionPages.ListPath(siteBaseUrl))
                : await CollectionPageEndpoints.RenderAsync(id, context, store, accounts, options, cancellationToken, status: CollectionEndpoints.StatusOf(outcome.Verdict));
        }

        // The sign-in page, coming back to `path` (WebSessions.ReturnPath decides it may).
        private static string SignInBackTo(string siteBaseUrl, string path) =>
            AccountPages.Link(siteBaseUrl, "sign-in") + "?return=" + Uri.EscapeDataString(path);

        // ###########################################################################################
        // The page as this reader sees it - `adjust` adds what they are doing (an edit, a delete asked, a refusal).
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
                id, async token => (await ReadReaderAsync(token)).Account, store, accounts, DateTimeOffset.UtcNow, cancellationToken, evenIfDeleted: true);

            if (shown is null)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotFound(), signedIn: false, StatusCodes.Status404NotFound);

            CollectionRecord collection = shown.Collection;

            if (collection.IsDeleted)
                return CollectionPageEndpoints.Page(context, CollectionPages.Deleted(collection), signedIn: false, StatusCodes.Status410Gone);

            // Its owner stopped sharing it, and this reader is neither its owner nor the administrator: only that it is not shared
            // at the moment - its link comes back once it is shared again (migration 0028). Never kept by a cache: it may be back
            // any moment.
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
                Comments = CollectionCommentFlows.ShownTo(await store.ListCommentsAsync(collection.Id, cancellationToken), collection, viewer),
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
                accounts,
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
                accounts,
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

        // ---- The comment forms (step 3) ------------------------------------------------------------

        private static Task<IResult> AddCommentAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, form) =>
            {
                string? image = AccountPageEndpoints.Field(form, "image");
                string? body = AccountPageEndpoints.Field(form, "body");

                CommentOutcome outcome = await CollectionCommentFlows.AddAsync(
                    account, id, image, body, store, accounts, mailer, AccountPages.BaseUrl(options.PublicApiBaseUrl), DateTimeOffset.UtcNow, cancellationToken);

                return (outcome, view => view with { ErrorOn = image, Errors = outcome.Errors, Typed = body });
            });

        private static Task<IResult> EditCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, form) =>
            {
                string? body = AccountPageEndpoints.Field(form, "body");

                CommentOutcome outcome = await CollectionCommentFlows.EditAsync(account, id, commentId, body, store, accounts, DateTimeOffset.UtcNow, cancellationToken);

                return (outcome, view => view with
                {
                    EditingId = commentId,
                    ErrorOn = commentId.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    Errors = outcome.Errors,
                    Typed = body
                });
            });

        private static Task<IResult> DeleteCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, _) =>
                (await CollectionCommentFlows.DeleteAsync(account, id, commentId, store, DateTimeOffset.UtcNow, cancellationToken), null));

        private static Task<IResult> HideCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, _) =>
                (await CollectionCommentFlows.SetHiddenAsync(account, id, commentId, hidden: true, store, DateTimeOffset.UtcNow, cancellationToken), null));

        private static Task<IResult> ShowCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, _) =>
                (await CollectionCommentFlows.SetHiddenAsync(account, id, commentId, hidden: false, store, DateTimeOffset.UtcNow, cancellationToken), null));

        // ###########################################################################################
        // EVERY COMMENT FORM, ONE WAY: refused from another website, then the reader's session and the form's token - both must
        // be there and agree, or nothing is done and the page asks to sign in - then the flow. Done: back to the page, at the
        // comment. Refused: the page again with why (`refusedView`), or "not found" as the flow says.
        // ###########################################################################################
        private static async Task<IResult> FormAsync(
            string id,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken,
            Func<AccountRecord, IFormCollection?, Task<(CommentOutcome Outcome, Func<CollectionPageView, CollectionPageView>? RefusedView)>> act)
        {
            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);
            SignedInForm signedIn = await CollectionPageEndpoints.SignedInFormAsync(context, accounts, options, CollectionPages.PagePath(siteBaseUrl, id), cancellationToken);

            if (signedIn.Instead is { } instead)
                return instead;

            (CommentOutcome outcome, Func<CollectionPageView, CollectionPageView>? refusedView) = await act(signedIn.Account!, signedIn.Form);

            if (outcome.IsDone)
                return AccountPageEndpoints.SeeOther(context, $"{CollectionPages.PagePath(siteBaseUrl, id)}#comment-{outcome.CommentId}");

            // The refusal's words where its form can show them (a comment typed, an edit), else at the top of the page - a hide,
            // a delete, or a comment gone meanwhile, whose form is no longer drawn.
            bool saidAtTheTop = refusedView is null || outcome.Verdict == CollectionVerdict.NotFound;

            return await CollectionPageEndpoints.RenderAsync(
                id, context, store, accounts, options, cancellationToken,
                view => (refusedView ?? (shown => shown))(view) with { PageErrors = saidAtTheTop ? outcome.Errors : [] },
                CollectionEndpoints.StatusOf(outcome.Verdict));
        }

        // ###########################################################################################
        // EVERY FORM HERE, THE SAME THREE STEPS: refused from another website before anything is read (AccountPages
        // .IsFromAnotherSite), then the reader's session and the form's token - both there and agreeing (WebSessions.Csrf) - or
        // nothing is done and the reader goes to sign in, coming back to `returnPath`. `Instead` is the answer when a step
        // refused; else the account and the form.
        // ###########################################################################################
        private static async Task<SignedInForm> SignedInFormAsync(
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            string returnPath,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return new SignedInForm(null, null, refused);

            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);

            // Sharing not set up, no form acts - as the pages and every /api/collections route answer 503 then (code review,
            // 2026-11-27: a page left open still added a comment, mailed its owner, and landed on "not available").
            if (!options.IsCollectionSharingConfigured)
            {
                return new SignedInForm(null, null, CollectionPageEndpoints.Page(
                    context, CollectionPages.NotSetUp(), signedIn: false, StatusCodes.Status503ServiceUnavailable));
            }

            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            (AccountRecord? account, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            return account is null || token is null || !WebSessions.CsrfMatches(token, AccountPageEndpoints.Field(form, "csrf"))
                ? new SignedInForm(null, form, AccountPageEndpoints.SeeOther(context, CollectionPageEndpoints.SignInBackTo(siteBaseUrl, returnPath)))
                : new SignedInForm(account, form, null);
        }

        private sealed record SignedInForm(AccountRecord? Account, IFormCollection? Form, IResult? Instead);

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
