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
    //   GET  /collections/{id}                             the page - 404 for no such collection (or one whose owner's account
    //                                                      is locked), 410 for one deleted, saying when. ?edit=<c> and
    //                                                      ?delete=<c> show a comment being edited or asked about.
    //   GET  /collections/{id}/{sha256}.png                a capture - only one the collection shows, while it is shown
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

        public static void MapCollectionPageEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder pages = app.MapGroup(CollectionRoutes.PagePath);

            pages.MapGet("/{id}", CollectionPageEndpoints.PageAsync);
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
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            await CollectionPageEndpoints.RenderAsync(
                id, context, store, accounts, options, cancellationToken,
                view => view with { EditingId = edit, DeletingId = delete });

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
            CollectionBaselines baselines = CollectionBaselines.For(options);

            if (!options.IsCollectionSharingConfigured)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotSetUp(siteBaseUrl), baselines, signedIn: false, StatusCodes.Status503ServiceUnavailable);

            CollectionRecord? collection = await CollectionFlows.FindShownAsync(id, store, accounts, cancellationToken, evenIfDeleted: true);

            if (collection is null)
                return CollectionPageEndpoints.Page(context, CollectionPages.NotFound(siteBaseUrl), baselines, signedIn: false, StatusCodes.Status404NotFound);

            if (collection.IsDeleted)
                return CollectionPageEndpoints.Page(context, CollectionPages.Deleted(siteBaseUrl, collection), baselines, signedIn: false, StatusCodes.Status410Gone);

            (AccountRecord? viewer, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            var view = new CollectionPageView
            {
                Viewer = viewer is null ? null : new CollectionViewer(viewer.Id, viewer.DisplayName, viewer.IsAdministrator),
                Csrf = token is null ? null : WebSessions.Csrf(token),
                Comments = CollectionCommentFlows.ShownTo(await store.ListCommentsAsync(collection.Id, cancellationToken), collection, viewer)
            };

            if (adjust is not null)
                view = adjust(view);

            return CollectionPageEndpoints.Page(context, CollectionPages.Page(siteBaseUrl, collection, baselines.UrlOf, view), baselines, viewer is not null, status);
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

            CollectionRecord? collection = await CollectionFlows.FindShownAsync(id, store, accounts, cancellationToken);

            if (collection is null || !collection.Images.Any(image => image.Entry.Sha256 == sha256))
                return Results.NotFound();

            Stream? bytes = new CollectionFiles(options.CollectionsRoot!).OpenRead(sha256);

            if (bytes is null)
                return Results.NotFound();

            context.Response.Headers.XContentTypeOptions = "nosniff";
            context.Response.Headers.CacheControl = "public, max-age=3600";

            return Results.Stream(bytes, CollectionPng.ContentType);
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

                CommentOutcome outcome = await CollectionCommentFlows.EditAsync(account, commentId, body, store, accounts, DateTimeOffset.UtcNow, cancellationToken);

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
                (await CollectionCommentFlows.DeleteAsync(account, commentId, store, DateTimeOffset.UtcNow, cancellationToken), null));

        private static Task<IResult> HideCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, _) =>
                (await CollectionCommentFlows.SetHiddenAsync(account, commentId, hidden: true, store, DateTimeOffset.UtcNow, cancellationToken), null));

        private static Task<IResult> ShowCommentAsync(
            string id,
            long commentId,
            HttpContext context,
            ICollectionStore store,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken) =>
            CollectionPageEndpoints.FormAsync(id, context, store, accounts, options, cancellationToken, async (account, _) =>
                (await CollectionCommentFlows.SetHiddenAsync(account, commentId, hidden: false, store, DateTimeOffset.UtcNow, cancellationToken), null));

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
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            string siteBaseUrl = AccountPages.BaseUrl(options.PublicApiBaseUrl);
            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            (AccountRecord? account, string? token) = await CollectionPageEndpoints.ReaderAsync(context, accounts, options, cancellationToken);

            if (account is null || token is null || !WebSessions.CsrfMatches(token, AccountPageEndpoints.Field(form, "csrf")))
            {
                string signIn = AccountPages.Link(siteBaseUrl, "sign-in") + "?return=" + Uri.EscapeDataString(CollectionPages.PagePath(siteBaseUrl, id));
                return AccountPageEndpoints.SeeOther(context, signIn);
            }

            (CommentOutcome outcome, Func<CollectionPageView, CollectionPageView>? refusedView) = await act(account, form);

            if (outcome.IsDone)
                return AccountPageEndpoints.SeeOther(context, $"{CollectionPages.PagePath(siteBaseUrl, id)}#comment-{outcome.CommentId}");

            return await CollectionPageEndpoints.RenderAsync(
                id, context, store, accounts, options, cancellationToken,
                refusedView ?? (view => view),
                CollectionEndpoints.StatusOf(outcome.Verdict));
        }

        // The reader signed in on the website, by the cookie - the session slid as every signed-in request slides it.
        private static async Task<(AccountRecord? Account, string? Token)> ReaderAsync(
            HttpContext context,
            IAccountStore accounts,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string? token = WebSessions.TokenOf(context.Request, AccountPages.BaseUrl(options.PublicApiBaseUrl));

            if (token is null)
                return (null, null);

            AccountRecord? account = await AccountFlows.AuthenticateAsync(token, accounts, DateTimeOffset.UtcNow, cancellationToken, options);

            return account is null ? (null, null) : (account, token);
        }

        // A page signed in is the reader's own: never kept by anybody between them and the service.
        private static IResult Page(HttpContext context, string html, CollectionBaselines baselines, bool signedIn, int status = StatusCodes.Status200OK)
        {
            IHeaderDictionary headers = context.Response.Headers;

            headers.ContentSecurityPolicy = CollectionPages.ContentSecurityPolicy(baselines.Origin is { } origin ? [origin] : []);
            headers.XContentTypeOptions = "nosniff";
            headers.XFrameOptions = "DENY";
            headers["Referrer-Policy"] = AccountPages.ReferrerPolicy;
            headers["X-Robots-Tag"] = CollectionPageEndpoints.RobotsTag;
            headers.CacheControl = signedIn ? "private, no-store" : "no-cache";

            return Results.Content(html, "text/html; charset=utf-8", System.Text.Encoding.UTF8, status);
        }
    }

    // ###########################################################################################
    // WHERE A CAPTURE'S BASELINE PICTURE IS SERVED: the published data CRT reads - the stable tree when publishing to it is set
    // up, else the BETA tree - under its stored name (DataTreeLayout.StoredFilePath, by checksum and the board's name's
    // extension). Only a file that is there is linked, so the page never shows a broken picture.
    // ###########################################################################################
    public sealed class CollectionBaselines(string? treeRoot, string? publicBaseUrl)
    {
        public static CollectionBaselines For(ServerOptions options) =>
            options.IsProductionPublishingConfigured
                ? new CollectionBaselines(options.ProductionDataTreeRoot, options.ProductionPublicDataBaseUrl)
                : new CollectionBaselines(options.DataTreeRoot, options.PublicDataBaseUrl);

        // The origin the pictures come from, for the page's Content-Security-Policy.
        public string? Origin =>
            Uri.TryCreate(publicBaseUrl, UriKind.Absolute, out Uri? uri) ? uri.GetLeftPart(UriPartial.Authority) : null;

        public string? UrlOf(CollectionBaseline baseline)
        {
            if (string.IsNullOrWhiteSpace(treeRoot) || string.IsNullOrWhiteSpace(publicBaseUrl) || !DataTreeLayout.IsSha256(baseline.Sha256))
                return null;

            string stored = DataTreeLayout.StoredFilePath(baseline.Sha256!, baseline.FileName ?? ".png");

            return File.Exists(Path.Combine(treeRoot, stored)) ? $"{publicBaseUrl.TrimEnd('/')}/{stored}" : null;
        }
    }
}
