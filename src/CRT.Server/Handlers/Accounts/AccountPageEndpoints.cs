using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;
using Microsoft.Extensions.Primitives;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // THE ACCOUNT WEB PAGES ON HTTP (2026-11-25) - a RIM, as AccountEndpoints is: each reads the form into plain values,
    // calls the very flow the JSON route calls (AccountFlows), and hands the verdict to AccountPages, which makes the page.
    // Not one decision is made here but one - the sign-in's administrators alone; the anti-enumeration answers, the rate limits
    // and the mail budget are the flows'. Nobody makes an account here since server 17.0.0 (owner decision, 2026-11-28).
    //
    //   GET  /account/verify?token=    the mailed link   (GET /api/accounts/verify answers the same)
    //   GET  /account/forgot-password  ask for a code and set a password with it
    //   POST /account/forgot-password  AccountFlows.RequestPasswordResetAsync
    //   POST /account/reset-password   AccountFlows.CompletePasswordResetAsync
    //   GET  /account/sign-in          the website's sign-in - the administrator's   POST: AccountFlows.LoginAsync, the cookie kept
    //   POST /account/sign-out         the session ended, the cookie forgotten - only with the session's form token
    //
    // *** A FORM ANOTHER WEBSITE POSTS IS REFUSED *** (AccountPages.IsFromAnotherSite, before anything is read) - every form
    // here and on the collections' pages. A forged form would ride on the visitor's own IP address: asking for a code mails
    // somebody, and the mail budget is counted per IP address, so a hidden form on a busy page would spend every visitor's
    // (code review, 2026-11-25, Security). WHAT ACTS ON A SIGNED-IN PERSON NEEDS THE SESSION'S TOKEN TOO (WebSessions.Csrf): the
    // sign-out here, and the administrator's delete on a collection's page - signing in needs none, as it acts on no session.
    // Read by hand from the request (ReadFormAsync), never through [FromForm], which would ask for the framework's
    // anti-forgery middleware.
    //
    // A form's body is small, and limited where it is mapped (RequestBodyLimits.FormBytes). The pages are classified as the
    // web's, never refused as an outdated CRT (ClientVersionPolicy, ClientArea.Web) - a browser sends no CRT version.
    // ###########################################################################################
    public static class AccountPageEndpoints
    {
        public static void MapAccountPageEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder pages = app.MapGroup(AccountPages.Path);

            pages.MapGet("/verify", AccountPageEndpoints.VerifyAsync);
            pages.MapGet("/forgot-password", AccountPageEndpoints.ForgotPasswordForm);
            pages.MapPost("/forgot-password", AccountPageEndpoints.ForgotPasswordAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/reset-password", AccountPageEndpoints.ResetPasswordAsync).WithBodyLimit(RequestBodyLimits.FormBytes);

            // Signed in on the website - the administrator, for every shared collection (WebSessions).
            pages.MapGet("/sign-in", AccountPageEndpoints.SignInForm);
            pages.MapPost("/sign-in", AccountPageEndpoints.SignInAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/sign-out", AccountPageEndpoints.SignOutAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
        }

        private static IResult SignInForm(string? @return, HttpContext context, ServerOptions options)
        {
            string baseUrl = AccountPageEndpoints.BaseUrl(options);

            return AccountPageEndpoints.Page(context, AccountPages.SignIn(baseUrl, returnTo: WebSessions.ReturnPath(@return, baseUrl)));
        }

        // ###########################################################################################
        // POST /account/sign-in - AccountFlows.LoginAsync, the very sign-in CRT's is, its session kept in the browser's cookie;
        // then back to the page it came from, or "You are signed in". A wrong address or password says no reason (AccountFlows).
        // *** THE ADMINISTRATOR'S ALONE (owner decision, 2026-11-28): *** right ones of any other account are refused in words,
        // and the session the sign-in made is ended at once - no cookie is ever written for it.
        // ###########################################################################################
        private static async Task<IResult> SignInAsync(
            HttpContext context,
            IAccountStore store,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            string baseUrl = AccountPageEndpoints.BaseUrl(options);
            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            string? email = AccountPageEndpoints.Field(form, "email");
            string? returnTo = WebSessions.ReturnPath(AccountPageEndpoints.Field(form, "return"), baseUrl);

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest(email, AccountPageEndpoints.Field(form, "password"), context.Request.Headers.UserAgent.ToString(), context.Connection.RemoteIpAddress?.ToString()),
                store,
                hasher,
                options,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (!outcome.IsSuccess || outcome.Session is null || outcome.Account is null)
            {
                return AccountPageEndpoints.Page(
                    context,
                    AccountPages.SignIn(baseUrl, email, returnTo, [outcome.IsRateLimited ? AccountPages.SignInSlowDown : AccountPages.SignInRefused]),
                    outcome.IsRateLimited ? StatusCodes.Status429TooManyRequests : StatusCodes.Status401Unauthorized);
            }

            if (!outcome.Account.IsAdministrator)
            {
                await AccountFlows.LogoutAsync(outcome.Session.RefreshToken, store, DateTimeOffset.UtcNow, CancellationToken.None);

                return AccountPageEndpoints.Page(
                    context,
                    AccountPages.SignIn(baseUrl, email, returnTo, [AccountPages.SignInAdministratorsOnly]),
                    StatusCodes.Status403Forbidden);
            }

            // A form from another site never gets this far (RefuseFromAnotherSite).
            await WebSessions.EndEarlierAsync(
                WebSessions.TokenOf(context.Request, baseUrl), outcome.Session.RefreshToken, store, DateTimeOffset.UtcNow, cancellationToken);

            WebSessions.Keep(context.Response, baseUrl, outcome.Session.RefreshToken, outcome.Session.ExpiresUtc);

            return returnTo is not null
                ? AccountPageEndpoints.SeeOther(context, returnTo)
                : AccountPageEndpoints.Page(context, AccountPages.SignedIn(baseUrl, outcome.Account.DisplayName));
        }

        // ###########################################################################################
        // POST /account/sign-out - the session ended on the server (AccountFlows.LogoutAsync) and the cookie forgotten. The
        // form carries its token, so no other site can sign a visitor out.
        // ###########################################################################################
        private static async Task<IResult> SignOutAsync(
            HttpContext context,
            IAccountStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            string baseUrl = AccountPageEndpoints.BaseUrl(options);
            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            string? token = WebSessions.TokenOf(context.Request, baseUrl);

            // The session is ended on the server only with its form's token - no other site ends it. But THIS BROWSER forgets
            // its cookie whatever the token says (code review, 2026-11-27, Security: an old page's form, its token of an
            // earlier session, said "signed out" and left the browser signed in - on a shared computer, for the next person).
            // A form from another site never gets this far (RefuseFromAnotherSite).
            if (token is not null && WebSessions.CsrfMatches(token, AccountPageEndpoints.Field(form, "csrf")))
                await AccountFlows.LogoutAsync(token, store, DateTimeOffset.UtcNow, cancellationToken);

            if (token is not null)
                WebSessions.Forget(context.Response, baseUrl);

            string? returnTo = WebSessions.ReturnPath(AccountPageEndpoints.Field(form, "return"), baseUrl);

            return returnTo is not null
                ? AccountPageEndpoints.SeeOther(context, returnTo)
                : AccountPageEndpoints.Page(context, AccountPages.SignedOut(baseUrl));
        }

        // A form done: on to `path` (a GET, so a reload posts nothing again).
        internal static IResult SeeOther(HttpContext context, string path)
        {
            context.Response.Headers.Location = path;
            context.Response.Headers.CacheControl = "no-store";

            return Results.StatusCode(StatusCodes.Status303SeeOther);
        }

        // ###########################################################################################
        // A page, sent with the headers every page needs: the policy allowing nothing but the page's own style
        // (AccountPages.ContentSecurityPolicy), no sniffing a type, no Referer leaving with a verify link's token
        // (AccountPages.ReferrerPolicy), no frame on another site, and no copy kept.
        // ###########################################################################################
        internal static IResult Page(HttpContext context, string html, int status = StatusCodes.Status200OK)
        {
            IHeaderDictionary headers = context.Response.Headers;

            headers.ContentSecurityPolicy = AccountPages.ContentSecurityPolicy;
            headers.XContentTypeOptions = "nosniff";
            headers.XFrameOptions = "DENY";
            headers["Referrer-Policy"] = AccountPages.ReferrerPolicy;
            headers.CacheControl = "no-store";

            return Results.Content(html, "text/html; charset=utf-8", System.Text.Encoding.UTF8, status);
        }

        // A form posted from another website: 403 and a page saying so, nothing read and nothing done (the header).
        internal static IResult? RefuseFromAnotherSite(HttpContext context, ServerOptions options)
        {
            IHeaderDictionary headers = context.Request.Headers;
            string baseUrl = AccountPageEndpoints.BaseUrl(options);

            return AccountPages.IsFromAnotherSite(headers["Sec-Fetch-Site"].ToString(), headers.Origin.ToString(), baseUrl)
                ? AccountPageEndpoints.Page(context, AccountPages.FromAnotherSite(baseUrl), StatusCodes.Status403Forbidden)
                : null;
        }

        private static async Task<IResult> VerifyAsync(
            string? token,
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            TokenRedemptionOutcome outcome = await AccountFlows.VerifyEmailAsync(token, store, mailer, DateTimeOffset.UtcNow, cancellationToken);

            (string html, int status) = AccountPages.Verified(AccountPageEndpoints.BaseUrl(options), outcome);

            return AccountPageEndpoints.Page(context, html, status);
        }

        private static IResult ForgotPasswordForm(HttpContext context, ServerOptions options) =>
            AccountPageEndpoints.Page(context, AccountPages.ForgotPassword(AccountPageEndpoints.BaseUrl(options)));

        // The same neutral answer for every address - a code may or may not be on its way (AccountFlows' header).
        private static async Task<IResult> ForgotPasswordAsync(
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            string email = AccountPageEndpoints.Field(form, "email")?.Trim() ?? string.Empty;

            if (email.Length > 0)
            {
                await AccountFlows.RequestPasswordResetAsync(
                    email, store, mailer, options, DateTimeOffset.UtcNow, context.Connection.RemoteIpAddress?.ToString(), cancellationToken);
            }

            return AccountPageEndpoints.Page(context, AccountPages.ForgotPassword(AccountPageEndpoints.BaseUrl(options), email, codeAsked: email.Length > 0));
        }

        private static async Task<IResult> ResetPasswordAsync(
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            if (AccountPageEndpoints.RefuseFromAnotherSite(context, options) is IResult refused)
                return refused;

            string baseUrl = AccountPageEndpoints.BaseUrl(options);
            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);
            string? newPassword = AccountPageEndpoints.Field(form, "newPassword");

            TokenRedemptionOutcome outcome = await AccountFlows.CompletePasswordResetAsync(
                AccountPageEndpoints.Field(form, "code")?.Trim(), newPassword, store, mailer, hasher, DateTimeOffset.UtcNow, cancellationToken);

            if (outcome == TokenRedemptionOutcome.Success)
                return AccountPageEndpoints.Page(context, AccountPages.PasswordSet(baseUrl));

            // The reset route's own sentences (AccountEndpoints.ResetPasswordAsync), so the two answer alike.
            IReadOnlyList<string> errors = outcome switch
            {
                TokenRedemptionOutcome.PasswordRejected => AccountRules.ValidatePassword(newPassword),
                TokenRedemptionOutcome.Expired => ["That code has expired. Ask for a new one."],
                TokenRedemptionOutcome.AlreadyUsed => ["That code has already been used. Ask for a new one."],
                _ => ["That code is not valid. Check that the whole code was pasted."]
            };

            return AccountPageEndpoints.Page(context, AccountPages.ForgotPassword(baseUrl, resetErrors: errors), StatusCodes.Status400BadRequest);
        }

        private static string BaseUrl(ServerOptions options) => AccountPages.BaseUrl(options.PublicApiBaseUrl);

        // The posted form, or none for a request that sent no form or one that cannot be read - a field of none is null,
        // which the flows refuse in their own words. A body over the route's limit is the framework's 413, as everywhere.
        internal static async Task<IFormCollection?> FormOf(HttpContext context, CancellationToken cancellationToken)
        {
            if (!context.Request.HasFormContentType)
                return null;

            try
            {
                return await context.Request.ReadFormAsync(cancellationToken);
            }
            catch (InvalidDataException)
            {
                return null;
            }
        }

        // One value of a field; a field sent twice is no field (a form posts each once).
        internal static string? Field(IFormCollection? form, string name) =>
            form is not null && form.TryGetValue(name, out StringValues values) && values.Count == 1 ? values[0] : null;
    }
}
