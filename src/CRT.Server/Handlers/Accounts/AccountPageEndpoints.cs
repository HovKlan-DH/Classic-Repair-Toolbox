using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;
using Microsoft.Extensions.Primitives;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // THE ACCOUNT WEB PAGES ON HTTP (2026-11-25) - a RIM, as AccountEndpoints is: each reads the form into plain values,
    // calls the very flow the JSON route calls (AccountFlows), and hands the verdict to AccountPages, which makes the page.
    // Not one decision is made here; the anti-enumeration answers, the rate limits and the mail budget are the flows'.
    //
    //   GET  /account/                 what an account is, and the ways in
    //   GET  /account/register         the form          POST: AccountFlows.RegisterAsync
    //   GET  /account/verify?token=    the mailed link   (GET /api/accounts/verify answers the same)
    //   GET  /account/forgot-password  ask for a code and set a password with it
    //   POST /account/forgot-password  AccountFlows.RequestPasswordResetAsync
    //   POST /account/reset-password   AccountFlows.CompletePasswordResetAsync
    //
    // *** NO ANTI-FORGERY TOKEN, ON PURPOSE. *** A form another site posts here can do nothing its author could not do by
    // posting it directly: making an account mails the address named, asking for a code mails the address named (the same
    // neutral answer, the same mail budget), and setting a password needs the code from the mailbox. Nothing here acts on a
    // signed-in person - there is no sign-in on these pages - so there is nothing for a forged form to ride on. Read by hand
    // from the request (ReadFormAsync), never through [FromForm], which would ask for the framework's anti-forgery middleware.
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

            pages.MapGet("/", AccountPageEndpoints.Landing);
            pages.MapGet("/register", AccountPageEndpoints.RegisterForm);
            pages.MapPost("/register", AccountPageEndpoints.RegisterAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapGet("/verify", AccountPageEndpoints.VerifyAsync);
            pages.MapGet("/forgot-password", AccountPageEndpoints.ForgotPasswordForm);
            pages.MapPost("/forgot-password", AccountPageEndpoints.ForgotPasswordAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
            pages.MapPost("/reset-password", AccountPageEndpoints.ResetPasswordAsync).WithBodyLimit(RequestBodyLimits.FormBytes);
        }

        // ###########################################################################################
        // A page, sent with the headers every page needs: the policy allowing nothing but the page's own style
        // (AccountPages.ContentSecurityPolicy), no sniffing a type, no Referer leaving with a verify link's token, no frame on
        // another site, and no copy kept.
        // ###########################################################################################
        internal static IResult Page(HttpContext context, string html, int status = StatusCodes.Status200OK)
        {
            IHeaderDictionary headers = context.Response.Headers;

            headers.ContentSecurityPolicy = AccountPages.ContentSecurityPolicy;
            headers.XContentTypeOptions = "nosniff";
            headers.XFrameOptions = "DENY";
            headers["Referrer-Policy"] = "no-referrer";
            headers.CacheControl = "no-store";

            return Results.Content(html, "text/html; charset=utf-8", System.Text.Encoding.UTF8, status);
        }

        private static IResult Landing(HttpContext context, ServerOptions options) =>
            AccountPageEndpoints.Page(context, AccountPages.Landing(AccountPageEndpoints.BaseUrl(options)));

        private static IResult RegisterForm(HttpContext context, ServerOptions options) =>
            AccountPageEndpoints.Page(context, AccountPages.RegisterForm(AccountPageEndpoints.BaseUrl(options)));

        private static async Task<IResult> RegisterAsync(
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            string baseUrl = AccountPageEndpoints.BaseUrl(options);
            IFormCollection? form = await AccountPageEndpoints.FormOf(context, cancellationToken);

            string? displayName = AccountPageEndpoints.Field(form, "displayName");
            string? email = AccountPageEndpoints.Field(form, "email");

            // The field no person sees, filled: answered as a made account would be, and nothing made or mailed.
            if (!string.IsNullOrEmpty(AccountPageEndpoints.Field(form, AccountPages.TrapField)))
                return AccountPageEndpoints.Page(context, AccountPages.Registered(baseUrl, email ?? string.Empty));

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(email, AccountPageEndpoints.Field(form, "password"), displayName, context.Connection.RemoteIpAddress?.ToString()),
                store,
                mailer,
                hasher,
                options,
                DateTimeOffset.UtcNow,
                cancellationToken);

            return outcome.IsAccepted
                ? AccountPageEndpoints.Page(context, AccountPages.Registered(baseUrl, email?.Trim() ?? string.Empty))
                : AccountPageEndpoints.Page(context, AccountPages.RegisterForm(baseUrl, displayName, email, outcome.Failures), StatusCodes.Status400BadRequest);
        }

        private static async Task<IResult> VerifyAsync(string? token, HttpContext context, IAccountStore store, ServerOptions options, CancellationToken cancellationToken)
        {
            TokenRedemptionOutcome outcome = await AccountFlows.VerifyEmailAsync(token, store, DateTimeOffset.UtcNow, cancellationToken);

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
        private static async Task<IFormCollection?> FormOf(HttpContext context, CancellationToken cancellationToken)
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
        private static string? Field(IFormCollection? form, string name) =>
            form is not null && form.TryGetValue(name, out StringValues values) && values.Count == 1 ? values[0] : null;
    }
}
