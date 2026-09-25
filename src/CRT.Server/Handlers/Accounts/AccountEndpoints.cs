using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Maps the account flows onto HTTP. THIS FILE IS A RIM AND NOTHING ELSE: each handler reads
    // the request into plain values, calls one AccountFlows method, and turns the verdict into a
    // status code. Not one decision is made here.
    //
    // That is the rule the csproj header states, and the reason for it is visible in this file's
    // test coverage: AccountFlows has 52 tests, this has none, because there is nothing here that
    // could be wrong in an interesting way. A rule that migrated into a handler body would be a
    // rule no test could reach.
    //
    // THE STATUS CODES ARE PART OF THE SECURITY DESIGN, not presentation:
    //
    //   Registration and password reset ALWAYS answer 202 Accepted, whether or not the address is
    //   known. A 409 Conflict for a taken address - the obvious REST-shaped answer - is an account
    //   enumeration oracle. Anything that makes these two respond differently reopens it.
    //
    //   Login answers 401 for every failure with no reason attached, for the same reason.
    //
    //   Token redemption DOES distinguish its failures, because the token is the only subject and
    //   the user needs different things from an expired link and a mistyped one.
    // ###########################################################################################
    public static class AccountEndpoints
    {
        public static void MapAccountEndpoints(this WebApplication app)
        {
            ArgumentNullException.ThrowIfNull(app);

            RouteGroupBuilder accounts = app.MapGroup("/api/accounts");

            accounts.MapPost("/register", AccountEndpoints.RegisterAsync);
            accounts.MapGet("/verify", AccountEndpoints.VerifyAsync);
            accounts.MapPost("/login", AccountEndpoints.LoginAsync);
            accounts.MapPost("/refresh", AccountEndpoints.RefreshAsync);
            accounts.MapPost("/logout", AccountEndpoints.LogoutAsync);
            accounts.MapGet("/me", AccountEndpoints.MeAsync);
            accounts.MapPost("/forgot-password", AccountEndpoints.ForgotPasswordAsync);
            accounts.MapPost("/reset-password", AccountEndpoints.ResetPasswordAsync);
        }

        // ###########################################################################################
        // POST /api/accounts/register
        //
        // 202 for both a new address and one already registered - see the header.
        // ###########################################################################################
        private static async Task<IResult> RegisterAsync(
            RegisterBody body,
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(
                    body.Email,
                    body.Password,
                    body.DisplayName,
                    AccountEndpoints.ClientAddress(context)),
                store,
                mailer,
                hasher,
                options,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (!outcome.IsAccepted)
                return Results.BadRequest(new { errors = outcome.Failures });

            return Results.Accepted(value: new
            {
                message = "If that address can receive mail, a message is on its way to it."
            });
        }

        // ###########################################################################################
        // GET /api/accounts/verify?token=...
        //
        // A GET because it is reached by clicking a link in a mail client. That makes it
        // unprotectable against a prefetching mail client or scanner following the link, which is
        // why verification is not a security boundary on its own - it proves the address receives
        // mail, nothing more.
        // ###########################################################################################
        private static async Task<IResult> VerifyAsync(
            string? token,
            IAccountStore store,
            CancellationToken cancellationToken)
        {
            TokenRedemptionOutcome outcome =
                await AccountFlows.VerifyEmailAsync(token, store, DateTimeOffset.UtcNow, cancellationToken);

            return outcome switch
            {
                TokenRedemptionOutcome.Success =>
                    Results.Ok(new { message = "Your address is confirmed. You can close this page." }),

                TokenRedemptionOutcome.AlreadyUsed =>
                    Results.Ok(new { message = "That link has already been used. Your address is confirmed." }),

                TokenRedemptionOutcome.Expired =>
                    Results.BadRequest(new { message = "That link has expired. Ask for a new one from the app." }),

                _ => Results.BadRequest(new { message = "That link is not valid. Check you copied all of it." })
            };
        }

        // ###########################################################################################
        // POST /api/accounts/login
        //
        // 401 for every failure, 429 when rate limited. The 429 carries Retry-After so a client can
        // wait rather than hammer.
        // ###########################################################################################
        private static async Task<IResult> LoginAsync(
            LoginBody body,
            HttpContext context,
            IAccountStore store,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest(
                    body.Email,
                    body.Password,
                    AccountEndpoints.UserAgent(context),
                    AccountEndpoints.ClientAddress(context)),
                store,
                hasher,
                options,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (outcome.IsRateLimited)
            {
                context.Response.Headers.RetryAfter =
                    ((int)Math.Ceiling(outcome.RetryAfter.TotalSeconds))
                    .ToString(System.Globalization.CultureInfo.InvariantCulture);

                return Results.StatusCode(StatusCodes.Status429TooManyRequests);
            }

            if (!outcome.IsSuccess)
                return Results.Unauthorized();

            return Results.Ok(AccountEndpoints.SessionResponse(outcome.Account!, outcome.Session!));
        }

        private static async Task<IResult> RefreshAsync(
            RefreshBody body,
            HttpContext context,
            IAccountStore store,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            RefreshOutcome outcome = await AccountFlows.RefreshAsync(
                body.RefreshToken,
                AccountEndpoints.UserAgent(context),
                AccountEndpoints.ClientAddress(context),
                store,
                options,
                DateTimeOffset.UtcNow,
                cancellationToken);

            if (!outcome.IsSuccess)
                return Results.Unauthorized();

            return Results.Ok(AccountEndpoints.SessionResponse(outcome.Account!, outcome.Session!));
        }

        // ###########################################################################################
        // POST /api/accounts/logout
        //
        // Always 204, even for a token that was never valid - see AccountFlows.LogoutAsync.
        // ###########################################################################################
        private static async Task<IResult> LogoutAsync(
            RefreshBody body,
            IAccountStore store,
            CancellationToken cancellationToken)
        {
            await AccountFlows.LogoutAsync(body.RefreshToken, store, DateTimeOffset.UtcNow, cancellationToken);

            return Results.NoContent();
        }

        // ###########################################################################################
        // GET /api/accounts/me
        //
        // The first authenticated endpoint, and the shape every later one follows: pull the bearer
        // token, hand it to AuthenticateAsync, 401 if it does not resolve.
        //
        // NOTE WHAT IS NOT RETURNED: no password hash, no session list, no internal ids beyond the
        // account's own. A response shaped by "what does the client need" rather than "what does
        // the record hold".
        // ###########################################################################################
        private static async Task<IResult> MeAsync(
            HttpContext context,
            IAccountStore store,
            CancellationToken cancellationToken)
        {
            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                AccountEndpoints.BearerToken(context), store, DateTimeOffset.UtcNow, cancellationToken);

            if (account is null)
                return Results.Unauthorized();

            // The systems this account reviews (Phase 6 roles). Empty for an administrator, who is
            // in every pool by definition rather than by rows.
            IReadOnlySet<string> reviewerOf = await store.GetReviewedSystemIdsAsync(account.Id, cancellationToken);

            return Results.Ok(new
            {
                id = account.Id,
                email = account.Email,
                displayName = account.DisplayName,
                isVerified = account.IsVerified,
                isAdministrator = account.IsAdministrator,
                reviewerOf = reviewerOf.OrderBy(id => id, StringComparer.Ordinal).ToList(),
                createdUtc = account.CreatedUtc,
                lastLoginUtc = account.LastLoginUtc
            });
        }

        // ###########################################################################################
        // POST /api/accounts/forgot-password
        //
        // ALWAYS 202, whether or not the address is known - the same rule as registration.
        // ###########################################################################################
        private static async Task<IResult> ForgotPasswordAsync(
            ForgotPasswordBody body,
            HttpContext context,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            CancellationToken cancellationToken)
        {
            await AccountFlows.RequestPasswordResetAsync(
                body.Email,
                store,
                mailer,
                options,
                DateTimeOffset.UtcNow,
                AccountEndpoints.ClientAddress(context),
                cancellationToken);

            return Results.Accepted(value: new
            {
                message = "If that address has an account, a reset link is on its way to it."
            });
        }

        private static async Task<IResult> ResetPasswordAsync(
            ResetPasswordBody body,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            CancellationToken cancellationToken)
        {
            TokenRedemptionOutcome outcome = await AccountFlows.CompletePasswordResetAsync(
                body.Token, body.NewPassword, store, mailer, hasher, DateTimeOffset.UtcNow, cancellationToken);

            return outcome switch
            {
                TokenRedemptionOutcome.Success =>
                    Results.Ok(new { message = "Your password is set. You can sign in with it now." }),

                TokenRedemptionOutcome.PasswordRejected =>
                    Results.BadRequest(new { errors = AccountRules.ValidatePassword(body.NewPassword) }),

                TokenRedemptionOutcome.Expired =>
                    Results.BadRequest(new { message = "That link has expired. Ask for a new one." }),

                TokenRedemptionOutcome.AlreadyUsed =>
                    Results.BadRequest(new { message = "That link has already been used. Ask for a new one." }),

                _ => Results.BadRequest(new { message = "That link is not valid." })
            };
        }

        // -------------------------------------------------------------------------------------
        // Request and response shapes.
        // -------------------------------------------------------------------------------------

        private static object SessionResponse(AccountRecord account, IssuedSession session)
        {
            return new
            {
                refreshToken = session.RefreshToken,
                expiresUtc = session.ExpiresUtc,
                account = new
                {
                    id = account.Id,
                    email = account.Email,
                    displayName = account.DisplayName,
                    isVerified = account.IsVerified
                }
            };
        }

        // ###########################################################################################
        // The client's address, for rate limiting.
        //
        // UseForwardedHeaders has already run and is restricted to loopback proxies (see
        // Program.cs), so this is the real client address rather than Apache's - without that,
        // every per-IP limit would be one global bucket, a control that looks present and is not.
        // ###########################################################################################
        private static string? ClientAddress(HttpContext context)
        {
            return context.Connection.RemoteIpAddress?.ToString();
        }

        private static string? UserAgent(HttpContext context)
        {
            return context.Request.Headers.UserAgent.ToString() is { Length: > 0 } agent ? agent : null;
        }

        // Pulls the token out of "Authorization: Bearer <token>". Anything else yields null, which
        // the caller turns into a 401.
        private static string? BearerToken(HttpContext context)
        {
            string header = context.Request.Headers.Authorization.ToString();

            const string prefix = "Bearer ";

            if (!header.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                return null;

            string token = header[prefix.Length..].Trim();

            return token.Length == 0 ? null : token;
        }

        // Bodies are records so the JSON binder can construct them, and are deliberately all
        // nullable strings: validation is AccountRules' job, and a binder that rejected a missing
        // field would answer with a framework-shaped error rather than our own message.
        public sealed record RegisterBody(string? Email, string? Password, string? DisplayName);

        public sealed record LoginBody(string? Email, string? Password);

        public sealed record RefreshBody(string? RefreshToken);

        public sealed record ForgotPasswordBody(string? Email);

        public sealed record ResetPasswordBody(string? Token, string? NewPassword);
    }
}
