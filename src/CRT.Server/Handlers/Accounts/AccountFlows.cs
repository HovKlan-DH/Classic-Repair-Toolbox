using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Verification, login, refresh, logout and password reset. NO REGISTRATION since server 17.0.0
    // (owner decision, 2026-11-28: "I do not want the possibility for everyone to create an
    // account") - an account is made by accepting an invitation (MaintainerInvitationFlows), and
    // POST /api/accounts/register only says so.
    //
    // THIS IS WHERE THE DECISIONS LIVE. The endpoints in Program.cs read a request into plain
    // values, call one of these, and map the verdict onto a status code - nothing more. Everything
    // that could be wrong in an interesting way is here, and every method takes IAccountStore and
    // IEmailSender as arguments, so the whole file is testable against fakes with no database and
    // no network.
    //
    // FOUR PROPERTIES THIS FILE EXISTS TO GUARANTEE:
    //
    // 1. A RESET NEVER REVEALS WHETHER AN ADDRESS IS REGISTERED. It answers the same neutral
    //    result whatever it finds. What differs is whether a mail is sent, which only the mailbox
    //    owner sees. Any change that makes the RESULT depend on whether the account existed
    //    reopens enumeration.
    //
    //    *** AND AN ACCOUNT NOT CONFIRMED SAYS NOTHING EITHER *** (project owner, 2026-11-25: "I do
    //    not think any address should be revealed, no"). While anybody could register (server
    //    13.0.0 to 16.1.0), an account made was told apart by what followed: signing in with the
    //    password just given, and a name asked for being held. So an address not confirmed cannot
    //    sign in (LoginAsync, answered as a wrong password), and only a confirmed account holds a
    //    display name (ConfirmAddressAsync, IAccountStore.IsDisplayNameTakenAsync) - kept for the
    //    accounts left unconfirmed from then.
    //
    // 2. A LOGIN THAT FAILS TELLS YOU NOTHING ABOUT WHY. Unknown address, wrong password and
    //    locked account are one answer, given after the same Argon2 work. "No such account" -
    //    in words or in a quicker reply - would otherwise be a free account enumeration oracle
    //    on the one endpoint everybody can reach.
    //
    // 3. THE RATE LIMITER RUNS BEFORE THE HASHER. Argon2 allocates ~128 MiB per verification, so
    //    checking the password first and the limit second turns the limiter into decoration and
    //    leaves the memory exhaustion vector wide open.
    //
    // 4. REFRESH TOKENS ROTATE, AND REUSE REVOKES EVERYTHING. A token presented after it has
    //    already been rotated means two parties hold it. The chain is revoked rather than guessing
    //    which party is legitimate.
    // ###########################################################################################
    public static class AccountFlows
    {
        // Short, because a reset link is a live credential for an account whose owner may not be
        // the person who requested it.
        public static readonly TimeSpan PasswordResetTokenLifetime = TimeSpan.FromHours(AccountWording.PasswordResetCodeHours);

        // ###########################################################################################
        // VERIFY an email address from the link in the verification mail. Nothing mails such a link since server 17.0.0 - an
        // account is made by invitation alone (owner decision, 2026-11-28: no accounts for everybody) - but a link mailed
        // before still answers, and GET /api/accounts/verify stays mapped for it (an installed CRT 4.0's surface).
        //
        // Unlike a reset request, this DOES distinguish its failures - an expired link and an invalid
        // one need different things from the user ("I forgot my password", whose code confirms the
        // address too; check you copied it all). Neither answer reveals whether any account exists,
        // because the token is the only subject.
        // ###########################################################################################
        public static async Task<TokenRedemptionOutcome> VerifyEmailAsync(
            string? token,
            IAccountStore store,
            IEmailSender mailer,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);

            if (string.IsNullOrWhiteSpace(token))
                return TokenRedemptionOutcome.Invalid;

            AccountTokenRecord? record = await store.FindTokenByHashAsync(SecureToken.Hash(token), cancellationToken);

            if (record is null || record.Purpose != TokenPurpose.EmailVerification)
                return TokenRedemptionOutcome.Invalid;

            // A consumed token is not reusable. Without this, a verification link in a mailbox
            // stays live forever.
            if (record.ConsumedUtc is not null)
                return TokenRedemptionOutcome.AlreadyUsed;

            if (record.ExpiresUtc <= now)
                return TokenRedemptionOutcome.Expired;

            await store.ConsumeTokenAsync(record.Id, now, cancellationToken);
            await AccountFlows.ConfirmAddressAsync(record.AccountId, store, mailer, now, cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(record.AccountId, $"account:{record.AccountId}", "verify.email",
                    $"account:{record.AccountId}", null, now),
                cancellationToken);

            return TokenRedemptionOutcome.Success;
        }

        // ###########################################################################################
        // *** AN ADDRESS CONFIRMED - AND ITS DISPLAY NAME CLAIMED *** (project owner, 2026-11-25: "I do not think any address
        // should be revealed, no"). A name is held only by an account whose address is confirmed (IAccountStore
        // .IsDisplayNameTakenAsync): while anybody could register (server 13.0.0 to 16.1.0), a name held from registration on
        // told whether the account was made - so whether the address had an account. So two accounts left unconfirmed from then
        // may hold one name, and the second to confirm finds it held: it is given the name with "-2" after it (or the next
        // number free, the whole within DisplayNames' 30) and told so by mail.
        //
        // Every way an address becomes confirmed comes here: the mailed link (VerifyEmailAsync), a new password set with a
        // mailed code (CompletePasswordResetAsync) and a new address's code (AccountSelfServiceFlows). An accepted invitation
        // makes an account confirmed from the start, its name asked of the confirmed accounts as it is made.
        // ###########################################################################################
        internal static async Task ConfirmAddressAsync(
            long accountId,
            IAccountStore store,
            IEmailSender mailer,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            await store.SetVerifiedAsync(accountId, cancellationToken);

            if (await store.FindByIdAsync(accountId, cancellationToken) is not AccountRecord account)
                return;

            string name = account.DisplayName.Trim();

            if (!await store.IsDisplayNameTakenAsync(name, exceptAccountId: account.Id, cancellationToken))
                return;

            for (int number = 2; number < 1000; number++)
            {
                string given = AccountFlows.NumberedName(name, number);

                if (await store.IsDisplayNameTakenAsync(given, exceptAccountId: account.Id, cancellationToken))
                    continue;

                await store.SetDisplayNameAsync(account.Id, given, cancellationToken);
                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, account.Email, AccountFlows.NameGivenAtConfirmationAction, $"account:{account.Id}",
                        $"\"{name}\" -> \"{given}\"", now),
                    cancellationToken);
                await mailer.SendAsync(EmailTemplates.DisplayNameTakenMeanwhile(account.Email, name, given), cancellationToken);
                return;
            }
        }

        // "Dennis" and 2 -> "Dennis-2"; a name too long for the number is cut short first, so the whole stays within the rule.
        internal static string NumberedName(string name, int number)
        {
            string suffix = "-" + number.ToString(System.Globalization.CultureInfo.InvariantCulture);
            string kept = name.Length + suffix.Length > DisplayNames.MaximumLength
                ? name[..Math.Max(0, DisplayNames.MaximumLength - suffix.Length)]
                : name;

            return kept + suffix;
        }

        // The audit actions of the two above - an unconfirmed address refused, a name given as it was confirmed.
        public const string SignInUnconfirmedAction = "login.unconfirmed";

        public const string NameGivenAtConfirmationAction = "account.name_given";

        // ###########################################################################################
        // LOG IN.
        //
        // THE ORDER OF OPERATIONS IN HERE IS THE SECURITY DESIGN. Read it before changing it:
        //
        //   1. Rate limit FIRST, before any hashing. Argon2 allocates ~128 MiB per verification.
        //   2. Find the account. If missing or locked, STILL record a failure, STILL run one
        //      Argon2 verification (Argon2PasswordHasher.VerifyWithoutAccountAsync), and answer the
        //      generic result - otherwise the failure count and the response time both leak which
        //      addresses exist. The time did leak, until the security review of 2026-10-11.
        //   3. Verify the password.
        //   4. An account whose address is NOT CONFIRMED is no sign-in - answered as a wrong password.
        //   5. Every failure answers the same thing.
        //
        // *** AN UNCONFIRMED ADDRESS CANNOT SIGN IN *** (project owner, 2026-11-25: "I do not think any
        // address should be revealed, no"). While accounts were open to everybody (server 13.0.0 to
        // 16.1.0), signing in was the second request that told whether an address had an account:
        // register it with a password of your own, then sign in with that password - it worked only
        // when no account had existed. So the right password of an unconfirmed account is a failure
        // like any other: recorded for the limiter (a limiter that tripped only for wrong passwords
        // would say the same thing), after the same Argon2 work, in the same words. An account left
        // unconfirmed from then confirms its address with "I forgot my password", whose code confirms
        // it too - nothing needs a session first.
        // ###########################################################################################
        public static async Task<LoginOutcome> LoginAsync(
            LoginRequest request,
            IAccountStore store,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(hasher);
            ArgumentNullException.ThrowIfNull(options);

            if (string.IsNullOrWhiteSpace(request.Email) || string.IsNullOrEmpty(request.Password))
                return LoginOutcome.Failed();

            string normalised = AccountRules.NormaliseEmail(request.Email);

            // ---- 1. Rate limit, BEFORE the hasher. ----
            DateTimeOffset accountSince = now - AuthRateLimitPolicy.AccountWindow;
            DateTimeOffset addressSince = now - AuthRateLimitPolicy.AddressWindow;

            IReadOnlyList<DateTimeOffset> accountFailures =
                await store.GetRecentAuthFailuresAsync(normalised, null, accountSince, cancellationToken);

            // This account from THIS address - the tight bucket (code review, 2026-10-13). With no
            // address known, the account's own failures stand in for it, the strict side.
            IReadOnlyList<DateTimeOffset> pairFailures = request.IpAddress is null
                ? accountFailures
                : await store.GetRecentAuthFailuresAsync(normalised, request.IpAddress, accountSince, cancellationToken);

            IReadOnlyList<DateTimeOffset> addressFailures = request.IpAddress is null
                ? Array.Empty<DateTimeOffset>()
                : await store.GetRecentAuthFailuresAsync(null, request.IpAddress, addressSince, cancellationToken);

            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(pairFailures, accountFailures, addressFailures, now);

            if (!verdict.IsAllowed)
                return LoginOutcome.RateLimited(verdict.RetryAfter);

            // ---- 2. Find the account. ----
            AccountRecord? account = await store.FindByNormalisedEmailAsync(normalised, cancellationToken);

            if (account is null)
            {
                // Record the failure even though there is no account. Without this, an attacker
                // can probe addresses indefinitely without ever tripping the limiter, and the
                // difference in behaviour is itself the oracle.
                //
                // And spend what checking a password costs: answering before Argon2 ran was
                // hundreds of milliseconds faster than a wrong password (security review,
                // 2026-10-11).
                await hasher.VerifyWithoutAccountAsync(request.Password, cancellationToken);
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);
                return LoginOutcome.Failed();
            }

            if (account.IsLocked)
            {
                await hasher.VerifyWithoutAccountAsync(request.Password, cancellationToken);
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);

                // The SAME answer as a wrong password, in the same time. A distinct "your account
                // is locked" reply - or a quicker one - confirms the address exists.
                return LoginOutcome.Failed();
            }

            // ---- 3. Verify. ----
            PasswordVerificationResult verification = await hasher.VerifyAsync(request.Password, account.PasswordHash, cancellationToken);

            if (!verification.IsValid)
            {
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);

                // *** THIS ACTION MUST NOT BE THE ONE THE LIMITER COUNTS. *** RecordAuthFailure
                // above is what the rate limiter reads, and in the MySQL store both it and this
                // call land in the SAME `audit` table - the limiter counts rows by their action
                // string. Writing "login.failed" here too made one wrong password count TWICE
                // against the per-IP bucket, halving 30 attempts to 15; and because
                // ClearAuthFailuresAsync only clears rows whose subject is the address, this
                // second row was never cleared by a successful login and accumulated against the
                // address for ever.
                //
                // So this row - which exists to name WHICH ACCOUNT was targeted, something the
                // limiter's own row deliberately does not record - carries its own action. (Since
                // 2026-10-13 the limiter counts its own table, rate_limit_events, so the coupling is
                // gone - the separate action is kept so the audit trail reads as before.)
                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, account.Email, "login.failed.account", $"account:{account.Id}",
                        request.IpAddress, now),
                    cancellationToken);

                return LoginOutcome.Failed();
            }

            // ---- 4. The address not confirmed: a failure like any other (the header). ----
            if (!account.IsVerified)
            {
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);
                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, account.Email, AccountFlows.SignInUnconfirmedAction, $"account:{account.Id}",
                        request.IpAddress, now),
                    cancellationToken);

                return LoginOutcome.Failed();
            }

            // ---- 5. Success. ----

            // The parameters were raised since this password was last hashed, and the plaintext is
            // legitimately in hand for this one instant - so upgrade it now.
            if (verification.NeedsRehash)
                await store.SetPasswordHashAsync(account.Id, await hasher.HashAsync(request.Password, cancellationToken), cancellationToken);

            // A success clears the count: someone who gets it right on the fourth try is not an
            // attacker and must not be locked out afterwards.
            await store.ClearAuthFailuresAsync(normalised, cancellationToken);
            await store.SetLastLoginAsync(account.Id, now, cancellationToken);

            IssuedSession session = await AccountFlows.IssueSessionAsync(
                account.Id, request.UserAgent, request.IpAddress, store, options, now, cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, "login", $"account:{account.Id}", request.IpAddress, now),
                cancellationToken);

            return LoginOutcome.Succeeded(account, session);
        }

        // ###########################################################################################
        // REFRESH a session.
        //
        // ROTATION WITH REUSE DETECTION. Every refresh issues a new token and retires the old one.
        // Presenting an already-rotated token means two parties hold it, so the entire chain is
        // revoked - the legitimate client is logged out too, which is the correct trade: an
        // inconvenient re-login beats leaving a thief with a live session.
        // ###########################################################################################
        public static async Task<RefreshOutcome> RefreshAsync(
            string? refreshToken,
            string? userAgent,
            string? ipAddress,
            IAccountStore store,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(options);

            if (string.IsNullOrWhiteSpace(refreshToken))
                return RefreshOutcome.Failed();

            SessionRecord? session =
                await store.FindSessionByHashAsync(SecureToken.Hash(refreshToken), cancellationToken);

            if (session is null)
                return RefreshOutcome.Failed();

            // REUSE DETECTION. This token was already exchanged for a newer one, so somebody is
            // presenting a copy.
            if (session.ReplacedById is not null)
            {
                await store.RevokeAllSessionsAsync(
                    session.AccountId, "refresh token reused", now, cancellationToken);

                await store.WriteAuditAsync(
                    new AuditEntry(
                        session.AccountId,
                        $"account:{session.AccountId}",
                        "session.reuse_detected",
                        $"session:{session.Id}",
                        $"A rotated refresh token was presented again from {ipAddress ?? "an unknown address"}. " +
                        "Every session for this account was revoked.",
                        now),
                    cancellationToken);

                return RefreshOutcome.Failed();
            }

            if (session.RevokedUtc is not null || session.ExpiresUtc <= now)
                return RefreshOutcome.Failed();

            AccountRecord? account = await store.FindByIdAsync(session.AccountId, cancellationToken);

            // An account locked or deleted since the session was issued must not be refreshable -
            // this is the check that makes "removing someone takes effect immediately" true, and
            // is exactly what a stateless JWT could not do.
            if (account is null || account.IsLocked)
            {
                await store.RevokeSessionAsync(session.Id, "account unavailable", now, cancellationToken);
                return RefreshOutcome.Failed();
            }

            IssuedSession issued = await AccountFlows.IssueSessionAsync(
                account.Id, userAgent, ipAddress, store, options, now, cancellationToken);

            await store.RotateSessionAsync(session.Id, issued.SessionId, now, cancellationToken);

            return RefreshOutcome.Succeeded(account, issued);
        }

        // ###########################################################################################
        // LOG OUT. Revokes the one session presented, never the whole account: logging out on a
        // laptop must not sign you out on the bench machine.
        //
        // An unknown or already-revoked token answers SUCCESS. There is nothing useful to say, the
        // caller's intent (not being logged in) is satisfied either way, and distinguishing them
        // would turn logout into a way to test whether a token is live.
        // ###########################################################################################
        public static async Task LogoutAsync(
            string? refreshToken,
            IAccountStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);

            if (string.IsNullOrWhiteSpace(refreshToken))
                return;

            SessionRecord? session =
                await store.FindSessionByHashAsync(SecureToken.Hash(refreshToken), cancellationToken);

            if (session is null || session.RevokedUtc is not null)
                return;

            await store.RevokeSessionAsync(session.Id, "signed out", now, cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(session.AccountId, $"account:{session.AccountId}", "logout",
                    $"session:{session.Id}", null, now),
                cancellationToken);
        }

        // ###########################################################################################
        // AUTHENTICATE a request carrying a token, for GET /api/accounts/me and everything after
        // it.
        //
        // ONE INDEXED SINGLE-ROW LOOKUP PER REQUEST, which is the cost of opaque tokens and is the
        // point: a revoked session stops working immediately rather than at its natural expiry.
        // ###########################################################################################
        public static async Task<AccountRecord?> AuthenticateAsync(
            string? accessToken,
            IAccountStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken = default,
            ServerOptions? options = null)
        {
            ArgumentNullException.ThrowIfNull(store);

            if (string.IsNullOrWhiteSpace(accessToken))
                return null;

            SessionRecord? session =
                await store.FindSessionByHashAsync(SecureToken.Hash(accessToken), cancellationToken);

            if (session is null || session.RevokedUtc is not null || session.ExpiresUtc <= now)
                return null;

            // A rotated session's token is no longer current, even before the old row expires.
            if (session.ReplacedById is not null)
                return null;

            AccountRecord? account = await store.FindByIdAsync(session.AccountId, cancellationToken);

            if (account is null || account.IsLocked)
                return null;

            // ###########################################################################################
            // SLIDING EXPIRY (owner request, 2026-09-22).
            //
            // *** THIS RUNS ONLY AFTER EVERY CHECK ABOVE HAS PASSED, AND THAT ORDERING IS THE
            // DESIGN. *** A revoked, rotated or expired session, or one whose account has since
            // been locked, has already returned null - so extension can never prolong a credential
            // that should have stopped working. Moving this earlier would make locking an account
            // stop taking effect immediately, which Phase 6's definition of done forbids.
            //
            // The store repeats these conditions in its own WHERE clause rather than trusting this
            // caller; see MySqlAccountStore.ExtendSessionAsync for why.
            //
            // *** A FAILURE HERE MUST NOT FAIL THE REQUEST. *** The caller is authenticated and
            // entitled to what they asked for; losing an expiry bump is a convenience, not a
            // correctness matter. The worst case is that the maintainer signs in again sooner than
            // they otherwise would - never that a valid request is refused because a bookkeeping
            // UPDATE could not be written.
            // ###########################################################################################
            if (options is not null)
            {
                TimeSpan lifetime = TimeSpan.FromDays(options.RefreshTokenDays);

                if (SessionExtensionRules.ShouldExtend(session.ExpiresUtc, now, lifetime))
                {
                    try
                    {
                        await store.ExtendSessionAsync(
                            session.Id,
                            SessionExtensionRules.NewExpiry(now, lifetime),
                            now,
                            cancellationToken);
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        await AccountFlows.TryWriteAuditAsync(
                            store,
                            new AuditEntry(
                                account.Id,
                                $"account:{account.Id}",
                                "session.extend_failed",
                                $"session:{session.Id}",
                                $"A session expiry could not be extended: {ex.Message}",
                                now),
                            cancellationToken);
                    }
                }
            }

            return account;
        }

        // ###########################################################################################
        // Writes an audit entry, swallowing any failure.
        //
        // Used only from paths where the audit IS the error handling - if recording that something
        // failed itself throws, there is nowhere left to report it, and letting it propagate would
        // turn a tolerated failure into the very request failure the catch existed to prevent.
        // ###########################################################################################
        private static async Task TryWriteAuditAsync(
            IAccountStore store,
            AuditEntry entry,
            CancellationToken cancellationToken)
        {
            try
            {
                await store.WriteAuditAsync(entry, cancellationToken);
            }
            catch
            {
                // Deliberately ignored - see the header.
            }
        }

        // ###########################################################################################
        // REQUEST A PASSWORD RESET.
        //
        // Answers the same neutral result whatever it finds. An unknown address gets NO MAIL AT ALL -
        // mailing a stranger "you have no account here" turns this endpoint into a way to send
        // unsolicited mail to any address.
        // ###########################################################################################
        public static async Task RequestPasswordResetAsync(
            string? email,
            IAccountStore store,
            IEmailSender mailer,
            ServerOptions options,
            DateTimeOffset now,
            string? ipAddress = null,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(options);

            if (string.IsNullOrWhiteSpace(email) || !AccountRules.IsPlausibleEmail(email))
                return;

            // The per-IP mail budget (MaySendMailAsync), the signed-in email change's too: this sends a
            // message to an address the caller names. Returning silently is already this method's
            // answer for every other case, so a refusal is indistinguishable from an unknown
            // address - which is exactly the property the anti-enumeration design wants.
            if (!await AccountFlows.MaySendMailAsync(ipAddress, store, now, cancellationToken))
                return;

            AccountRecord? account =
                await store.FindByNormalisedEmailAsync(AccountRules.NormaliseEmail(email), cancellationToken);

            if (account is null)
                return;

            string token = SecureToken.Create();

            await store.CreateTokenAsync(
                new NewAccountToken(
                    account.Id,
                    TokenPurpose.PasswordReset,
                    SecureToken.Hash(token),
                    now,
                    now + AccountFlows.PasswordResetTokenLifetime),
                cancellationToken);

            await mailer.SendAsync(
                EmailTemplates.PasswordReset(
                    account.Email,
                    account.DisplayName,
                    token,
                    (int)AccountFlows.PasswordResetTokenLifetime.TotalHours,
                        AccountPages.Link(AccountPages.BaseUrl(options.PublicApiBaseUrl), "forgot-password")),
                cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, "password.reset_requested",
                    $"account:{account.Id}", null, now),
                cancellationToken);
        }

        // ###########################################################################################
        // COMPLETE a password reset.
        //
        // Three things happen together and all three matter: the token is consumed, every OTHER
        // outstanding reset token for that account is consumed too (an attacker who triggered one
        // earlier must not still hold a way in), and every session is revoked (if the account was
        // taken over, the thief's session dies with the password change).
        // ###########################################################################################
        public static async Task<TokenRedemptionOutcome> CompletePasswordResetAsync(
            string? token,
            string? newPassword,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(hasher);

            if (string.IsNullOrWhiteSpace(token))
                return TokenRedemptionOutcome.Invalid;

            IReadOnlyList<string> passwordFailures = AccountRules.ValidatePassword(newPassword);

            if (passwordFailures.Count > 0)
                return TokenRedemptionOutcome.PasswordRejected;

            AccountTokenRecord? record =
                await store.FindTokenByHashAsync(SecureToken.Hash(token), cancellationToken);

            if (record is null || record.Purpose != TokenPurpose.PasswordReset)
                return TokenRedemptionOutcome.Invalid;

            if (record.ConsumedUtc is not null)
                return TokenRedemptionOutcome.AlreadyUsed;

            if (record.ExpiresUtc <= now)
                return TokenRedemptionOutcome.Expired;

            AccountRecord? account = await store.FindByIdAsync(record.AccountId, cancellationToken);

            if (account is null)
                return TokenRedemptionOutcome.Invalid;

            await store.SetPasswordHashAsync(account.Id, await hasher.HashAsync(newPassword!, cancellationToken), cancellationToken);
            await store.ConsumeTokenAsync(record.Id, now, cancellationToken);

            // Any OTHER reset link already in flight stops working.
            await store.ConsumeOutstandingTokensAsync(
                account.Id, TokenPurpose.PasswordReset, now, cancellationToken);

            // Every session dies. If the account had been taken over, this is what ends it.
            await store.RevokeAllSessionsAsync(account.Id, "password changed", now, cancellationToken);

            // Completing a reset proves control of the mailbox, so the address is verified too.
            // Someone who registered, never clicked verify, then reset their password has
            // demonstrated exactly what verification was asking for - and since 2026-11-25 it is the
            // way on from an expired link, as an unconfirmed account cannot sign in for a new one.
            if (!account.IsVerified)
                await AccountFlows.ConfirmAddressAsync(account.Id, store, mailer, now, cancellationToken);

            await mailer.SendAsync(
                EmailTemplates.PasswordChanged(account.Email, account.DisplayName),
                cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(account.Id, account.Email, "password.reset_completed",
                    $"account:{account.Id}", null, now),
                cancellationToken);

            return TokenRedemptionOutcome.Success;
        }

        // -------------------------------------------------------------------------------------
        // Helpers.
        // -------------------------------------------------------------------------------------

        // ###########################################################################################
        // Issues a session: a refresh token and an access token.
        //
        // BOTH ARE THE SAME KIND OF OPAQUE VALUE, stored as hashes against the same session row,
        // with the access token simply expiring sooner. That is a deliberate simplification over
        // two separate mechanisms: one lookup path, one revocation path, one place to get wrong.
        // ###########################################################################################
        private static async Task<IssuedSession> IssueSessionAsync(
            long accountId,
            string? userAgent,
            string? ipAddress,
            IAccountStore store,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            string refreshToken = SecureToken.Create();

            // User-agent and IP come from the request and are therefore untrusted. The database
            // columns are bounded, and an over-long value would throw on insert rather than being
            // silently cut - so they are cut here, by the one shared rule, which never leaves half a
            // character either: utf8mb4 refuses that too (code review, 2026-10-29).
            long sessionId = await store.CreateSessionAsync(
                new NewSession(
                    accountId,
                    SecureToken.Hash(refreshToken),
                    now,
                    now + TimeSpan.FromDays(options.RefreshTokenDays),
                    BoardViewRules.Clip(userAgent, 255),
                    BoardViewRules.Clip(ipAddress, 45)),
                cancellationToken);

            return new IssuedSession(
                sessionId,
                refreshToken,
                now + TimeSpan.FromDays(options.RefreshTokenDays));
        }

        // ###########################################################################################
        // May this caller cause another mail to be sent, and count it if so.
        //
        // COUNTS THE REQUEST AS IT ALLOWS IT, so the two cannot drift apart - a check that relied
        // on its caller remembering to record afterwards is a limit that silently stops working
        // the first time somebody adds a third mail-sending endpoint.
        //
        // *** NO ADDRESS MEANS NO LIMIT, and that is a deliberate, narrow hole. *** The IP is
        // absent only when the request did not arrive over HTTP - which in practice means a test.
        // UseForwardedHeaders is restricted to loopback proxies (see Program.cs), so a real caller
        // cannot suppress their own address to escape this: they would have to stop Apache from
        // setting it.
        // ###########################################################################################
        private static async Task<bool> MaySendMailAsync(
            string? ipAddress,
            IAccountStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            return (await AccountFlows.CheckMailBudgetAsync(ipAddress, store, now, cancellationToken)).IsAllowed;
        }

        // The same check with its verdict, for a caller that answers a refusal with 429 and
        // Retry-After rather than silently - the signed-in email change (AccountSelfServiceFlows),
        // whose answer must not pretend a code was sent. Counts the request as it allows it, for
        // the reason above.
        internal static async Task<RateLimitVerdict> CheckMailBudgetAsync(
            string? ipAddress,
            IAccountStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken)
        {
            if (string.IsNullOrWhiteSpace(ipAddress))
                return RateLimitVerdict.Allowed();

            IReadOnlyList<DateTimeOffset> recent = await store.GetRecentMailRequestsAsync(
                ipAddress, now - AuthRateLimitPolicy.MailWindow, cancellationToken);

            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckMailRequest(recent, now);

            if (!verdict.IsAllowed)
                return verdict;

            await store.RecordMailRequestAsync(ipAddress, now, cancellationToken);

            return verdict;
        }
    }
}
