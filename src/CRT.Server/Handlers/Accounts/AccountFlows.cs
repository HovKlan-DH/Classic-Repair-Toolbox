using CRT.Server.Configuration;
using CRT.Server.Handlers.Email;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // Registration, verification, login, refresh, logout and password reset.
    //
    // THIS IS WHERE THE DECISIONS LIVE. The endpoints in Program.cs read a request into plain
    // values, call one of these, and map the verdict onto a status code - nothing more. Everything
    // that could be wrong in an interesting way is here, and every method takes IAccountStore and
    // IEmailSender as arguments, so the whole file is testable against fakes with no database and
    // no network.
    //
    // FOUR PROPERTIES THIS FILE EXISTS TO GUARANTEE:
    //
    // 1. REGISTRATION AND RESET NEVER REVEAL WHETHER AN ADDRESS IS REGISTERED. Both answer the
    //    same neutral result whatever they find. What differs is which mail is sent, which only
    //    the mailbox owner sees. Any change that makes the RESULT depend on whether the account
    //    existed reopens enumeration.
    //
    // 2. A LOGIN THAT FAILS TELLS YOU NOTHING ABOUT WHY. Unknown address, wrong password and
    //    locked account are one answer. "No such account" would otherwise be a free account
    //    enumeration oracle on the one endpoint everybody can reach.
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
        // How long a verification link lasts. Long, because someone may register in the evening
        // and read their mail the next day, and an expired link is a support burden for a
        // volunteer project with no support desk.
        public static readonly TimeSpan VerificationTokenLifetime = TimeSpan.FromHours(24);

        // Short, because a reset link is a live credential for an account whose owner may not be
        // the person who requested it.
        public static readonly TimeSpan PasswordResetTokenLifetime = TimeSpan.FromHours(2);

        // ###########################################################################################
        // REGISTER.
        //
        // Answers the same neutral "accepted" for an available address and for one already taken.
        // The difference is which mail goes out, and that is visible only to the mailbox owner.
        //
        // Validation failures DO answer differently, and that is correct: "your password is too
        // short" is about the input, not about who exists.
        // ###########################################################################################
        public static async Task<RegistrationOutcome> RegisterAsync(
            RegistrationRequest request,
            IAccountStore store,
            IEmailSender mailer,
            Argon2PasswordHasher hasher,
            ServerOptions options,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(request);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(mailer);
            ArgumentNullException.ThrowIfNull(hasher);
            ArgumentNullException.ThrowIfNull(options);

            var failures = new List<string>();

            if (!AccountRules.IsPlausibleEmail(request.Email))
                failures.Add("That does not look like an email address.");

            failures.AddRange(AccountRules.ValidatePassword(request.Password));
            failures.AddRange(AccountRules.ValidateDisplayName(request.DisplayName));

            if (failures.Count > 0)
                return RegistrationOutcome.Invalid(failures);

            // ---- The mail rate limit -----------------------------------------------------------
            //
            // *** BOTH BRANCHES BELOW SEND MAIL, so this endpoint is a way to put a message into
            // any address somebody names. *** Unlimited, that is an inbox-flooding tool wearing
            // this service's reputation, and the security model requires a limit here
            // (NewContributeStrategy.md, threat 6).
            //
            // AFTER validation and BEFORE the account lookup: a malformed request costs nothing
            // and must not spend the budget, while everything past this point either writes a row
            // or sends a mail.
            //
            // A refusal answers the SAME neutral Accepted as everything else here. Saying "you
            // are being rate limited" would tell a caller their probing was noticed, and - worse -
            // a distinct answer that appeared only sometimes would itself be a signal to correlate
            // against the addresses tried.
            if (!await AccountFlows.MaySendMailAsync(request.IpAddress, store, now, cancellationToken))
                return RegistrationOutcome.Accepted();

            // IsPlausibleEmail and ValidateDisplayName have already rejected null for both of
            // these, but the compiler cannot see that across the failure list - hence the
            // assertions rather than a null-forgiving operator, so a future reordering of the
            // checks fails loudly here instead of producing a NullReferenceException deeper in.
            ArgumentException.ThrowIfNullOrWhiteSpace(request.Email);
            ArgumentException.ThrowIfNullOrWhiteSpace(request.DisplayName);
            ArgumentException.ThrowIfNullOrEmpty(request.Password);

            string normalised = AccountRules.NormaliseEmail(request.Email);

            AccountRecord? existing = await store.FindByNormalisedEmailAsync(normalised, cancellationToken);

            if (existing is not null)
            {
                // The address is taken. Send the EXISTING owner a mail carrying a reset link,
                // using THEIR display name - never the one supplied by whoever is registering,
                // which would let a stranger put text into somebody else's inbox.
                string resetToken = SecureToken.Create();

                await store.CreateTokenAsync(
                    new NewAccountToken(
                        existing.Id,
                        TokenPurpose.PasswordReset,
                        SecureToken.Hash(resetToken),
                        now,
                        now + AccountFlows.PasswordResetTokenLifetime),
                    cancellationToken);

                await mailer.SendAsync(
                    EmailTemplates.AlreadyRegistered(
                        existing.Email,
                        existing.DisplayName,
                        resetToken,
                        (int)AccountFlows.PasswordResetTokenLifetime.TotalHours),
                    cancellationToken);

                await store.WriteAuditAsync(
                    new AuditEntry(
                        existing.Id,
                        existing.Email,
                        "register.duplicate",
                        $"account:{existing.Id}",
                        "Registration attempted with an address that already has an account.",
                        now),
                    cancellationToken);

                // The SAME outcome as a successful registration. This is the anti-enumeration
                // guarantee; changing it reopens the hole.
                return RegistrationOutcome.Accepted();
            }

            long accountId = await store.CreateAccountAsync(
                new NewAccount(
                    request.Email.Trim(),
                    normalised,
                    hasher.Hash(request.Password!),
                    request.DisplayName!.Trim(),
                    now),
                cancellationToken);

            string verificationToken = SecureToken.Create();

            await store.CreateTokenAsync(
                new NewAccountToken(
                    accountId,
                    TokenPurpose.EmailVerification,
                    SecureToken.Hash(verificationToken),
                    now,
                    now + AccountFlows.VerificationTokenLifetime),
                cancellationToken);

            await mailer.SendAsync(
                EmailTemplates.Verification(
                    request.Email.Trim(),
                    request.DisplayName.Trim(),
                    AccountFlows.BuildLink(options, "verify", verificationToken),
                    (int)AccountFlows.VerificationTokenLifetime.TotalHours),
                cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(accountId, request.Email.Trim(), "register", $"account:{accountId}", null, now),
                cancellationToken);

            return RegistrationOutcome.Accepted();
        }

        // ###########################################################################################
        // VERIFY an email address from the link in the verification mail.
        //
        // Unlike registration, this DOES distinguish its failures - an expired link and an invalid
        // one need different things from the user (ask for a new one; check you copied it all).
        // Neither answer reveals whether any account exists, because the token is the only subject.
        // ###########################################################################################
        public static async Task<TokenRedemptionOutcome> VerifyEmailAsync(
            string? token,
            IAccountStore store,
            DateTimeOffset now,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(store);

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
            await store.SetVerifiedAsync(record.AccountId, cancellationToken);

            await store.WriteAuditAsync(
                new AuditEntry(record.AccountId, $"account:{record.AccountId}", "verify.email",
                    $"account:{record.AccountId}", null, now),
                cancellationToken);

            return TokenRedemptionOutcome.Success;
        }

        // ###########################################################################################
        // LOG IN.
        //
        // THE ORDER OF OPERATIONS IN HERE IS THE SECURITY DESIGN. Read it before changing it:
        //
        //   1. Rate limit FIRST, before any hashing. Argon2 allocates ~128 MiB per verification.
        //   2. Find the account. If missing, STILL record a failure and answer the generic
        //      result - otherwise the response time and the failure count both leak which
        //      addresses exist.
        //   3. Verify the password.
        //   4. Every failure answers the same thing.
        //
        // An UNVERIFIED account may log in. That is deliberate: blocking it makes "resend my
        // verification mail" unreachable, since that endpoint needs to know who is asking. What an
        // unverified account may not do is submit, which is enforced where submissions are
        // accepted.
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

            IReadOnlyList<DateTimeOffset> addressFailures = request.IpAddress is null
                ? Array.Empty<DateTimeOffset>()
                : await store.GetRecentAuthFailuresAsync(null, request.IpAddress, addressSince, cancellationToken);

            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(accountFailures, addressFailures, now);

            if (!verdict.IsAllowed)
                return LoginOutcome.RateLimited(verdict.RetryAfter);

            // ---- 2. Find the account. ----
            AccountRecord? account = await store.FindByNormalisedEmailAsync(normalised, cancellationToken);

            if (account is null)
            {
                // Record the failure even though there is no account. Without this, an attacker
                // can probe addresses indefinitely without ever tripping the limiter, and the
                // difference in behaviour is itself the oracle.
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);
                return LoginOutcome.Failed();
            }

            if (account.IsLocked)
            {
                await store.RecordAuthFailureAsync(normalised, request.IpAddress, now, cancellationToken);

                // The SAME answer as a wrong password. A distinct "your account is locked" reply
                // confirms the address exists.
                return LoginOutcome.Failed();
            }

            // ---- 3. Verify. ----
            PasswordVerificationResult verification = hasher.Verify(request.Password, account.PasswordHash);

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
                // limiter's own row deliberately does not record - carries its own action.
                await store.WriteAuditAsync(
                    new AuditEntry(account.Id, account.Email, "login.failed.account", $"account:{account.Id}",
                        request.IpAddress, now),
                    cancellationToken);

                return LoginOutcome.Failed();
            }

            // ---- 4. Success. ----

            // The parameters were raised since this password was last hashed, and the plaintext is
            // legitimately in hand for this one instant - so upgrade it now.
            if (verification.NeedsRehash)
                await store.SetPasswordHashAsync(account.Id, hasher.Hash(request.Password), cancellationToken);

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
        // Answers the same neutral result whatever it finds, exactly like registration. An unknown
        // address gets NO MAIL AT ALL - mailing a stranger "you have no account here" turns this
        // endpoint into a way to send unsolicited mail to any address.
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

            // The same mail budget registration spends from, and for the same reason: this sends a
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
                    (int)AccountFlows.PasswordResetTokenLifetime.TotalHours),
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

            await store.SetPasswordHashAsync(account.Id, hasher.Hash(newPassword!), cancellationToken);
            await store.ConsumeTokenAsync(record.Id, now, cancellationToken);

            // Any OTHER reset link already in flight stops working.
            await store.ConsumeOutstandingTokensAsync(
                account.Id, TokenPurpose.PasswordReset, now, cancellationToken);

            // Every session dies. If the account had been taken over, this is what ends it.
            await store.RevokeAllSessionsAsync(account.Id, "password changed", now, cancellationToken);

            // Completing a reset proves control of the mailbox, so the address is verified too.
            // Someone who registered, never clicked verify, then reset their password has
            // demonstrated exactly what verification was asking for.
            if (!account.IsVerified)
                await store.SetVerifiedAsync(account.Id, cancellationToken);

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

            long sessionId = await store.CreateSessionAsync(
                new NewSession(
                    accountId,
                    SecureToken.Hash(refreshToken),
                    now,
                    now + TimeSpan.FromDays(options.RefreshTokenDays),
                    AccountFlows.Truncate(userAgent, 255),
                    AccountFlows.Truncate(ipAddress, 45)),
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
            if (string.IsNullOrWhiteSpace(ipAddress))
                return true;

            IReadOnlyList<DateTimeOffset> recent = await store.GetRecentMailRequestsAsync(
                ipAddress, now - AuthRateLimitPolicy.MailWindow, cancellationToken);

            if (!AuthRateLimitPolicy.CheckMailRequest(recent, now).IsAllowed)
                return false;

            await store.RecordMailRequestAsync(ipAddress, now, cancellationToken);

            return true;
        }

        // ###########################################################################################
        // Builds a clickable mail link. The base URL is validated as an absolute https URL at
        // startup, so this only has to join the parts without doubling or dropping the separator.
        //
        // *** ONLY "verify" MAY BE PASSED HERE, AND THAT IS NOT A STYLE POINT. *** A link in a mail
        // is followed by a GET, so the path handed to this method must be one the server actually
        // maps with MapGet. Exactly one does: "/verify". It was also called with "reset" until
        // 2026-09-22, producing ".../api/accounts/reset?token=..." - a path nothing is mapped at,
        // because resetting needs a new PASSWORD and so is POST /reset-password. Every reset mail
        // ever sent therefore led to a 404, unnoticed from Phase 3 until the project owner clicked
        // one. Reset mails now carry a code to paste into the Maintainer tab instead.
        //
        // Before adding a second caller, confirm the path is mapped with MapGet in
        // AccountEndpoints - a wrong value here fails in a mail client days later, not at build.
        // ###########################################################################################
        private static string BuildLink(ServerOptions options, string path, string token)
        {
            string baseUrl = options.PublicApiBaseUrl!.TrimEnd('/');

            return $"{baseUrl}/accounts/{path}?token={Uri.EscapeDataString(token)}";
        }

        // User-agent and IP come from the request and are therefore untrusted. The database
        // columns are bounded, and an over-long value would throw on insert rather than being
        // silently cut - so cut it here, where it is a deliberate decision.
        private static string? Truncate(string? value, int maximumLength)
        {
            if (string.IsNullOrEmpty(value))
                return null;

            return value.Length <= maximumLength ? value : value[..maximumLength];
        }
    }
}
