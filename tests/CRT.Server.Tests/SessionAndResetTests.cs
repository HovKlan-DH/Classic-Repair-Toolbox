using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers session refresh, reuse detection, logout, authentication and password reset.
    //
    // THE REUSE-DETECTION TESTS ARE THE POINT OF THIS FILE. Rotation with reuse detection is what
    // makes a 30-day refresh lifetime acceptable: a stolen token is usable only until the
    // legitimate client refreshes next, at which point the theft is detected and the whole chain
    // dies. Without it, a stolen refresh token is a month of access.
    //
    // They are also what a stateless JWT cannot do at all, which is why the strategy document
    // requires opaque database-backed tokens - see Phase 6's "removing a maintainer takes effect
    // immediately, proven by a test using an already-issued token".
    // ###########################################################################################
    public class SessionAndResetTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private const string GoodPassword = "correct horse battery staple";

        private static ServerOptions Options()
        {
            return new ServerOptions
            {
                Argon2MemoryKib = 8192,
                Argon2Iterations = 1,
                Argon2Parallelism = 1,
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk",
                RefreshTokenDays = 30
            };
        }

        private static Argon2PasswordHasher Hasher() => new(SessionAndResetTests.Options());

        // Registers and logs in, returning everything a session test needs.
        private static async Task<(FakeAccountStore Store, FakeEmailSender Mailer, string RefreshToken)> LoggedInAsync()
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", SessionAndResetTests.GoodPassword, "Dennis", null),
                store, mailer, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(), SessionAndResetTests.Now);

            LoginOutcome login = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", SessionAndResetTests.GoodPassword, "CRT/2.6", "192.0.2.1"),
                store, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(), SessionAndResetTests.Now);

            return (store, mailer, login.Session!.RefreshToken);
        }

        // -----------------------------------------------------------------------------------
        // Refresh and rotation.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_refresh_token_can_be_exchanged_for_a_new_session()
        {
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            RefreshOutcome outcome = await AccountFlows.RefreshAsync(
                token, "CRT/2.6", "192.0.2.1", store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            Assert.True(outcome.IsSuccess);
            Assert.NotNull(outcome.Session);
            Assert.NotEqual(token, outcome.Session!.RefreshToken);
        }

        [Fact]
        public async Task Refreshing_RETIRES_the_old_token()
        {
            // Rotation: the old token must stop working the moment it is exchanged, or a copy
            // taken at any point stays valid for the full refresh lifetime.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RefreshAsync(
                token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            Assert.NotNull(store.Sessions[1].ReplacedById);
        }

        [Fact]
        public async Task Presenting_an_ALREADY_ROTATED_token_revokes_every_session()
        {
            // REUSE DETECTION. Two parties hold this token - the legitimate client and whoever
            // copied it - and there is no way to tell which is which, so the whole chain dies.
            // The legitimate user is logged out too: an inconvenient re-login beats leaving a
            // thief with a live session.
            (FakeAccountStore store, _, string stolenToken) = await SessionAndResetTests.LoggedInAsync();

            // The legitimate client refreshes normally.
            RefreshOutcome legitimate = await AccountFlows.RefreshAsync(
                stolenToken, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            Assert.True(legitimate.IsSuccess);

            // The thief then presents their copy of the OLD token.
            RefreshOutcome thief = await AccountFlows.RefreshAsync(
                stolenToken, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(25));

            Assert.False(thief.IsSuccess);

            // Every session, including the legitimate client's new one, is now revoked.
            Assert.All(store.Sessions.Values, session => Assert.NotNull(session.RevokedUtc));
        }

        [Fact]
        public async Task Reuse_detection_is_recorded_in_the_audit_trail()
        {
            // The user is about to be logged out of everything with no explanation; the audit row
            // is how the maintainer can tell them why.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RefreshAsync(token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            await AccountFlows.RefreshAsync(token, null, "203.0.113.9", store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(25));

            AuditEntry entry = Assert.Single(store.Audit, row => row.Action == "session.reuse_detected");

            Assert.Contains("203.0.113.9", entry.Detail!, StringComparison.Ordinal);
        }

        [Fact]
        public async Task The_thief_is_told_nothing_about_having_triggered_a_revocation()
        {
            // Reuse detection answers exactly like any other failed refresh.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RefreshAsync(token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            RefreshOutcome reuse = await AccountFlows.RefreshAsync(
                token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(25));

            RefreshOutcome nonsense = await AccountFlows.RefreshAsync(
                "a token that never existed", null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(25));

            Assert.Equal(nonsense, reuse);
        }

        [Fact]
        public async Task An_EXPIRED_refresh_token_is_refused()
        {
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            DateTimeOffset tooLate = SessionAndResetTests.Now + TimeSpan.FromDays(31);

            Assert.False((await AccountFlows.RefreshAsync(
                token, null, null, store, SessionAndResetTests.Options(), tooLate)).IsSuccess);
        }

        [Fact]
        public async Task A_LOCKED_account_cannot_refresh_even_with_a_valid_token()
        {
            // This is what makes "removing someone takes effect immediately" true, and is exactly
            // what a stateless JWT could not do - it would stay valid until it expired.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            store.Accounts[1] = store.Accounts[1] with { IsLocked = true };

            RefreshOutcome outcome = await AccountFlows.RefreshAsync(
                token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            Assert.False(outcome.IsSuccess);
            Assert.NotNull(store.Sessions[1].RevokedUtc);
        }

        // -----------------------------------------------------------------------------------
        // Logout.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Logging_out_revokes_the_session()
        {
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.LogoutAsync(token, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.NotNull(store.Sessions[1].RevokedUtc);
        }

        [Fact]
        public async Task Logging_out_does_not_revoke_the_account_s_OTHER_sessions()
        {
            // Signing out on a laptop must not sign you out on the bench machine.
            (FakeAccountStore store, _, string firstToken) = await SessionAndResetTests.LoggedInAsync();

            LoginOutcome second = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", SessionAndResetTests.GoodPassword, "CRT/2.6", "198.51.100.4"),
                store, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(), SessionAndResetTests.Now);

            await AccountFlows.LogoutAsync(firstToken, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            AccountRecord? stillValid = await AccountFlows.AuthenticateAsync(
                second.Session!.RefreshToken, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(6));

            Assert.NotNull(stillValid);
        }

        [Fact]
        public async Task Logging_out_with_an_unknown_token_is_harmless()
        {
            // Nothing useful to say, and distinguishing it would turn logout into a way to test
            // whether a token is live.
            (FakeAccountStore store, _, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.LogoutAsync("never existed", store, SessionAndResetTests.Now);

            Assert.Null(store.Sessions[1].RevokedUtc);
        }

        // -----------------------------------------------------------------------------------
        // Authentication (GET /api/accounts/me and everything after it).
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_live_token_authenticates()
        {
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.NotNull(account);
            Assert.Equal("dennis@example.com", account!.Email);
        }

        [Fact]
        public async Task A_REVOKED_session_stops_authenticating_immediately()
        {
            // The whole reason for opaque database-backed tokens.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.LogoutAsync(token, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(1));

            Assert.Null(await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(2)));
        }

        [Fact]
        public async Task A_ROTATED_token_stops_authenticating()
        {
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RefreshAsync(token, null, null, store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(20));

            Assert.Null(await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now + TimeSpan.FromMinutes(21)));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("not-a-token")]
        public async Task A_missing_or_bogus_token_does_not_authenticate(string? token)
        {
            (FakeAccountStore store, _, _) = await SessionAndResetTests.LoggedInAsync();

            Assert.Null(await AccountFlows.AuthenticateAsync(token, store, SessionAndResetTests.Now));
        }

        // -----------------------------------------------------------------------------------
        // Password reset.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_reset_link_sets_a_new_password_and_the_old_one_stops_working()
        {
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            string resetToken = mailer.ExtractTokenFromLastMail();

            TokenRedemptionOutcome outcome = await AccountFlows.CompletePasswordResetAsync(
                resetToken, "a brand new passphrase", store, mailer, SessionAndResetTests.Hasher(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.Equal(TokenRedemptionOutcome.Success, outcome);

            // The new one works...
            Assert.True((await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "a brand new passphrase", null, null),
                store, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(10))).IsSuccess);

            // ...and the old one does not.
            Assert.False((await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", SessionAndResetTests.GoodPassword, null, null),
                store, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(10))).IsSuccess);
        }

        [Fact]
        public async Task Completing_a_reset_revokes_every_existing_session()
        {
            // If the account had been taken over, this is what ends the thief's access.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            await AccountFlows.CompletePasswordResetAsync(
                mailer.ExtractTokenFromLastMail(), "a brand new passphrase", store, mailer,
                SessionAndResetTests.Hasher(), SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.All(store.Sessions.Values, session => Assert.NotNull(session.RevokedUtc));
        }

        [Fact]
        public async Task Completing_a_reset_invalidates_OTHER_outstanding_reset_links()
        {
            // An attacker who triggered a reset earlier must not still hold a way in after the
            // real owner has reset their password.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            string attackerToken = mailer.ExtractTokenFromLastMail();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(1));

            string ownerToken = mailer.ExtractTokenFromLastMail();

            await AccountFlows.CompletePasswordResetAsync(
                ownerToken, "the owner's new passphrase", store, mailer, SessionAndResetTests.Hasher(),
                SessionAndResetTests.Now + TimeSpan.FromMinutes(2));

            Assert.Equal(
                TokenRedemptionOutcome.AlreadyUsed,
                await AccountFlows.CompletePasswordResetAsync(
                    attackerToken, "the attacker's password", store, mailer, SessionAndResetTests.Hasher(),
                    SessionAndResetTests.Now + TimeSpan.FromMinutes(3)));
        }

        [Fact]
        public async Task Completing_a_reset_also_VERIFIES_the_address()
        {
            // Clicking a link in that mailbox proves control of it - exactly what verification was
            // asking for. Someone who registered, never verified, then reset has demonstrated it.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", SessionAndResetTests.GoodPassword, "Dennis", null),
                store, mailer, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(), SessionAndResetTests.Now);

            Assert.False(store.Accounts.Values.Single().IsVerified);

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            await AccountFlows.CompletePasswordResetAsync(
                mailer.ExtractTokenFromLastMail(), "a brand new passphrase", store, mailer,
                SessionAndResetTests.Hasher(), SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.True(store.Accounts.Values.Single().IsVerified);
        }

        [Fact]
        public async Task A_reset_request_for_an_UNKNOWN_address_sends_no_mail_at_all()
        {
            // Mailing a stranger "you have no account here" turns this endpoint into a way to send
            // unsolicited mail to any address.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RequestPasswordResetAsync(
                "nobody@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            Assert.Empty(mailer.Sent);
        }

        [Fact]
        public async Task Completing_a_reset_with_a_WEAK_password_is_refused_and_consumes_nothing()
        {
            // The token must survive a rejected password, or a typo costs the user their only
            // reset link.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            string token = mailer.ExtractTokenFromLastMail();

            Assert.Equal(
                TokenRedemptionOutcome.PasswordRejected,
                await AccountFlows.CompletePasswordResetAsync(
                    token, "short", store, mailer, SessionAndResetTests.Hasher(),
                    SessionAndResetTests.Now + TimeSpan.FromMinutes(1)));

            // Still usable with a good password.
            Assert.Equal(
                TokenRedemptionOutcome.Success,
                await AccountFlows.CompletePasswordResetAsync(
                    token, "a proper passphrase now", store, mailer, SessionAndResetTests.Hasher(),
                    SessionAndResetTests.Now + TimeSpan.FromMinutes(2)));
        }

        [Fact]
        public async Task An_EXPIRED_reset_token_is_refused()
        {
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            DateTimeOffset tooLate = SessionAndResetTests.Now + AccountFlows.PasswordResetTokenLifetime + TimeSpan.FromMinutes(1);

            Assert.Equal(
                TokenRedemptionOutcome.Expired,
                await AccountFlows.CompletePasswordResetAsync(
                    mailer.ExtractTokenFromLastMail(), "a brand new passphrase", store, mailer,
                    SessionAndResetTests.Hasher(), tooLate));
        }

        [Fact]
        public async Task A_VERIFICATION_token_cannot_be_used_to_reset_a_password()
        {
            // The twin of the verification-side test. One bug in a WHERE clause must not let a
            // verification link become a password reset - that would be a full account takeover
            // from a link that is sent to every new registration.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", SessionAndResetTests.GoodPassword, "Dennis", null),
                store, mailer, SessionAndResetTests.Hasher(), SessionAndResetTests.Options(), SessionAndResetTests.Now);

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.CompletePasswordResetAsync(
                    mailer.ExtractTokenFromLastMail(), "an attacker's password", store, mailer,
                    SessionAndResetTests.Hasher(), SessionAndResetTests.Now + TimeSpan.FromMinutes(1)));
        }

        [Fact]
        public async Task A_completed_reset_notifies_the_owner()
        {
            // The one message that tells someone whose account was taken over that it happened.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, SessionAndResetTests.Options(), SessionAndResetTests.Now);

            await AccountFlows.CompletePasswordResetAsync(
                mailer.ExtractTokenFromLastMail(), "a brand new passphrase", store, mailer,
                SessionAndResetTests.Hasher(), SessionAndResetTests.Now + TimeSpan.FromMinutes(5));

            Assert.Contains("was changed", mailer.Last!.Subject, StringComparison.Ordinal);
        }

        // -----------------------------------------------------------------------------------
        // SLIDING EXPIRY (maintainer request, 2026-09-22).
        //
        // The promise being kept: a reviewer signs in once and keeps working. The promise being
        // preserved alongside it: signing out, locking an account, or a session expiring still
        // stops access immediately. These tests assert BOTH halves, because a change that keeps
        // someone signed in by accidentally reviving dead sessions would pass the first half
        // alone.
        //
        // *** EXTENSION IS NOT ROTATION, and no test here should ever assert a changed token. ***
        // See SessionExtensionRules' header: rotating a token a desktop app keeps in a file puts a
        // disk write inside the window where the old value is already poisoned, so a power cut
        // revokes every session and looks like an attack in the audit log.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task USING_a_session_past_halfway_pushes_its_expiry_forward()
        {
            // The whole point: come back on day 20 of 30 and the clock resets, so the password is
            // never asked for again by someone who keeps using the app.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            DateTimeOffset later = SessionAndResetTests.Now.AddDays(20);

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, later, CancellationToken.None, SessionAndResetTests.Options());

            Assert.NotNull(account);

            SessionRecord session = store.Sessions.Values.Single();

            Assert.Equal(later.AddDays(30), session.ExpiresUtc);
            Assert.Equal(later, session.LastUsedUtc);
        }

        [Fact]
        public async Task Extension_does_NOT_change_the_token()
        {
            // *** THE ANTI-ROTATION ASSERTION. *** If a future change makes extension issue a new
            // token, the desktop app's stored copy silently becomes a poisoned value whose next
            // use revokes the entire account. The same token must keep working afterwards.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            DateTimeOffset later = SessionAndResetTests.Now.AddDays(20);

            await AccountFlows.AuthenticateAsync(
                token, store, later, CancellationToken.None, SessionAndResetTests.Options());

            SessionRecord session = store.Sessions.Values.Single();

            Assert.Null(session.ReplacedById);
            Assert.Null(session.RevokedUtc);

            // And it still authenticates - the real proof the client is not holding a dead value.
            Assert.NotNull(await AccountFlows.AuthenticateAsync(
                token, store, later.AddDays(1), CancellationToken.None, SessionAndResetTests.Options()));
        }

        [Fact]
        public async Task A_FRESH_session_is_not_rewritten_on_every_request()
        {
            // The review app calls the queue on launch, on every refresh click and after every
            // decision. Extending on each would turn a read-only screen into a stream of UPDATEs
            // that move the expiry by seconds.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            for (int request = 0; request < 10; request++)
            {
                await AccountFlows.AuthenticateAsync(
                    token, store, SessionAndResetTests.Now.AddMinutes(request),
                    CancellationToken.None, SessionAndResetTests.Options());
            }

            Assert.Equal(0, store.ExtensionCount);
        }

        [Fact]
        public async Task A_SIGNED_OUT_session_is_never_revived_by_a_later_request()
        {
            // *** SIGNING OUT MUST BE FINAL. *** The maintainer named logging off as a thing that
            // must forget the login, so a revoked session being extended back to life by a
            // request already in flight would break the one guarantee they asked for by name.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.LogoutAsync(token, store, SessionAndResetTests.Now.AddDays(20));

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now.AddDays(20), CancellationToken.None,
                SessionAndResetTests.Options());

            Assert.Null(account);
            Assert.Equal(0, store.ExtensionCount);
            Assert.NotNull(store.Sessions.Values.Single().RevokedUtc);
        }

        [Fact]
        public async Task A_LOCKED_account_stops_immediately_and_its_session_is_not_extended()
        {
            // *** THE OTHER THING THE MAINTAINER NAMED BY NAME. *** Phase 6's definition of done
            // requires removal to bite on the very next request rather than at natural expiry.
            // Extension runs only AFTER the lock check, so a locked account cannot prolong the
            // credential it is no longer entitled to.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            long accountId = store.Sessions.Values.Single().AccountId;
            store.Accounts[accountId] = store.Accounts[accountId] with { IsLocked = true };

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now.AddDays(20), CancellationToken.None,
                SessionAndResetTests.Options());

            Assert.Null(account);
            Assert.Equal(0, store.ExtensionCount);
            Assert.Equal(
                SessionAndResetTests.Now.AddDays(30),
                store.Sessions.Values.Single().ExpiresUtc);
        }

        [Fact]
        public async Task An_EXPIRED_session_is_never_resurrected()
        {
            // *** THE CASE THAT WOULD MAKE SESSIONS IMMORTAL. *** A token recovered from an old
            // backup or a stolen disk must not be extendable back into life. This is also what
            // keeps a sliding window SHORTER than the long fixed lifetime it replaced: a machine
            // that stops being used goes cold on its own.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now.AddDays(31), CancellationToken.None,
                SessionAndResetTests.Options());

            Assert.Null(account);
            Assert.Equal(0, store.ExtensionCount);
        }

        [Fact]
        public async Task A_ROTATED_session_is_not_extended_either()
        {
            // The review app never calls refresh, but AccountFlows.RefreshAsync still exists and
            // a future client might. A superseded session must not be handed a fresh lease.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            await AccountFlows.RefreshAsync(
                token, "CRT/2.6", "192.0.2.1", store, SessionAndResetTests.Options(),
                SessionAndResetTests.Now.AddDays(20));

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now.AddDays(21), CancellationToken.None,
                SessionAndResetTests.Options());

            Assert.Null(account);
            Assert.Equal(0, store.ExtensionCount);
        }

        [Fact]
        public async Task WITHOUT_options_nothing_is_extended_and_authentication_still_works()
        {
            // The parameter is optional so existing callers compile unchanged. A caller that does
            // not pass it must lose the convenience and nothing else - never the ability to
            // authenticate.
            (FakeAccountStore store, _, string token) = await SessionAndResetTests.LoggedInAsync();

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, SessionAndResetTests.Now.AddDays(20), CancellationToken.None);

            Assert.NotNull(account);
            Assert.Equal(0, store.ExtensionCount);
        }
    }
}
