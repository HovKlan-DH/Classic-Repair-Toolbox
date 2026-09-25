using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // "Who may do what", asserted at the level the endpoints actually enforce it: every protected
    // operation resolves its caller through AccountFlows.AuthenticateAsync, so these tests drive
    // that one function with every kind of credential a request can carry.
    //
    // WHY NOT THROUGH HTTP. Testing the real pipeline needs WebApplicationFactory, which starts an
    // in-process HTTP server - ruled out by CLAUDE.md test rule 6 and by the test csproj's own
    // header. The endpoints are deliberately thin enough that there is nothing between the request
    // and this function worth testing: each handler pulls a bearer token, calls AuthenticateAsync,
    // and returns 401 on null. What could be WRONG is which credentials that function accepts,
    // which is exactly what is below.
    //
    // THE THEME: a credential that was valid once must stop working the moment the reason for its
    // validity ends - a logout, a rotation, an expiry, a lock, a password change. That immediacy
    // is the entire justification for opaque database-backed tokens over stateless JWTs, and it is
    // what Phase 6 will depend on when removing a maintainer has to take effect at once.
    // ###########################################################################################
    public class EndpointAuthorizationTests
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

        private static Argon2PasswordHasher Hasher() => new(EndpointAuthorizationTests.Options());

        private static async Task<(FakeAccountStore Store, FakeEmailSender Mailer, string Token)> SignedInAsync(
            string email = "dennis@example.com")
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest(email, EndpointAuthorizationTests.GoodPassword, "Dennis", null),
                store, mailer, EndpointAuthorizationTests.Hasher(), EndpointAuthorizationTests.Options(),
                EndpointAuthorizationTests.Now);

            LoginOutcome login = await AccountFlows.LoginAsync(
                new LoginRequest(email, EndpointAuthorizationTests.GoodPassword, "CRT/2.6", "192.0.2.1"),
                store, EndpointAuthorizationTests.Hasher(), EndpointAuthorizationTests.Options(),
                EndpointAuthorizationTests.Now);

            return (store, mailer, login.Session!.RefreshToken);
        }

        // -----------------------------------------------------------------------------------
        // What counts as authenticated.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_valid_token_resolves_to_the_account_that_owns_it()
        {
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            Assert.NotNull(account);
            Assert.Equal("dennis@example.com", account!.Email);
        }

        [Fact]
        public async Task One_account_s_token_never_resolves_to_ANOTHER_account()
        {
            // The most basic authorization property there is, and the one whose absence would be
            // catastrophic: tokens must not be confusable between accounts.
            (FakeAccountStore store, FakeEmailSender mailer, string dennisToken) =
                await EndpointAuthorizationTests.SignedInAsync();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("other@example.com", EndpointAuthorizationTests.GoodPassword, "Other", null),
                store, mailer, EndpointAuthorizationTests.Hasher(), EndpointAuthorizationTests.Options(),
                EndpointAuthorizationTests.Now);

            LoginOutcome otherLogin = await AccountFlows.LoginAsync(
                new LoginRequest("other@example.com", EndpointAuthorizationTests.GoodPassword, null, null),
                store, EndpointAuthorizationTests.Hasher(), EndpointAuthorizationTests.Options(),
                EndpointAuthorizationTests.Now);

            AccountRecord? viaDennis = await AccountFlows.AuthenticateAsync(
                dennisToken, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            AccountRecord? viaOther = await AccountFlows.AuthenticateAsync(
                otherLogin.Session!.RefreshToken, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            Assert.Equal("dennis@example.com", viaDennis!.Email);
            Assert.Equal("other@example.com", viaOther!.Email);
            Assert.NotEqual(viaDennis.Id, viaOther.Id);
        }

        [Fact]
        public async Task An_account_flagged_administrator_is_reported_as_one()
        {
            // The flag the review app will authorise against from Phase 6. Asserted here so the
            // plumbing is known good before anything depends on it.
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            store.Accounts[1] = store.Accounts[1] with { IsAdministrator = true };

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            Assert.True(account!.IsAdministrator);
        }

        [Fact]
        public async Task An_ordinary_account_is_neither_administrator_nor_reviewer()
        {
            // Privilege is never the default. A new account gets nothing.
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            AccountRecord? account = await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            Assert.False(account!.IsAdministrator);
            Assert.False(account.IsReviewer);
        }

        // -----------------------------------------------------------------------------------
        // Credentials that must stop working IMMEDIATELY. This is the JWT-versus-opaque argument,
        // expressed as tests.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Locking_an_account_takes_effect_on_the_very_next_request()
        {
            // Phase 6's definition of done in miniature: "removing a maintainer takes effect
            // immediately, proven by a test using an already-issued token". A stateless JWT would
            // keep working until it expired, which is why this design does not use one.
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            Assert.NotNull(await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1)));

            store.Accounts[1] = store.Accounts[1] with { IsLocked = true };

            Assert.Null(await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1)));
        }

        [Fact]
        public async Task A_deleted_account_s_token_stops_working()
        {
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            store.Accounts.Remove(1);

            Assert.Null(await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1)));
        }

        [Fact]
        public async Task A_password_reset_invalidates_every_token_the_account_had()
        {
            // If the account had been taken over, this is the moment the thief loses access - so
            // it has to be immediate, not at the token's natural expiry.
            (FakeAccountStore store, FakeEmailSender mailer, string token) =
                await EndpointAuthorizationTests.SignedInAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, EndpointAuthorizationTests.Options(),
                EndpointAuthorizationTests.Now);

            await AccountFlows.CompletePasswordResetAsync(
                mailer.ExtractTokenFromLastMail(), "a brand new passphrase", store, mailer,
                EndpointAuthorizationTests.Hasher(), EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(1));

            Assert.Null(await AccountFlows.AuthenticateAsync(
                token, store, EndpointAuthorizationTests.Now + TimeSpan.FromMinutes(2)));
        }

        [Fact]
        public async Task An_expired_session_stops_authenticating_on_the_stroke()
        {
            // Boundary: exactly at the expiry the token is already dead, not still alive for that
            // instant. An "expires at" that is inclusive is an off-by-one in the wrong direction.
            (FakeAccountStore store, _, string token) = await EndpointAuthorizationTests.SignedInAsync();

            DateTimeOffset expiry = store.Sessions[1].ExpiresUtc;

            Assert.NotNull(await AccountFlows.AuthenticateAsync(
                token, store, expiry - TimeSpan.FromSeconds(1)));

            Assert.Null(await AccountFlows.AuthenticateAsync(token, store, expiry));
        }
    }
}
