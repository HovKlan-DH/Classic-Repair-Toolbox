using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // What happens when the input is hostile, malformed or simply absent.
    //
    // THE RULE THIS FILE ENFORCES: a bad request produces a refusal, never an exception and never
    // a partial write. An unhandled exception in a handler is a 500, and a 500 is information -
    // it tells an attacker that this particular input reached somewhere interesting. It also
    // takes the process down at startup-time code paths, as this project has already seen with
    // .NET's abort-on-unhandled-exception.
    //
    // Several of these look trivial. They are here because "obviously it handles null" is exactly
    // the assumption that produces a NullReferenceException in production, and because a future
    // refactor that reorders a validation check will fail these rather than shipping.
    // ###########################################################################################
    public class EndpointNegativeTests
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

        private static Argon2PasswordHasher Hasher() => new(EndpointNegativeTests.Options());

        // -----------------------------------------------------------------------------------
        // Registration with rubbish input.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null, null, null)]
        [InlineData("", "", "")]
        [InlineData("   ", "   ", "   ")]
        [InlineData("dennis@example.com", null, "Dennis")]
        [InlineData(null, "correct horse battery staple", "Dennis")]
        [InlineData("dennis@example.com", "correct horse battery staple", null)]
        public async Task A_registration_with_missing_fields_is_refused_without_throwing(
            string? email, string? password, string? displayName)
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(email, password, displayName, null),
                store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.NotEmpty(outcome.Failures);

            // Nothing was written and nothing was sent - a rejected request must leave no trace.
            Assert.Empty(store.Accounts);
            Assert.Empty(store.Tokens);
            Assert.Empty(mailer.Sent);
        }

        [Fact]
        public async Task A_registration_carrying_a_newline_in_the_display_name_is_refused()
        {
            // A newline in a name reaches a mail body and a log line. Refusing it here is cheaper
            // than escaping it everywhere it is later rendered.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(
                    "dennis@example.com",
                    EndpointNegativeTests.GoodPassword,
                    "Dennis\r\nBcc: someone@example.com",
                    null),
                store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsAccepted);
            Assert.Empty(store.Accounts);
        }

        [Fact]
        public async Task A_registration_with_a_SQL_shaped_address_is_treated_as_ordinary_text()
        {
            // Not because SQL injection is plausible here - every value goes through a parameter -
            // but because the refusal must come from the address being malformed, with no special
            // handling and no exception.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(
                    "'; DROP TABLE accounts; --",
                    EndpointNegativeTests.GoodPassword,
                    "Dennis",
                    null),
                store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsAccepted);
        }

        [Fact]
        public async Task A_display_name_of_exactly_the_maximum_length_is_accepted()
        {
            // Boundary in the permissive direction: an off-by-one here silently rejects a legal
            // name, which nobody reports as a bug - they just give up.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest(
                    "dennis@example.com",
                    EndpointNegativeTests.GoodPassword,
                    new string('a', AccountRules.MaximumDisplayNameLength),
                    null),
                store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.True(outcome.IsAccepted);
        }

        // -----------------------------------------------------------------------------------
        // Login with rubbish input.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null, null)]
        [InlineData("", "")]
        [InlineData("dennis@example.com", null)]
        [InlineData(null, "correct horse battery staple")]
        [InlineData("not-an-address", "correct horse battery staple")]
        public async Task A_login_with_missing_or_malformed_fields_fails_without_throwing(
            string? email, string? password)
        {
            var store = new FakeAccountStore();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest(email, password, null, null),
                store, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsSuccess);
            Assert.Null(outcome.Session);
        }

        [Fact]
        public async Task A_login_against_an_EMPTY_database_does_not_throw()
        {
            // The very first request a freshly migrated service ever receives.
            var store = new FakeAccountStore();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("nobody@example.com", "anything at all", null, "192.0.2.1"),
                store, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsSuccess);
        }

        [Fact]
        public async Task A_login_with_an_ABSURDLY_long_password_does_not_reach_the_hasher()
        {
            // Unbounded input reaching Argon2 is a memory amplification: the caller controls the
            // input size and the hasher allocates per call. It must fail on length, fast.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", EndpointNegativeTests.GoodPassword, "Dennis", null),
                store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", new string('x', 100_000), null, null),
                store, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                EndpointNegativeTests.Now);

            Assert.False(outcome.IsSuccess);
        }

        // -----------------------------------------------------------------------------------
        // Tokens that are absent, malformed or from somewhere else.
        // -----------------------------------------------------------------------------------

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData("../../etc/passwd")]
        [InlineData("<script>alert(1)</script>")]
        [InlineData("00000000000000000000000000000000000000000000")]
        public async Task A_rubbish_verification_token_is_refused_without_throwing(string? token)
        {
            var store = new FakeAccountStore();

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.VerifyEmailAsync(token, store, EndpointNegativeTests.Now));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not-a-token")]
        public async Task A_rubbish_refresh_token_is_refused_without_throwing(string? token)
        {
            var store = new FakeAccountStore();

            RefreshOutcome outcome = await AccountFlows.RefreshAsync(
                token, null, null, store, EndpointNegativeTests.Options(), EndpointNegativeTests.Now);

            Assert.False(outcome.IsSuccess);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not-a-token")]
        public async Task Logging_out_with_a_rubbish_token_does_not_throw(string? token)
        {
            var store = new FakeAccountStore();

            await AccountFlows.LogoutAsync(token, store, EndpointNegativeTests.Now);
        }

        [Fact]
        public async Task A_reset_completion_with_no_token_and_no_password_does_not_throw()
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            TokenRedemptionOutcome outcome = await AccountFlows.CompletePasswordResetAsync(
                null, null, store, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Now);

            Assert.Equal(TokenRedemptionOutcome.Invalid, outcome);
            Assert.Empty(mailer.Sent);
        }

        [Fact]
        public async Task A_forgot_password_request_with_a_malformed_address_does_nothing_quietly()
        {
            // No mail, no exception, and - critically - the same observable behaviour as a
            // well-formed address that has no account.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RequestPasswordResetAsync(
                "not-an-address", store, mailer, EndpointNegativeTests.Options(), EndpointNegativeTests.Now);

            await AccountFlows.RequestPasswordResetAsync(
                null, store, mailer, EndpointNegativeTests.Options(), EndpointNegativeTests.Now);

            Assert.Empty(mailer.Sent);
        }

        // -----------------------------------------------------------------------------------
        // Argument contracts. These throw on purpose - a null STORE is a wiring defect in our own
        // code, not hostile input, and must fail loudly at the first call rather than producing a
        // confusing null deref three frames down.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_missing_collaborator_throws_rather_than_failing_quietly()
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await Assert.ThrowsAsync<ArgumentNullException>(() =>
                AccountFlows.RegisterAsync(
                    new RegistrationRequest("a@b.com", EndpointNegativeTests.GoodPassword, "A", null),
                    null!, mailer, EndpointNegativeTests.Hasher(), EndpointNegativeTests.Options(),
                    EndpointNegativeTests.Now));

            await Assert.ThrowsAsync<ArgumentNullException>(() =>
                AccountFlows.LoginAsync(
                    new LoginRequest("a@b.com", "x", null, null),
                    store, null!, EndpointNegativeTests.Options(), EndpointNegativeTests.Now));

            await Assert.ThrowsAsync<ArgumentNullException>(() =>
                AccountFlows.AuthenticateAsync("token", null!, EndpointNegativeTests.Now));
        }
    }
}
