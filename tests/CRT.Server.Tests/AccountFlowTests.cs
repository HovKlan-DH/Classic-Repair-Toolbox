using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers registration, verification, login and password reset end to end against fakes.
    //
    // These are the flows an attacker reaches first, so the tests are written around what must NOT
    // leak as much as around what must work. The enumeration tests in particular assert that two
    // different situations produce the SAME answer - a test shape that looks odd until you
    // remember the failure it prevents: a registration endpoint that says "that address is taken"
    // is a list of everyone with an account.
    //
    // Weak Argon2 parameters throughout - see PasswordHashingTests for why.
    // ###########################################################################################
    public class AccountFlowTests
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

        private static Argon2PasswordHasher Hasher() => new(AccountFlowTests.Options());

        // Registers an account and returns the store, mailer and the verification token that was
        // mailed - the starting point for most of the tests below.
        private static async Task<(FakeAccountStore Store, FakeEmailSender Mailer, string Token)> RegisteredAsync(
            string email = "dennis@example.com")
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest(email, AccountFlowTests.GoodPassword, "Dennis", "192.0.2.1"),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            return (store, mailer, mailer.ExtractTokenFromLastMail());
        }

        // -----------------------------------------------------------------------------------
        // Registration.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Registering_creates_an_unverified_account_and_mails_a_verification_link()
        {
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            AccountRecord account = Assert.Single(store.Accounts.Values);
            Assert.Equal("dennis@example.com", account.Email);
            Assert.False(account.IsVerified);

            EmailMessageAssert(mailer, "Confirm your");
        }

        [Fact]
        public async Task The_stored_password_is_hashed_not_the_plaintext()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            AccountRecord account = Assert.Single(store.Accounts.Values);

            Assert.DoesNotContain(AccountFlowTests.GoodPassword, account.PasswordHash, StringComparison.Ordinal);
            Assert.StartsWith("$argon2id$", account.PasswordHash, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Only_the_token_HASH_is_stored_never_the_token_itself()
        {
            // A database read must not yield working verification or reset links.
            (FakeAccountStore store, _, string token) = await AccountFlowTests.RegisteredAsync();

            string stored = Assert.Single(store.TokenHashes.Values);

            Assert.NotEqual(token, stored);
            Assert.Equal(SecureToken.Hash(token), stored);
        }

        [Fact]
        public async Task Registering_with_an_address_that_ALREADY_EXISTS_gives_the_SAME_answer()
        {
            // THE anti-enumeration property. A different status, a different message or a
            // different set of failures here would make this endpoint a way to test whether any
            // given address has an account.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            RegistrationOutcome second = await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", "a completely different password", "Someone Else", "203.0.113.9"),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(second.IsAccepted);
            Assert.Empty(second.Failures);
        }

        [Fact]
        public async Task Registering_with_an_existing_address_does_not_create_a_second_account()
        {
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", "a completely different password", "Someone Else", null),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Single(store.Accounts);
        }

        [Fact]
        public async Task Registering_with_an_existing_address_does_not_change_the_existing_password()
        {
            // Otherwise anyone could overwrite anyone's password by "registering" again.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            string originalHash = store.Accounts.Values.Single().PasswordHash;

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", "attacker chosen password", "Someone Else", null),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal(originalHash, store.Accounts.Values.Single().PasswordHash);
        }

        [Fact]
        public async Task The_already_registered_mail_goes_to_the_OWNER_and_uses_THEIR_name()
        {
            // The mail goes to the existing account's inbox, so the display name supplied by
            // whoever is registering must never appear in it - that would be an injection into
            // somebody else's mail.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", "another password", "ATTACKER TEXT", null),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal("dennis@example.com", mailer.Last!.ToAddress);
            Assert.Contains("Hello Dennis,", mailer.Last.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("ATTACKER TEXT", mailer.Last.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Registering_with_a_differently_cased_address_is_treated_as_the_same_account()
        {
            // The unique index is on the normalised address, so this must not create a second
            // account - and must produce the already-registered mail instead.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("DENNIS@Example.COM", AccountFlowTests.GoodPassword, "Dennis", null),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Single(store.Accounts);
            Assert.Contains("already have", mailer.Last!.Subject, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task An_invalid_registration_is_rejected_and_sends_no_mail()
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = await AccountFlows.RegisterAsync(
                new RegistrationRequest("not-an-address", "short", "", null),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.False(outcome.IsAccepted);

            // Every problem at once, not just the first.
            Assert.True(outcome.Failures.Count >= 3, $"Got: {string.Join(" | ", outcome.Failures)}");

            Assert.Empty(mailer.Sent);
            Assert.Empty(store.Accounts);
        }

        // -----------------------------------------------------------------------------------
        // Verification.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_verification_token_from_the_mail_verifies_the_account()
        {
            (FakeAccountStore store, _, string token) = await AccountFlowTests.RegisteredAsync();

            TokenRedemptionOutcome outcome = await AccountFlows.VerifyEmailAsync(
                token, store, AccountFlowTests.Now);

            Assert.Equal(TokenRedemptionOutcome.Success, outcome);
            Assert.True(store.Accounts.Values.Single().IsVerified);
        }

        [Fact]
        public async Task A_verification_token_cannot_be_used_twice()
        {
            // Without this, a verification link sitting in a mailbox stays live forever.
            (FakeAccountStore store, _, string token) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.VerifyEmailAsync(token, store, AccountFlowTests.Now);

            Assert.Equal(
                TokenRedemptionOutcome.AlreadyUsed,
                await AccountFlows.VerifyEmailAsync(token, store, AccountFlowTests.Now));
        }

        [Fact]
        public async Task An_expired_verification_token_is_refused()
        {
            (FakeAccountStore store, _, string token) = await AccountFlowTests.RegisteredAsync();

            DateTimeOffset tooLate = AccountFlowTests.Now + AccountFlows.VerificationTokenLifetime + TimeSpan.FromMinutes(1);

            Assert.Equal(
                TokenRedemptionOutcome.Expired,
                await AccountFlows.VerifyEmailAsync(token, store, tooLate));

            Assert.False(store.Accounts.Values.Single().IsVerified);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not-a-real-token")]
        public async Task A_bogus_verification_token_is_refused(string? token)
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.VerifyEmailAsync(token, store, AccountFlowTests.Now));
        }

        [Fact]
        public async Task A_PASSWORD_RESET_token_cannot_be_used_to_verify_an_address()
        {
            // The purposes are separate for exactly this reason: one bug in a WHERE clause must
            // not let a reset token act as a verification, or vice versa.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, AccountFlowTests.Options(), AccountFlowTests.Now);

            string resetToken = mailer.ExtractTokenFromLastMail();

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.VerifyEmailAsync(resetToken, store, AccountFlowTests.Now));
        }

        // -----------------------------------------------------------------------------------
        // Login.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_correct_password_logs_in_and_issues_a_session()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, "CRT/2.6", "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);
            Assert.NotNull(outcome.Session);
            Assert.NotEmpty(outcome.Session!.RefreshToken);
            Assert.Single(store.Sessions);
        }

        [Fact]
        public async Task An_UNVERIFIED_account_can_still_log_in()
        {
            // Deliberate: blocking login makes "resend my verification mail" unreachable, since
            // that endpoint needs to know who is asking. What an unverified account may not do is
            // submit, which is enforced where submissions are accepted.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            Assert.False(store.Accounts.Values.Single().IsVerified);

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);
        }

        [Fact]
        public async Task A_wrong_password_and_an_UNKNOWN_ADDRESS_fail_identically()
        {
            // The enumeration guarantee on the one endpoint everybody can reach. If these two
            // answered differently - a different status, a different message, even a different
            // shape - login would be a way to discover who has an account.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            LoginOutcome wrongPassword = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "the wrong password entirely", null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            LoginOutcome unknownAddress = await AccountFlows.LoginAsync(
                new LoginRequest("nobody@example.com", "the wrong password entirely", null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal(wrongPassword, unknownAddress);
        }

        [Fact]
        public async Task A_LOCKED_account_fails_exactly_like_a_wrong_password()
        {
            // A distinct "your account is locked" reply would confirm the address exists.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            store.Accounts[1] = store.Accounts[1] with { IsLocked = true };

            LoginOutcome locked = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.False(locked.IsSuccess);
            Assert.False(locked.IsRateLimited);
            Assert.Null(locked.Session);
        }

        // -----------------------------------------------------------------------------------
        // The mail-sending budget.
        //
        // Registration and "forgot my password" both send a message to an address the CALLER
        // names, so without a limit this service is a way to flood a stranger's inbox under its
        // own reputation. NewContributeStrategy.md's threat 6 requires the limit; these pin that
        // it is actually applied, which AuthRateLimitPolicy's own tests cannot - the policy was
        // written, tested and then wired to nothing at all.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Registration_stops_sending_mail_once_the_per_IP_budget_is_spent()
        {
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            // One more than the budget, every one from the same address and each naming a
            // DIFFERENT victim - the shape inbox flooding actually takes.
            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
            {
                await AccountFlows.RegisterAsync(
                    new RegistrationRequest(
                        $"victim{attempt}@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                    store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            Assert.Equal(AuthRateLimitPolicy.MaxMailRequestsPerAddress, mailer.Sent.Count);
        }

        [Fact]
        public async Task A_rate_limited_registration_still_answers_ACCEPTED()
        {
            // The anti-enumeration guarantee survives the limit. A distinct "you are being rate
            // limited" answer would tell a caller their probing was noticed, and an answer that
            // changed only sometimes would itself be a signal to correlate against the addresses
            // that had been tried.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            RegistrationOutcome outcome = RegistrationOutcome.Accepted();

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
            {
                outcome = await AccountFlows.RegisterAsync(
                    new RegistrationRequest(
                        $"victim{attempt}@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                    store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            Assert.True(outcome.IsAccepted);

            // And no account was created for the refused one, so the limit is not merely silencing
            // the mail while still doing the work.
            Assert.Equal(AuthRateLimitPolicy.MaxMailRequestsPerAddress, store.Accounts.Count);
        }

        [Fact]
        public async Task An_INVALID_registration_does_not_spend_the_mail_budget()
        {
            // The limit sits after validation deliberately: a malformed request sends nothing, so
            // charging it would let a clumsy client lock itself out of registering at all.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxMailRequestsPerAddress * 2; attempt++)
            {
                await AccountFlows.RegisterAsync(
                    new RegistrationRequest("not-an-address", "short", "Dennis", "192.0.2.9"),
                    store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            Assert.Empty(store.MailRequests);

            // A valid one afterwards still gets through.
            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Single(mailer.Sent);
        }

        [Fact]
        public async Task Password_reset_requests_share_the_same_per_IP_budget()
        {
            // Both endpoints send mail to a caller-named address, so a limit on only one of them
            // just moves the flooding to the other.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.RegisteredAsync();

            int alreadySent = mailer.Sent.Count;

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
            {
                await AccountFlows.RequestPasswordResetAsync(
                    "dennis@example.com",
                    store,
                    mailer,
                    AccountFlowTests.Options(),
                    AccountFlowTests.Now,
                    "192.0.2.9");
            }

            // The registration in RegisteredAsync came from a different address, so the whole
            // budget was available here - and exactly that many reset mails went out.
            Assert.Equal(alreadySent + AuthRateLimitPolicy.MaxMailRequestsPerAddress, mailer.Sent.Count);
        }

        [Fact]
        public async Task The_mail_budget_is_per_ADDRESS_so_one_caller_cannot_block_everybody()
        {
            // A single global bucket would mean one abuser stops every other person on the
            // internet from registering - the limit becoming the denial of service.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
            {
                await AccountFlows.RegisterAsync(
                    new RegistrationRequest(
                        $"victim{attempt}@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                    store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            int spent = mailer.Sent.Count;

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("someone@example.com", AccountFlowTests.GoodPassword, "Dennis", "198.51.100.4"),
                store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal(spent + 1, mailer.Sent.Count);
        }

        [Fact]
        public async Task The_mail_budget_EXPIRES_rather_than_locking_an_address_out_for_ever()
        {
            // The same reasoning as the login lockout: a limit that never clears is a permanent
            // ban on whatever shares that address - a household, a workplace, a whole country
            // behind CGNAT.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
            {
                await AccountFlows.RegisterAsync(
                    new RegistrationRequest(
                        $"victim{attempt}@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                    store, mailer, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            int spent = mailer.Sent.Count;

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("later@example.com", AccountFlowTests.GoodPassword, "Dennis", "192.0.2.9"),
                store,
                mailer,
                AccountFlowTests.Hasher(),
                AccountFlowTests.Options(),
                AccountFlowTests.Now + AuthRateLimitPolicy.MailWindow + TimeSpan.FromMinutes(1));

            Assert.Equal(spent + 1, mailer.Sent.Count);
        }

        [Fact]
        public async Task A_wrong_password_counts_as_ONE_failure_against_the_rate_limiter()
        {
            // *** THE AUDIT ROW MUST NOT REUSE THE LIMITER'S OWN ACTION STRING. ***
            //
            // In MySqlAccountStore both RecordAuthFailureAsync and WriteAuditAsync land in the SAME
            // `audit` table, and the limiter counts rows by matching `action`. This fake keeps the
            // two in separate lists, so it CANNOT reproduce that - which is precisely how the bug
            // shipped unnoticed. The invariant is therefore asserted here directly, on the action
            // string, because that string is the whole coupling between the two writes.
            //
            // The consequences of getting it wrong were both real: one wrong password counted
            // twice against the per-IP bucket (30 attempts became 15), and because
            // ClearAuthFailuresAsync only clears rows whose subject is the address, the extra
            // "account:{id}" row survived every successful login and accumulated for ever.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "the wrong password entirely", null, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            // Exactly one row is what the limiter is allowed to see for one wrong password.
            Assert.Single(store.AuthFailures);

            // The descriptive audit row still exists - it names WHICH account was targeted, which
            // the limiter's own row deliberately does not record - but under its own action.
            Assert.DoesNotContain(store.Audit, entry => entry.Action == "login.failed");
            Assert.Contains(store.Audit, entry => entry.Action == "login.failed.account");
        }

        [Fact]
        public async Task A_failed_login_against_an_UNKNOWN_address_still_records_a_failure()
        {
            // Otherwise an attacker can probe addresses indefinitely without ever tripping the
            // limiter - and the difference in behaviour is itself the oracle.
            var store = new FakeAccountStore();

            await AccountFlows.LoginAsync(
                new LoginRequest("nobody@example.com", "guess", null, "203.0.113.9"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Single(store.AuthFailures);
        }

        [Fact]
        public async Task Repeated_failures_trip_the_rate_limiter()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxFailuresPerAccount; attempt++)
            {
                await AccountFlows.LoginAsync(
                    new LoginRequest("dennis@example.com", "wrong", null, "192.0.2.1"),
                    store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            // Even the CORRECT password is refused while the limit holds.
            Assert.True(outcome.IsRateLimited);
            Assert.False(outcome.IsSuccess);
            Assert.True(outcome.RetryAfter > TimeSpan.Zero);
        }

        [Fact]
        public async Task A_successful_login_clears_the_failure_count()
        {
            // Someone who gets it right on the fourth try is not an attacker and must not be
            // locked out afterwards.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            for (int attempt = 0; attempt < 3; attempt++)
            {
                await AccountFlows.LoginAsync(
                    new LoginRequest("dennis@example.com", "wrong", null, "192.0.2.1"),
                    store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Empty(store.AuthFailures);
        }

        [Fact]
        public async Task The_rate_limiter_runs_BEFORE_the_password_is_checked()
        {
            // Argon2 allocates ~128 MiB per verification, so a limiter that runs after the hash is
            // decoration and leaves the memory-exhaustion vector open. Proven by observing that a
            // rate-limited request does not record another failure - it never reached the code
            // that would.
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxFailuresPerAccount; attempt++)
            {
                await AccountFlows.LoginAsync(
                    new LoginRequest("dennis@example.com", "wrong", null, "192.0.2.1"),
                    store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            int failuresBefore = store.AuthFailures.Count;

            await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "wrong again", null, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal(failuresBefore, store.AuthFailures.Count);
        }

        [Fact]
        public async Task Logging_in_with_a_differently_cased_address_works()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.RegisteredAsync();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("DENNIS@EXAMPLE.COM", AccountFlowTests.GoodPassword, null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);
        }

        [Fact]
        public async Task A_password_hashed_with_weaker_parameters_is_upgraded_on_login()
        {
            // The rehash-on-login path, end to end: the plaintext is legitimately in hand for this
            // one instant, so a below-policy hash is replaced.
            var store = new FakeAccountStore();
            var mailer = new FakeEmailSender();

            var weak = new ServerOptions
            {
                Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1,
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk", RefreshTokenDays = 30
            };

            await AccountFlows.RegisterAsync(
                new RegistrationRequest("dennis@example.com", AccountFlowTests.GoodPassword, "Dennis", null),
                store, mailer, new Argon2PasswordHasher(weak), weak, AccountFlowTests.Now);

            string hashBefore = store.Accounts.Values.Single().PasswordHash;

            // ServerOptions is a plain class, not a record, so this is a fresh instance rather
            // than a "with" expression.
            var strong = new ServerOptions
            {
                Argon2MemoryKib = 16384, Argon2Iterations = 2, Argon2Parallelism = 1,
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk", RefreshTokenDays = 30
            };

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, null),
                store, new Argon2PasswordHasher(strong), strong, AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);

            string hashAfter = store.Accounts.Values.Single().PasswordHash;

            Assert.NotEqual(hashBefore, hashAfter);
            Assert.Contains("m=16384", hashAfter, StringComparison.Ordinal);
        }

        private static void EmailMessageAssert(FakeEmailSender mailer, string expectedSubjectFragment)
        {
            EmailMessage message = Assert.Single(mailer.Sent);
            Assert.Contains(expectedSubjectFragment, message.Subject, StringComparison.Ordinal);
        }
    }
}
