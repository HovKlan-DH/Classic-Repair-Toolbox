using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers verification, login and password reset end to end against fakes. (Registration went with server 17.0.0 - owner
    // decision, 2026-11-28: no accounts for everybody; an account is made by an invitation.)
    //
    // These are the flows an attacker reaches first, so the tests are written around what must NOT
    // leak as much as around what must work. The enumeration tests in particular assert that two
    // different situations produce the SAME answer - a test shape that looks odd until you
    // remember the failure it prevents: a sign-in that says "no such account" is a list of
    // everyone with one.
    //
    // Weak Argon2 parameters throughout - see PasswordHashingTests for why.
    // ###########################################################################################
    public class AccountFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private const string GoodPassword = "correct horse battery staple";

        // How long the verification link the helper below makes lasts - as the links mailed before server 17.0.0 did.
        private static readonly TimeSpan VerificationLinkLasts = TimeSpan.FromHours(48);

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

        // ###########################################################################################
        // An account with GoodPassword, as one registered before server 17.0.0 is: unconfirmed, with the verification link its
        // mail carried (the token returned) - the starting point for most of the tests below. `confirmed`: its address
        // confirmed too, as opening the link does - an address not confirmed cannot sign in (since 2026-11-25).
        // ###########################################################################################
        private static async Task<(FakeAccountStore Store, FakeEmailSender Mailer, string Token)> AccountAsync(
            string email = "dennis@example.com",
            bool confirmed = false,
            string name = "Dennis",
            FakeAccountStore? store = null,
            Argon2PasswordHasher? hasher = null)
        {
            store ??= new FakeAccountStore();

            long id = await store.CreateAccountAsync(new NewAccount(
                email, AccountRules.NormaliseEmail(email), await (hasher ?? AccountFlowTests.Hasher()).HashAsync(AccountFlowTests.GoodPassword), name, AccountFlowTests.Now));

            string token = SecureToken.Create();

            await store.CreateTokenAsync(new NewAccountToken(
                id, TokenPurpose.EmailVerification, SecureToken.Hash(token), AccountFlowTests.Now, AccountFlowTests.Now + AccountFlowTests.VerificationLinkLasts));

            if (confirmed)
                store.ConfirmAddress(email);

            return (store, new FakeEmailSender(), token);
        }

        // A reset's token is stored as its hash alone - a database read must not yield working reset links.
        [Fact]
        public async Task Only_the_token_HASH_is_stored_never_the_token_itself()
        {
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            await AccountFlows.RequestPasswordResetAsync("dennis@example.com", store, mailer, AccountFlowTests.Options(), AccountFlowTests.Now);

            string token = mailer.ExtractTokenFromLastMail();

            Assert.Contains(SecureToken.Hash(token), store.TokenHashes.Values);
            Assert.DoesNotContain(token, store.TokenHashes.Values);
        }

        // -----------------------------------------------------------------------------------
        // Verification - a link mailed before server 17.0.0 still answers.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_verification_token_from_the_mail_verifies_the_account()
        {
            (FakeAccountStore store, _, string token) = await AccountFlowTests.AccountAsync();

            TokenRedemptionOutcome outcome = await AccountFlows.VerifyEmailAsync(
                token, store, new FakeEmailSender(), AccountFlowTests.Now);

            Assert.Equal(TokenRedemptionOutcome.Success, outcome);
            Assert.True(store.Accounts.Values.Single().IsVerified);
        }

        [Fact]
        public async Task A_verification_token_cannot_be_used_twice()
        {
            // Without this, a verification link sitting in a mailbox stays live forever.
            (FakeAccountStore store, _, string token) = await AccountFlowTests.AccountAsync();

            await AccountFlows.VerifyEmailAsync(token, store, new FakeEmailSender(), AccountFlowTests.Now);

            Assert.Equal(
                TokenRedemptionOutcome.AlreadyUsed,
                await AccountFlows.VerifyEmailAsync(token, store, new FakeEmailSender(), AccountFlowTests.Now));
        }

        [Fact]
        public async Task An_expired_verification_token_is_refused()
        {
            (FakeAccountStore store, _, string token) = await AccountFlowTests.AccountAsync();

            DateTimeOffset tooLate = AccountFlowTests.Now + AccountFlowTests.VerificationLinkLasts + TimeSpan.FromMinutes(1);

            Assert.Equal(
                TokenRedemptionOutcome.Expired,
                await AccountFlows.VerifyEmailAsync(token, store, new FakeEmailSender(), tooLate));

            Assert.False(store.Accounts.Values.Single().IsVerified);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not-a-real-token")]
        public async Task A_bogus_verification_token_is_refused(string? token)
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync();

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.VerifyEmailAsync(token, store, new FakeEmailSender(), AccountFlowTests.Now));
        }

        [Fact]
        public async Task A_PASSWORD_RESET_token_cannot_be_used_to_verify_an_address()
        {
            // The purposes are separate for exactly this reason: one bug in a WHERE clause must
            // not let a reset token act as a verification, or vice versa.
            (FakeAccountStore store, FakeEmailSender mailer, _) = await AccountFlowTests.AccountAsync();

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", store, mailer, AccountFlowTests.Options(), AccountFlowTests.Now);

            string resetToken = mailer.ExtractTokenFromLastMail();

            Assert.Equal(
                TokenRedemptionOutcome.Invalid,
                await AccountFlows.VerifyEmailAsync(resetToken, store, mailer, AccountFlowTests.Now));
        }

        // -----------------------------------------------------------------------------------
        // Login.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_correct_password_logs_in_and_issues_a_session()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, "CRT/2.6", "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);
            Assert.NotNull(outcome.Session);
            Assert.NotEmpty(outcome.Session!.RefreshToken);
            Assert.Single(store.Sessions);
        }

        // ###########################################################################################
        // *** A USER AGENT CUT TO ITS COLUMN NEVER ENDS IN HALF A CHARACTER (code review, 2026-10-29). ***
        // Cut at 255 UTF-16 units, an emoji across the cut left a lone high surrogate, which MariaDB's
        // utf8mb4 refuses - the session's insert failed, and with it the sign-in. The one shared helper
        // (BoardViewRules.Clip) steps back off it.
        // ###########################################################################################
        [Fact]
        public async Task A_long_user_agent_is_cut_without_splitting_a_character()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);
            string userAgent = new string('a', 254) + "\U0001F600 and more";

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, userAgent, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(outcome.IsSuccess);
            string stored = Assert.Single(store.NewSessions.Values).UserAgent!;
            Assert.Equal(new string('a', 254), stored);
        }

        // ###########################################################################################
        // *** AN ADDRESS NOT CONFIRMED CANNOT SIGN IN, AND IT LOOKS LIKE A WRONG PASSWORD (project owner,
        // 2026-11-25: "I do not think any address should be revealed, no"). *** It could until then, for
        // "resend my verification mail" - and that made signing in the second request that told whether
        // an address had an account: register it with a password of your own, then sign in with it. The
        // right password of an unconfirmed account now fails as a wrong one does - the same answer, and a
        // failure counted for the limiter, which would otherwise trip only for an address that had one.
        // Confirmed, it signs in.
        // ###########################################################################################
        [Fact]
        public async Task An_UNCONFIRMED_account_cannot_sign_in_and_fails_as_a_wrong_password_does()
        {
            (FakeAccountStore store, _, string token) = await AccountFlowTests.AccountAsync();

            Assert.False(store.Accounts.Values.Single().IsVerified);

            LoginOutcome rightPassword = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "198.51.100.7"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            LoginOutcome wrongPassword = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "the wrong password entirely", null, "198.51.100.7"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.Equal(wrongPassword, rightPassword);
            Assert.False(rightPassword.IsSuccess);
            Assert.Empty(store.NewSessions);
            Assert.Equal(2, store.AuthFailures.Count);
            Assert.Contains(store.Audit, entry => entry.Action == AccountFlows.SignInUnconfirmedAction);

            await AccountFlows.VerifyEmailAsync(token, store, new FakeEmailSender(), AccountFlowTests.Now);

            LoginOutcome confirmed = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "198.51.100.7"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(confirmed.IsSuccess);
        }

        // ###########################################################################################
        // Two accounts left unconfirmed from when anybody could register asked for one name, and the first confirmed: the
        // second, confirming, is given the name with "-2" after it and told so by mail (AccountFlows.ConfirmAddressAsync) - a
        // confirmed name is never shared.
        // ###########################################################################################
        [Fact]
        public async Task The_second_to_confirm_a_name_is_given_it_numbered_and_told_by_mail()
        {
            (FakeAccountStore store, FakeEmailSender mailer, string annasLink) = await AccountFlowTests.AccountAsync("anna@example.com", name: "Anna");
            (_, _, string othersLink) = await AccountFlowTests.AccountAsync("other@example.com", name: "ANNA", store: store);

            Assert.Equal(TokenRedemptionOutcome.Success, await AccountFlows.VerifyEmailAsync(annasLink, store, mailer, AccountFlowTests.Now));
            int mailsBefore = mailer.Sent.Count;

            Assert.Equal(TokenRedemptionOutcome.Success, await AccountFlows.VerifyEmailAsync(othersLink, store, mailer, AccountFlowTests.Now));

            Assert.Equal("Anna", store.Accounts.Values.Single(account => account.Email == "anna@example.com").DisplayName);
            Assert.Equal("ANNA-2", store.Accounts.Values.Single(account => account.Email == "other@example.com").DisplayName);

            EmailMessage told = Assert.Single(mailer.Sent.Skip(mailsBefore));
            Assert.Equal("other@example.com", told.ToAddress);
            Assert.Contains("ANNA-2", told.Body, StringComparison.Ordinal);
            Assert.Contains(store.Audit, entry => entry.Action == AccountFlows.NameGivenAtConfirmationAction);
        }

        // A name too long for its number is cut short first, so the whole stays within the rule's 30 characters.
        [Theory]
        [InlineData("Anna", 2, "Anna-2")]
        [InlineData("123456789012345678901234567890", 2, "1234567890123456789012345678-2")]
        [InlineData("123456789012345678901234567890", 12, "123456789012345678901234567-12")]
        public void A_numbered_name_keeps_within_the_rule(string name, int number, string expected)
        {
            string given = AccountFlows.NumberedName(name, number);

            Assert.Equal(expected, given);
            Assert.Null(DisplayNames.Problem(given));
        }

        [Fact]
        public async Task A_wrong_password_and_an_UNKNOWN_ADDRESS_fail_identically()
        {
            // The enumeration guarantee on the one endpoint everybody can reach. If these two
            // answered differently - a different status, a different message, even a different
            // shape - login would be a way to discover who has an account.
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

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
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            store.Accounts[1] = store.Accounts[1] with { IsLocked = true };

            LoginOutcome locked = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, null),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.False(locked.IsSuccess);
            Assert.False(locked.IsRateLimited);
            Assert.Null(locked.Session);
        }

        // ###########################################################################################
        // *** EVERY FAILED LOGIN COSTS ONE ARGON2 RUN, WHATEVER IT FOUND (security review,
        // 2026-10-11). *** The ANSWER was already the same for an unknown address, a locked account
        // and a wrong password, but only the wrong password ran Argon2 - hundreds of milliseconds at
        // the configured cost - so the response TIME said which addresses have an account, one
        // request per address. Counted rather than timed, because a wall-clock test would be flaky.
        // ###########################################################################################
        [Theory]
        [InlineData("nobody@example.com", "the wrong password entirely", false)]
        [InlineData("dennis@example.com", "the wrong password entirely", false)]
        [InlineData("dennis@example.com", AccountFlowTests.GoodPassword, true)]
        public async Task Every_failed_login_costs_exactly_one_password_verification(
            string email, string password, bool locked)
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            if (locked)
                store.Accounts[1] = store.Accounts[1] with { IsLocked = true };

            Argon2PasswordHasher hasher = AccountFlowTests.Hasher();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest(email, password, null, "192.0.2.1"),
                store, hasher, AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.False(outcome.IsSuccess);
            Assert.False(outcome.IsRateLimited);
            Assert.Equal(1, hasher.Computations);
        }

        // An account whose stored hash cannot be read fails like any other - and costs the same.
        [Fact]
        public async Task A_login_to_an_account_with_an_unreadable_stored_hash_costs_one_verification()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            store.Accounts[1] = store.Accounts[1] with { PasswordHash = "not a password hash" };

            Argon2PasswordHasher hasher = AccountFlowTests.Hasher();

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "192.0.2.1"),
                store, hasher, AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.False(outcome.IsSuccess);
            Assert.Equal(1, hasher.Computations);
        }

        // ###########################################################################################
        // *** AFTER THE COST IS RAISED, A WRONG PASSWORD STILL COSTS WHAT AN UNKNOWN ADDRESS COSTS
        // (code review, 2026-10-11). *** A hash keeps the cost it was made with until its account
        // signs in again, while an unknown address costs the configured one - so every account not
        // yet rehashed answered a wrong password sooner, and could be told apart by the clock.
        // Counted as Argon2 work (memory x passes over lanes), not timed.
        // ###########################################################################################
        [Fact]
        public async Task After_the_cost_is_raised_a_wrong_password_costs_the_same_work_as_an_unknown_address()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            ServerOptions raised = AccountFlowTests.Options();
            raised.Argon2MemoryKib = 16384;
            raised.Argon2Iterations = 2;

            var forUnknownAddress = new Argon2PasswordHasher(raised);
            var forOldAccount = new Argon2PasswordHasher(raised);

            await AccountFlows.LoginAsync(
                new LoginRequest("nobody@example.com", "the wrong password entirely", null, "192.0.2.1"),
                store, forUnknownAddress, raised, AccountFlowTests.Now);

            LoginOutcome outcome = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "the wrong password entirely", null, "192.0.2.2"),
                store, forOldAccount, raised, AccountFlowTests.Now);

            Assert.False(outcome.IsSuccess);
            Assert.Equal(Argon2PasswordHasher.Cost(16384, 2, 1), forUnknownAddress.Work);
            Assert.Equal(forUnknownAddress.Work, forOldAccount.Work);
        }

        // -----------------------------------------------------------------------------------
        // The mail-sending budget.
        //
        // "Forgot my password" sends a message to an address the CALLER names, so without a limit
        // this service is a way to flood a stranger's inbox under its own reputation. These pin that
        // the limit is actually applied, which AuthRateLimitPolicy's own tests cannot - the policy
        // was once written, tested and then wired to nothing at all.
        // -----------------------------------------------------------------------------------

        // `count` accounts, victim0@ to victim(N-1)@, each of which a reset mails.
        private static async Task<(FakeAccountStore Store, FakeEmailSender Mailer)> VictimsAsync(int count)
        {
            var store = new FakeAccountStore();

            for (int i = 0; i < count; i++)
                await AccountFlowTests.AccountAsync($"victim{i}@example.com", name: $"Victim{i}", store: store);

            return (store, new FakeEmailSender());
        }

        private static Task ResetAsync(FakeAccountStore store, FakeEmailSender mailer, string email, string ip, DateTimeOffset? when = null) =>
            AccountFlows.RequestPasswordResetAsync(email, store, mailer, AccountFlowTests.Options(), when ?? AccountFlowTests.Now, ip);

        [Fact]
        public async Task Reset_requests_stop_sending_mail_once_the_per_IP_budget_is_spent()
        {
            // One more than the budget, every one from the same address and each naming a
            // DIFFERENT victim - the shape inbox flooding actually takes.
            (FakeAccountStore store, FakeEmailSender mailer) = await AccountFlowTests.VictimsAsync(AuthRateLimitPolicy.MaxMailRequestsPerAddress + 1);

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
                await AccountFlowTests.ResetAsync(store, mailer, $"victim{attempt}@example.com", "192.0.2.9");

            Assert.Equal(AuthRateLimitPolicy.MaxMailRequestsPerAddress, mailer.Sent.Count);
        }

        [Fact]
        public async Task An_INVALID_reset_request_does_not_spend_the_mail_budget()
        {
            // The limit sits after validation deliberately: a malformed request sends nothing, so
            // charging it would let a clumsy client lock itself out.
            (FakeAccountStore store, FakeEmailSender mailer) = await AccountFlowTests.VictimsAsync(1);

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxMailRequestsPerAddress * 2; attempt++)
                await AccountFlowTests.ResetAsync(store, mailer, "not-an-address", "192.0.2.9");

            Assert.Empty(store.MailRequests);

            // A valid one afterwards still gets through.
            await AccountFlowTests.ResetAsync(store, mailer, "victim0@example.com", "192.0.2.9");

            Assert.Single(mailer.Sent);
        }

        [Fact]
        public async Task The_mail_budget_is_per_ADDRESS_so_one_caller_cannot_block_everybody()
        {
            // A single global bucket would mean one abuser stops every other person on the
            // internet from resetting a password - the limit becoming the denial of service.
            (FakeAccountStore store, FakeEmailSender mailer) = await AccountFlowTests.VictimsAsync(AuthRateLimitPolicy.MaxMailRequestsPerAddress + 2);

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
                await AccountFlowTests.ResetAsync(store, mailer, $"victim{attempt}@example.com", "192.0.2.9");

            int spent = mailer.Sent.Count;

            await AccountFlowTests.ResetAsync(store, mailer, $"victim{AuthRateLimitPolicy.MaxMailRequestsPerAddress + 1}@example.com", "198.51.100.4");

            Assert.Equal(spent + 1, mailer.Sent.Count);
        }

        [Fact]
        public async Task The_mail_budget_EXPIRES_rather_than_locking_an_address_out_for_ever()
        {
            // The same reasoning as the login lockout: a limit that never clears is a permanent
            // ban on whatever shares that address - a household, a whole country behind CGNAT.
            (FakeAccountStore store, FakeEmailSender mailer) = await AccountFlowTests.VictimsAsync(AuthRateLimitPolicy.MaxMailRequestsPerAddress + 2);

            for (int attempt = 0; attempt <= AuthRateLimitPolicy.MaxMailRequestsPerAddress; attempt++)
                await AccountFlowTests.ResetAsync(store, mailer, $"victim{attempt}@example.com", "192.0.2.9");

            int spent = mailer.Sent.Count;

            await AccountFlowTests.ResetAsync(
                store,
                mailer,
                $"victim{AuthRateLimitPolicy.MaxMailRequestsPerAddress + 1}@example.com",
                "192.0.2.9",
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
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

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
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxFailuresPerAccountAndAddress; attempt++)
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

        // A stranger's wrong passwords lock out the stranger's address, never the owner signing in
        // from their own (code review, 2026-10-13) - the administrator's address is in every mail.
        [Fact]
        public async Task Failures_from_another_address_do_not_stop_the_owner_signing_in()
        {
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxFailuresPerAccountAndAddress; attempt++)
            {
                await AccountFlows.LoginAsync(
                    new LoginRequest("dennis@example.com", "wrong", null, "198.51.100.66"),
                    store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);
            }

            LoginOutcome stranger = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", "wrong again", null, "198.51.100.66"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            LoginOutcome owner = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountFlowTests.GoodPassword, null, "192.0.2.1"),
                store, AccountFlowTests.Hasher(), AccountFlowTests.Options(), AccountFlowTests.Now);

            Assert.True(stranger.IsRateLimited);
            Assert.True(owner.IsSuccess);
        }

        [Fact]
        public async Task A_successful_login_clears_the_failure_count()
        {
            // Someone who gets it right on the fourth try is not an attacker and must not be
            // locked out afterwards.
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

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
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

            for (int attempt = 0; attempt < AuthRateLimitPolicy.MaxFailuresPerAccountAndAddress; attempt++)
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
            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true);

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
            var weak = new ServerOptions
            {
                Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1,
                PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
                MailFromAddress = "noreply@classic-repair-toolbox.dk", RefreshTokenDays = 30
            };

            (FakeAccountStore store, _, _) = await AccountFlowTests.AccountAsync(confirmed: true, hasher: new Argon2PasswordHasher(weak));

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
    }
}
