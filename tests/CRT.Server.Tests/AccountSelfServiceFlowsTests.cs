using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers AccountSelfServiceFlows - a signed-in maintainer changing their own name, address and
    // password (owner request, 2026-10-03).
    //
    // THE ADDRESS TESTS ARE THE POINT OF THIS FILE. The address is the account's identity and where
    // "I forgot my password" sends its code, so changing it is the one operation here that could
    // hand an account to somebody else: a typo must not strand the account (a code proves the new
    // mailbox), the answer must not reveal which addresses are registered (a taken one answers like
    // a free one), and once it is done, every other session and every reset code still in the OLD
    // mailbox must stop working, and the old address is told.
    //
    // NO CURRENT PASSWORD IS ASKED FOR (owner decision, 2026-10-03: "as I see it as you are already
    // logged in") - the session is enough, an accepted risk; see AccountSelfServiceFlows' header.
    // ###########################################################################################
    public class AccountSelfServiceFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 3, 12, 0, 0, TimeSpan.Zero);

        private const string GoodPassword = "correct horse battery staple";
        private const string NewPassword = "a different horse entirely";
        private const string Ip = "192.0.2.1";

        private static ServerOptions Options() => new()
        {
            Argon2MemoryKib = 8192,
            Argon2Iterations = 1,
            Argon2Parallelism = 1,
            PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
            MailFromAddress = "noreply@classic-repair-toolbox.dk",
            RefreshTokenDays = 30
        };

        private static Argon2PasswordHasher Hasher() => new(AccountSelfServiceFlowsTests.Options());

        private sealed record Setup(FakeAccountStore Store, FakeEmailSender Mailer, string Token, long AccountId);

        // Registers and signs in one account, and empties the mail the registration sent.
        private static async Task<Setup> SignedInAsync(
            string email = "dennis@example.com",
            FakeAccountStore? store = null,
            FakeEmailSender? mailer = null,
            string displayName = "Dennis")
        {
            store ??= new FakeAccountStore();
            mailer ??= new FakeEmailSender();

            await AccountFlows.RegisterAsync(
                new RegistrationRequest(email, AccountSelfServiceFlowsTests.GoodPassword, displayName, null),
                store, mailer, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

            string token = await AccountSelfServiceFlowsTests.LogInAsync(store, email);

            mailer.Sent.Clear();

            return new Setup(store, mailer, token, store.Accounts.Values.Single(account => account.NormalisedEmail == email.ToLowerInvariant()).Id);
        }

        private static async Task<string> LogInAsync(FakeAccountStore store, string email, string password = GoodPassword)
        {
            LoginOutcome login = await AccountFlows.LoginAsync(
                new LoginRequest(email, password, "CRT 2.6", AccountSelfServiceFlowsTests.Ip),
                store, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

            Assert.True(login.IsSuccess, $"Could not sign in as {email}.");

            return login.Session!.RefreshToken;
        }

        private static Task<AccountChangeOutcome> RequestEmailAsync(Setup setup, string newEmail, string? token = null, string ip = Ip) =>
            AccountSelfServiceFlows.RequestEmailChangeAsync(
                token ?? setup.Token, newEmail, ip, setup.Store, setup.Mailer,
                AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

        private static Task<AccountChangeOutcome> ConfirmAsync(Setup setup, string? code, string? token = null, DateTimeOffset? at = null) =>
            AccountSelfServiceFlows.ConfirmEmailChangeAsync(
                token ?? setup.Token, code, setup.Store, setup.Mailer, AccountSelfServiceFlowsTests.Options(), at ?? AccountSelfServiceFlowsTests.Now);

        private static Task<AccountChangeOutcome> ChangePasswordAsync(Setup setup, string? next, string? token = null) =>
            AccountSelfServiceFlows.ChangePasswordAsync(
                token ?? setup.Token, next, AccountSelfServiceFlowsTests.Ip, setup.Store, setup.Mailer,
                AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

        private static AccountRecord Account(Setup setup) => setup.Store.Accounts[setup.AccountId];

        private static long SessionIdOf(Setup setup, string token) =>
            setup.Store.SessionHashes.Single(pair => pair.Value == SecureToken.Hash(token)).Key;

        // -----------------------------------------------------------------------------------
        // Name.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_signed_in_maintainer_changes_their_name_and_gets_the_account_back()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome outcome = await AccountSelfServiceFlows.ChangeNameAsync(
                setup.Token, "  Dennis H  ", setup.Store, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

            Assert.True(outcome.IsDone);
            Assert.Equal("Dennis H", AccountSelfServiceFlowsTests.Account(setup).DisplayName);
            Assert.Equal("Dennis H", outcome.Answer!.Account!.DisplayName);
            Assert.Equal("Your name is now Dennis H.", outcome.Answer.Message);

            AuditEntry row = Assert.Single(setup.Store.Audit, entry => entry.Action == AccountSelfServiceFlows.NameChangedAction);
            Assert.Equal("\"Dennis\" -> \"Dennis H\"", row.Detail);
        }

        [Fact]
        public async Task A_name_breaking_a_rule_is_refused_with_the_rules_and_changes_nothing()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome outcome = await AccountSelfServiceFlows.ChangeNameAsync(
                setup.Token, "D", setup.Store, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

            Assert.False(outcome.IsDone);
            Assert.NotEmpty(outcome.Errors);
            Assert.Equal("Dennis", AccountSelfServiceFlowsTests.Account(setup).DisplayName);
        }

        // Saving the name that is already there writes nothing and leaves no audit row behind.
        [Fact]
        public async Task The_same_name_again_writes_nothing()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome outcome = await AccountSelfServiceFlows.ChangeNameAsync(
                setup.Token, "Dennis", setup.Store, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);

            Assert.True(outcome.IsDone);
            Assert.DoesNotContain(setup.Store.Audit, entry => entry.Action == AccountSelfServiceFlows.NameChangedAction);
        }

        // Every change needs a LIVE session: none, a made-up one, and one signed out.
        [Fact]
        public async Task Nothing_changes_without_a_live_session()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            await AccountFlows.LogoutAsync(setup.Token, setup.Store, AccountSelfServiceFlowsTests.Now);

            foreach (string? token in new[] { null, "not-a-token", setup.Token })
            {
                Assert.True((await AccountSelfServiceFlows.ChangeNameAsync(
                    token, "Mallory", setup.Store, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now)).IsNotSignedIn);
                Assert.True((await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "mallory@example.com", token: token ?? string.Empty)).IsNotSignedIn);
                Assert.True((await AccountSelfServiceFlowsTests.ConfirmAsync(setup, "code", token: token ?? string.Empty)).IsNotSignedIn);
                Assert.True((await AccountSelfServiceFlowsTests.ChangePasswordAsync(setup, AccountSelfServiceFlowsTests.NewPassword, token: token ?? string.Empty)).IsNotSignedIn);
            }

            Assert.Equal("Dennis", AccountSelfServiceFlowsTests.Account(setup).DisplayName);
            Assert.Empty(setup.Mailer.Sent);
        }

        // -----------------------------------------------------------------------------------
        // Address.
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The whole round trip: asking changes NOTHING but sends a code to the new mailbox, and the
        // code coming back is what changes the address - which can then be signed in with.
        // ###########################################################################################
        [Fact]
        public async Task A_new_address_is_used_only_once_the_code_mailed_to_it_comes_back()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome asked = await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "New.Address@Example.com");

            Assert.True(asked.IsDone);
            Assert.True(asked.Answer!.CodeSent);
            Assert.Null(asked.Answer.Account);
            Assert.Equal(AccountSelfServiceFlows.CodeSentMessage("New.Address@Example.com"), asked.Answer.Message);
            Assert.Equal("dennis@example.com", AccountSelfServiceFlowsTests.Account(setup).Email);

            EmailMessage mail = Assert.Single(setup.Mailer.Sent);
            Assert.Equal("New.Address@Example.com", mail.ToAddress);
            Assert.Equal("Confirm your new Classic Repair Toolbox address", mail.Subject);

            AccountChangeOutcome confirmed = await AccountSelfServiceFlowsTests.ConfirmAsync(setup, setup.Mailer.ExtractTokenFromLastMail());

            Assert.True(confirmed.IsDone);
            Assert.Equal("New.Address@Example.com", confirmed.Answer!.Account!.Email);
            Assert.Equal("New.Address@Example.com", AccountSelfServiceFlowsTests.Account(setup).Email);
            Assert.Equal("new.address@example.com", AccountSelfServiceFlowsTests.Account(setup).NormalisedEmail);

            // The code reaching the mailbox proves what the verification link would have.
            Assert.True(AccountSelfServiceFlowsTests.Account(setup).IsVerified);

            await AccountSelfServiceFlowsTests.LogInAsync(setup.Store, "new.address@example.com");

            LoginOutcome old = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountSelfServiceFlowsTests.GoodPassword, null, null),
                setup.Store, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);
            Assert.False(old.IsSuccess);
        }

        // ###########################################################################################
        // *** THE SESSION IS ENOUGH (owner decision, 2026-10-03). *** No current password is asked
        // for - the requests carry none - and nothing counts as a failed login.
        // ###########################################################################################
        [Fact]
        public async Task A_new_address_needs_only_the_session()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome asked = await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "bench@example.com");

            Assert.True(asked.IsDone);
            Assert.Single(setup.Mailer.Sent);
            Assert.Empty(setup.Store.AuthFailures);
            Assert.DoesNotContain(typeof(ChangeEmailRequest).GetProperties(), property => property.Name.Contains("Password", StringComparison.Ordinal));
            Assert.DoesNotContain(typeof(ChangePasswordRequest).GetProperties(), property => property.Name.Contains("Current", StringComparison.Ordinal));
        }

        // ###########################################################################################
        // *** AN ADDRESS ANOTHER ACCOUNT HOLDS ANSWERS EXACTLY LIKE A FREE ONE (owner decision,
        // 2026-10-03). *** The difference is only the mail: the holder is told somebody tried, and
        // no code exists that could move anything.
        // ###########################################################################################
        [Fact]
        public async Task An_address_another_account_holds_gets_the_same_answer_and_its_owner_a_notice_instead_of_a_code()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            await AccountSelfServiceFlowsTests.SignedInAsync("anna@example.com", setup.Store, setup.Mailer, "Anna");

            AccountChangeOutcome taken = await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "Anna@Example.com");

            Assert.True(taken.IsDone);
            Assert.True(taken.Answer!.CodeSent);
            Assert.Equal(AccountSelfServiceFlows.CodeSentMessage("Anna@Example.com"), taken.Answer.Message);

            EmailMessage notice = Assert.Single(setup.Mailer.Sent);
            Assert.Equal("anna@example.com", notice.ToAddress);
            Assert.Equal("Someone tried to use your Classic Repair Toolbox address", notice.Subject);
            // Greeted as the HOLDER - nothing of the requester's but the address reaches this mail.
            Assert.StartsWith("Hi Anna,", notice.Body, StringComparison.Ordinal);

            Assert.DoesNotContain(setup.Store.Tokens.Values, token => token.Purpose == TokenPurpose.EmailChange);
            Assert.Equal("dennis@example.com", AccountSelfServiceFlowsTests.Account(setup).Email);
        }

        // Only the capitals changing is the same mailbox: done at once, no code and no mail.
        [Fact]
        public async Task Only_the_capitals_changing_needs_no_code()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "Dennis@Example.com");

            Assert.True(outcome.IsDone);
            Assert.False(outcome.Answer!.CodeSent);
            Assert.Equal("Dennis@Example.com", outcome.Answer.Account!.Email);
            Assert.Equal("Dennis@Example.com", AccountSelfServiceFlowsTests.Account(setup).Email);
            Assert.Empty(setup.Mailer.Sent);
        }

        // The same address, or something that is no address, is refused - and spends no mail budget.
        [Fact]
        public async Task The_same_address_or_no_address_is_refused_and_mails_nothing()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            Assert.Equal(AccountSelfServiceFlows.AlreadyYourAddress,
                (await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, " dennis@example.com ")).Message);
            Assert.Equal(AccountSelfServiceFlows.NotAnEmailAddress,
                (await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "not an address")).Message);

            Assert.Empty(setup.Mailer.Sent);
            Assert.Empty(setup.Store.MailRequests);
        }

        // ###########################################################################################
        // *** A CODE WORKS ONLY IN THE SESSION OF THE ACCOUNT IT WAS MADE FOR. *** Another
        // maintainer holding it - forwarded, or read over a shoulder - cannot move their own account
        // to that address with it, and is told no more than for a mistyped code.
        // ###########################################################################################
        [Fact]
        public async Task A_code_is_refused_in_another_accounts_session()
        {
            Setup dennis = await AccountSelfServiceFlowsTests.SignedInAsync();
            Setup anna = await AccountSelfServiceFlowsTests.SignedInAsync("anna@example.com", dennis.Store, dennis.Mailer);

            await AccountSelfServiceFlowsTests.RequestEmailAsync(dennis, "bench@example.com");
            string code = dennis.Mailer.ExtractTokenFromLastMail();

            AccountChangeOutcome stolen = await AccountSelfServiceFlowsTests.ConfirmAsync(anna, code);

            Assert.Equal(AccountSelfServiceFlows.CodeNotValid, stolen.Message);
            Assert.Equal("anna@example.com", AccountSelfServiceFlowsTests.Account(anna).Email);

            // Still good for the account it belongs to.
            Assert.True((await AccountSelfServiceFlowsTests.ConfirmAsync(dennis, code)).IsDone);
        }

        [Fact]
        public async Task A_code_works_once_and_not_after_it_expires()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "bench@example.com");
            string code = setup.Mailer.ExtractTokenFromLastMail();

            AccountChangeOutcome late = await AccountSelfServiceFlowsTests.ConfirmAsync(
                setup, code, at: AccountSelfServiceFlowsTests.Now + AccountSelfServiceFlows.EmailChangeCodeLifetime + TimeSpan.FromMinutes(1));
            Assert.Equal(AccountSelfServiceFlows.CodeExpired, late.Message);

            Assert.True((await AccountSelfServiceFlowsTests.ConfirmAsync(setup, code)).IsDone);
            Assert.Equal(AccountSelfServiceFlows.CodeAlreadyUsed, (await AccountSelfServiceFlowsTests.ConfirmAsync(setup, code)).Message);

            Assert.Equal(AccountSelfServiceFlows.PasteTheCode, (await AccountSelfServiceFlowsTests.ConfirmAsync(setup, "  ")).Message);
            Assert.Equal(AccountSelfServiceFlows.CodeNotValid, (await AccountSelfServiceFlowsTests.ConfirmAsync(setup, "made-up-code-0000000000")).Message);
        }

        // Asking again for a different address: only the newest code works, so the first address
        // cannot be confirmed after the maintainer changed their mind.
        [Fact]
        public async Task Only_the_newest_code_works()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "first@example.com");
            string first = setup.Mailer.ExtractTokenFromLastMail();

            await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "second@example.com");
            string second = setup.Mailer.ExtractTokenFromLastMail();

            Assert.False((await AccountSelfServiceFlowsTests.ConfirmAsync(setup, first)).IsDone);
            Assert.Equal("dennis@example.com", AccountSelfServiceFlowsTests.Account(setup).Email);

            Assert.True((await AccountSelfServiceFlowsTests.ConfirmAsync(setup, second)).IsDone);
            Assert.Equal("second@example.com", AccountSelfServiceFlowsTests.Account(setup).Email);
        }

        // An account registered at the address between the request and the code: refused, nothing
        // moved, and the code spent.
        [Fact]
        public async Task An_address_taken_while_the_code_was_on_its_way_is_refused()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "bench@example.com");
            string code = setup.Mailer.ExtractTokenFromLastMail();

            await AccountSelfServiceFlowsTests.SignedInAsync("bench@example.com", setup.Store, setup.Mailer);

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.ConfirmAsync(setup, code);

            Assert.Equal(AccountSelfServiceFlows.AddressTakenMeanwhile, outcome.Message);
            Assert.Equal("dennis@example.com", AccountSelfServiceFlowsTests.Account(setup).Email);
            Assert.Equal(AccountSelfServiceFlows.CodeAlreadyUsed, (await AccountSelfServiceFlowsTests.ConfirmAsync(setup, code)).Message);
        }

        // ###########################################################################################
        // *** ONCE THE ADDRESS HAS CHANGED, THE OLD WAYS IN ARE CLOSED (owner decision,
        // 2026-10-03). *** Every other session is signed out (this one kept), a reset code still
        // in the old mailbox stops working, and the old address is told - naming the new one.
        // ###########################################################################################
        [Fact]
        public async Task A_changed_address_signs_out_every_other_session_and_closes_the_old_mailbox()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            string otherComputer = await AccountSelfServiceFlowsTests.LogInAsync(setup.Store, "dennis@example.com");

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", setup.Store, setup.Mailer, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);
            string resetCode = setup.Mailer.ExtractTokenFromLastMail();

            await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "bench@example.com");
            Assert.True((await AccountSelfServiceFlowsTests.ConfirmAsync(setup, setup.Mailer.ExtractTokenFromLastMail())).IsDone);

            Assert.NotNull(await AccountFlows.AuthenticateAsync(setup.Token, setup.Store, AccountSelfServiceFlowsTests.Now));
            Assert.Null(await AccountFlows.AuthenticateAsync(otherComputer, setup.Store, AccountSelfServiceFlowsTests.Now));
            Assert.Equal("email changed", setup.Store.Sessions[AccountSelfServiceFlowsTests.SessionIdOf(setup, otherComputer)].RevokedReason);

            TokenRedemptionOutcome reset = await AccountFlows.CompletePasswordResetAsync(
                resetCode, "somebody else's password now", setup.Store, setup.Mailer, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Now);
            Assert.Equal(TokenRedemptionOutcome.AlreadyUsed, reset);

            EmailMessage notice = Assert.Single(setup.Mailer.Sent, mail => mail.Subject == "Your Classic Repair Toolbox address was changed");
            Assert.Equal("dennis@example.com", notice.ToAddress);
            Assert.Contains("bench@example.com", notice.Body, StringComparison.Ordinal);
            Assert.Contains(EmailTemplates.ContactAddress, notice.Body, StringComparison.Ordinal);

            AuditEntry row = Assert.Single(setup.Store.Audit, entry => entry.Action == AccountSelfServiceFlows.EmailChangedAction);
            Assert.Equal("dennis@example.com -> bench@example.com", row.Detail);
        }

        // The address-change mails come out of the same per-address budget registration and
        // "I forgot my password" spend - and a refusal is said, not hidden behind "a code is on its way".
        [Fact]
        public async Task Asking_for_codes_is_limited_by_the_same_mail_budget_as_registration()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();

            for (int sent = 0; sent < AuthRateLimitPolicy.MaxMailRequestsPerAddress; sent++)
                setup.Store.MailRequests.Add((AccountSelfServiceFlowsTests.Ip, AccountSelfServiceFlowsTests.Now - TimeSpan.FromMinutes(5)));

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.RequestEmailAsync(setup, "bench@example.com");

            Assert.NotNull(outcome.RetryAfter);
            Assert.Empty(setup.Mailer.Sent);
        }

        // -----------------------------------------------------------------------------------
        // Password.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_changed_password_works_the_old_one_does_not_and_other_computers_are_signed_out()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            string otherComputer = await AccountSelfServiceFlowsTests.LogInAsync(setup.Store, "dennis@example.com");

            await AccountFlows.RequestPasswordResetAsync(
                "dennis@example.com", setup.Store, setup.Mailer, AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);
            string resetCode = setup.Mailer.ExtractTokenFromLastMail();

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.ChangePasswordAsync(
                setup, AccountSelfServiceFlowsTests.NewPassword);

            Assert.True(outcome.IsDone);
            Assert.Equal(AccountSelfServiceFlows.PasswordChangedMessage, outcome.Answer!.Message);

            await AccountSelfServiceFlowsTests.LogInAsync(setup.Store, "dennis@example.com", AccountSelfServiceFlowsTests.NewPassword);

            LoginOutcome old = await AccountFlows.LoginAsync(
                new LoginRequest("dennis@example.com", AccountSelfServiceFlowsTests.GoodPassword, null, null),
                setup.Store, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Options(), AccountSelfServiceFlowsTests.Now);
            Assert.False(old.IsSuccess);

            Assert.NotNull(await AccountFlows.AuthenticateAsync(setup.Token, setup.Store, AccountSelfServiceFlowsTests.Now));
            Assert.Null(await AccountFlows.AuthenticateAsync(otherComputer, setup.Store, AccountSelfServiceFlowsTests.Now));

            // A reset code mailed before must not set the password back.
            Assert.Equal(
                TokenRedemptionOutcome.AlreadyUsed,
                await AccountFlows.CompletePasswordResetAsync(
                    resetCode, "yet another password here", setup.Store, setup.Mailer, AccountSelfServiceFlowsTests.Hasher(), AccountSelfServiceFlowsTests.Now));

            Assert.Contains(setup.Mailer.Sent, mail => mail.Subject == "Your Classic Repair Toolbox password was changed" && mail.ToAddress == "dennis@example.com");
            Assert.Contains(setup.Store.Audit, entry => entry.Action == AccountSelfServiceFlows.PasswordChangedAction);
        }

        // A new password breaking a rule is refused with every rule - before the mail budget is
        // spent or the hasher runs.
        [Fact]
        public async Task A_new_password_breaking_a_rule_is_refused_and_changes_nothing()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            string hashBefore = AccountSelfServiceFlowsTests.Account(setup).PasswordHash;

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.ChangePasswordAsync(setup, "short");

            Assert.NotEmpty(outcome.Errors);
            Assert.Equal(hashBefore, AccountSelfServiceFlowsTests.Account(setup).PasswordHash);
            Assert.Empty(setup.Store.MailRequests);
            Assert.Empty(setup.Mailer.Sent);
        }

        // ###########################################################################################
        // *** WITH NO PASSWORD CHECK IN FRONT OF IT, THE HASHER IS GUARDED BY THE MAIL BUDGET. ***
        // A password change runs the ~128 MiB Argon2 hasher and sends a mail; spent from the same
        // per-address budget as registration, it cannot be called without limit.
        // ###########################################################################################
        [Fact]
        public async Task Password_changes_spend_the_mail_budget_before_the_hasher_runs()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            string hashBefore = AccountSelfServiceFlowsTests.Account(setup).PasswordHash;

            for (int sent = 0; sent < AuthRateLimitPolicy.MaxMailRequestsPerAddress; sent++)
                setup.Store.MailRequests.Add((AccountSelfServiceFlowsTests.Ip, AccountSelfServiceFlowsTests.Now - TimeSpan.FromMinutes(5)));

            AccountChangeOutcome outcome = await AccountSelfServiceFlowsTests.ChangePasswordAsync(setup, AccountSelfServiceFlowsTests.NewPassword);

            Assert.NotNull(outcome.RetryAfter);
            Assert.Equal(hashBefore, AccountSelfServiceFlowsTests.Account(setup).PasswordHash);
            Assert.Empty(setup.Mailer.Sent);
        }

        // -----------------------------------------------------------------------------------
        // The account as /me answers it.
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task The_account_lists_the_systems_it_maintains_in_order()
        {
            Setup setup = await AccountSelfServiceFlowsTests.SignedInAsync();
            setup.Store.Maintainers.Add(("Commodore/C64/250407", setup.AccountId));
            setup.Store.Maintainers.Add(("Commodore/C128/310378", setup.AccountId));

            AccountAnswer answer = await AccountSelfServiceFlows.DescribeAsync(AccountSelfServiceFlowsTests.Account(setup), setup.Store);

            Assert.Equal(["Commodore/C128/310378", "Commodore/C64/250407"], answer.MaintainerOf);
            Assert.Equal("dennis@example.com", answer.Email);
            Assert.Equal("Dennis", answer.DisplayName);
        }
    }
}
