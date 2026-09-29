using CRT.Server.Configuration;
using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers MaintainerInvitationFlows - inviting a new maintainer by email (owner request,
    // 2026-09-27: "I should be able to either select an existing maintainer or invite a new
    // maintainer via email").
    //
    // The whole round trip is driven the way the people involved drive it: the administrator
    // invites, the code is taken OUT OF THE MAIL (never out of the store - the store holds only its
    // hash), the invitee accepts it with a name and a password, and then signs in and has authority
    // over the system.
    // ###########################################################################################
    public sealed class MaintainerInvitationFlowsTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 27, 12, 0, 0, TimeSpan.Zero);

        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";
        private const string GoodPassword = "correct horse battery staple";

        private static readonly IReadOnlyList<PublishedSystemLister.KnownSystem> Tree =
        [
            new(MaintainerInvitationFlowsTests.C64, "Commodore", "C64", "250407"),
            new(MaintainerInvitationFlowsTests.C128, "Commodore", "C128", "310378")
        ];

        private static ServerOptions Options() => new()
        {
            Argon2MemoryKib = 8192,
            Argon2Iterations = 1,
            Argon2Parallelism = 1,
            PublicApiBaseUrl = "https://classic-repair-toolbox.dk/api",
            MailFromAddress = "noreply@classic-repair-toolbox.dk",
            RefreshTokenDays = 30
        };

        private static Argon2PasswordHasher Hasher() => new(MaintainerInvitationFlowsTests.Options());

        private sealed record World(FakeAccountStore Accounts, FakeSubmissionStore Submissions, FakeEmailSender Mailer, ReviewAccess Admin, long AdminId);

        private static async Task<World> SetUpAsync()
        {
            var accounts = new FakeAccountStore();

            long adminId = await accounts.CreateAccountAsync(new NewAccount(
                "admin@example.com", "admin@example.com", "hash", "Dennis", MaintainerInvitationFlowsTests.Now));
            accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };

            return new World(accounts, new FakeSubmissionStore(), new FakeEmailSender(), ReviewAccess.For(accounts.Accounts[adminId]), adminId);
        }

        private static Task<MaintainerInvitationOutcome> InviteAsync(World world, string email, string systemId = C64, ReviewAccess? actor = null) =>
            MaintainerInvitationFlows.InviteAsync(
                actor ?? world.Admin, systemId, email, MaintainerInvitationFlowsTests.Tree,
                world.Accounts, world.Submissions, world.Mailer, MaintainerInvitationFlowsTests.Now);

        private static Task<InvitationAcceptOutcome> AcceptAsync(World world, string code, DateTimeOffset? at = null, string name = "Anna", string password = GoodPassword) =>
            MaintainerInvitationFlows.AcceptAsync(
                code, name, password, world.Accounts, MaintainerInvitationFlowsTests.Hasher(), at ?? MaintainerInvitationFlowsTests.Now.AddHours(1));

        // -----------------------------------------------------------------------------------
        // The round trip
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task An_invited_person_accepts_the_mailed_code_signs_in_and_maintains_the_system()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            MaintainerInvitationOutcome invited = await MaintainerInvitationFlowsTests.InviteAsync(world, "Anna@Example.com");

            Assert.True(invited.IsDone, invited.Message);
            Assert.Equal("An invitation is on its way to Anna@Example.com. Its code works for 14 days.", invited.Message);

            // The mail names the system, who invited, where the app is and the button to press.
            EmailMessage mail = world.Mailer.Last!;
            Assert.Equal("Anna@Example.com", mail.ToAddress);
            Assert.Equal("You are invited to maintain Commodore / C64 / 250407", mail.Subject);
            Assert.Contains("Dennis has invited you", mail.Body, StringComparison.Ordinal);
            Assert.Contains("\"I have an invitation\"", mail.Body, StringComparison.Ordinal);
            // CRT itself, and how to show its Maintainer tab (2026-09-29) - the separate CRT
            // Maintainer application and its release repository are gone.
            Assert.Contains("HovKlan-DH/Classic-Repair-Toolbox/releases", mail.Body, StringComparison.Ordinal);
            Assert.Contains("\"Enable Maintainer tab\"", mail.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("Classic-Repair-Toolbox-Maintainer", mail.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("CRT Maintainer", mail.Body, StringComparison.Ordinal);

            string code = world.Mailer.ExtractTokenFromLastMail();

            InvitationAcceptOutcome accepted = await MaintainerInvitationFlowsTests.AcceptAsync(world, code);

            Assert.True(accepted.IsAccepted, accepted.Message);
            Assert.Equal("Anna@Example.com", accepted.Answer!.Email);
            Assert.Equal([MaintainerInvitationFlowsTests.C64], accepted.Answer.SystemIds);
            Assert.Equal(
                "Your account is ready, and you now maintain Commodore/C64/250407. Sign in with your email address and the password you chose.",
                accepted.Message);

            // A VERIFIED account - the code reaching the mailbox proved it - that can sign in...
            AccountRecord anna = world.Accounts.Accounts.Values.Single(account => account.Email == "Anna@Example.com");
            Assert.True(anna.IsVerified);
            Assert.Equal("Anna", anna.DisplayName);

            LoginOutcome login = await AccountFlows.LoginAsync(
                new LoginRequest("anna@example.com", MaintainerInvitationFlowsTests.GoodPassword, "CRT 3.0.0", "192.0.2.1"),
                world.Accounts, MaintainerInvitationFlowsTests.Hasher(), MaintainerInvitationFlowsTests.Options(), MaintainerInvitationFlowsTests.Now.AddHours(2));

            Assert.True(login.IsSuccess);

            // ...and publishes to exactly the system she was invited to.
            ReviewAccess access = ReviewAccess.For(anna, await world.Accounts.GetReviewedSystemIdsAsync(anna.Id));
            Assert.True(ReviewAuthority.CanPublish(access, MaintainerInvitationFlowsTests.C64));
            Assert.False(ReviewAuthority.CanPublish(access, MaintainerInvitationFlowsTests.C128));
        }

        // *** ONLY THE HASH IS STORED *** - a database read must not yield a working invitation.
        [Fact]
        public async Task The_code_itself_is_never_stored()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");

            string code = world.Mailer.ExtractTokenFromLastMail();

            Assert.Equal(SecureToken.Hash(code), Assert.Single(world.Accounts.InvitationHashes.Values));
            Assert.DoesNotContain(world.Accounts.InvitationHashes.Values, hash => hash == code);
        }

        // One account answers every invitation its address holds.
        [Fact]
        public async Task Accepting_one_code_accepts_every_open_invitation_for_that_address()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com", MaintainerInvitationFlowsTests.C64);
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com", MaintainerInvitationFlowsTests.C128);

            InvitationAcceptOutcome accepted = await MaintainerInvitationFlowsTests.AcceptAsync(world, world.Mailer.ExtractTokenFromLastMail());

            Assert.True(accepted.IsAccepted, accepted.Message);
            Assert.Equal([MaintainerInvitationFlowsTests.C128, MaintainerInvitationFlowsTests.C64], accepted.Answer!.SystemIds.Order(StringComparer.Ordinal));

            long anna = world.Accounts.Accounts.Values.Single(account => account.Email == "anna@example.com").Id;
            Assert.Contains((MaintainerInvitationFlowsTests.C64, anna), world.Accounts.Maintainers);
            Assert.Contains((MaintainerInvitationFlowsTests.C128, anna), world.Accounts.Maintainers);
            Assert.Empty(await world.Accounts.ListOpenInvitationsAsync(MaintainerInvitationFlowsTests.Now.AddHours(2)));
        }

        // -----------------------------------------------------------------------------------
        // Inviting
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task Only_an_administrator_can_invite()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            long maintainerId = await world.Accounts.CreateAccountAsync(new NewAccount(
                "m@example.com", "m@example.com", "hash", "M", MaintainerInvitationFlowsTests.Now));
            world.Accounts.Accounts[maintainerId] = world.Accounts.Accounts[maintainerId] with { IsVerified = true };
            await world.Accounts.AddMaintainerAsync(MaintainerInvitationFlowsTests.C64, maintainerId, world.AdminId, MaintainerInvitationFlowsTests.Now);

            ReviewAccess maintainer = ReviewAccess.For(world.Accounts.Accounts[maintainerId], [MaintainerInvitationFlowsTests.C64]);

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com", actor: maintainer);

            Assert.True(outcome.IsForbidden);
            Assert.Empty(world.Mailer.Sent);
            Assert.Empty(world.Accounts.Invitations);
        }

        // ###########################################################################################
        // *** AN ADDRESS WITH AN ACCOUNT IS CHOSEN FROM THE LIST, NOT INVITED. *** The list says why
        // an account cannot be granted (unverified, locked, administrator); an invitation would be a
        // second way into the same pool that skipped those checks.
        // ###########################################################################################
        [Fact]
        public async Task An_address_that_already_has_an_account_is_not_invited()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlowsTests.InviteAsync(world, "ADMIN@example.com");

            Assert.False(outcome.IsDone);
            Assert.Equal("ADMIN@example.com already has an account - choose it from the list of accounts instead of inviting it.", outcome.Message);
            Assert.Empty(world.Mailer.Sent);
        }

        [Theory]
        [InlineData("")]
        [InlineData("not an address")]
        public async Task Something_that_is_not_an_address_is_refused(string email)
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlowsTests.InviteAsync(world, email);

            Assert.Equal("That does not look like an email address.", outcome.Message);
            Assert.Empty(world.Mailer.Sent);
        }

        [Fact]
        public async Task A_system_that_does_not_exist_is_refused()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            MaintainerInvitationOutcome outcome = await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com", "Commodore/VIC-20/999999");

            Assert.True(outcome.IsNotFound);
            Assert.Empty(world.Mailer.Sent);
        }

        // A shipped board nobody has submitted to has no systems row; the invitation's key needs one.
        [Fact]
        public async Task Inviting_to_a_board_with_no_systems_row_creates_the_row()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");

            Assert.True(world.Submissions.Systems.ContainsKey(MaintainerInvitationFlowsTests.C64));
        }

        // "Send it again": the earlier code stops working, so only the newest mail is good.
        [Fact]
        public async Task Inviting_the_same_address_again_replaces_the_open_invitation()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();

            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            string first = world.Mailer.ExtractTokenFromLastMail();

            await MaintainerInvitationFlowsTests.InviteAsync(world, "Anna@example.com");
            string second = world.Mailer.ExtractTokenFromLastMail();

            Assert.Single(await world.Accounts.ListOpenInvitationsAsync(MaintainerInvitationFlowsTests.Now));

            InvitationAcceptOutcome old = await MaintainerInvitationFlowsTests.AcceptAsync(world, first);
            Assert.Equal("That invitation has been withdrawn. Ask the administrator for a new one.", old.Message);

            Assert.True((await MaintainerInvitationFlowsTests.AcceptAsync(world, second)).IsAccepted);
        }

        [Fact]
        public async Task A_withdrawn_invitation_can_no_longer_be_accepted()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            string code = world.Mailer.ExtractTokenFromLastMail();
            long id = Assert.Single(world.Accounts.Invitations.Keys);

            MaintainerInvitationOutcome withdrawn = await MaintainerInvitationFlows.WithdrawAsync(
                world.Admin, id, world.Accounts, MaintainerInvitationFlowsTests.Now);

            Assert.True(withdrawn.IsDone, withdrawn.Message);
            Assert.Equal("The invitation to anna@example.com is withdrawn - its code no longer works.", withdrawn.Message);

            InvitationAcceptOutcome accepted = await MaintainerInvitationFlowsTests.AcceptAsync(world, code);
            Assert.False(accepted.IsAccepted);
            Assert.DoesNotContain(world.Accounts.Accounts.Values, account => account.Email == "anna@example.com");

            // Withdrawing it again says why nothing happened.
            Assert.False((await MaintainerInvitationFlows.WithdrawAsync(world.Admin, id, world.Accounts, MaintainerInvitationFlowsTests.Now)).IsDone);
        }

        // -----------------------------------------------------------------------------------
        // Accepting
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task An_expired_or_used_or_unknown_code_is_refused_with_what_to_do_next()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            string code = world.Mailer.ExtractTokenFromLastMail();

            InvitationAcceptOutcome expired = await MaintainerInvitationFlowsTests.AcceptAsync(
                world, code, MaintainerInvitationFlowsTests.Now + MaintainerInvitationFlows.InvitationLifetime);
            Assert.Equal("That invitation has expired. Ask the administrator for a new one.", expired.Message);

            InvitationAcceptOutcome unknown = await MaintainerInvitationFlowsTests.AcceptAsync(world, "not-the-code");
            Assert.StartsWith("That invitation code is not valid.", unknown.Message, StringComparison.Ordinal);

            Assert.True((await MaintainerInvitationFlowsTests.AcceptAsync(world, code)).IsAccepted);

            InvitationAcceptOutcome used = await MaintainerInvitationFlowsTests.AcceptAsync(world, code);
            Assert.Equal("That invitation has already been used. Sign in with your email address and password.", used.Message);
        }

        // The name and password rules registration keeps - every broken one listed.
        [Fact]
        public async Task A_weak_password_or_blank_name_is_refused_before_anything_is_written()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            string code = world.Mailer.ExtractTokenFromLastMail();

            InvitationAcceptOutcome outcome = await MaintainerInvitationFlowsTests.AcceptAsync(world, code, name: " ", password: "short");

            Assert.False(outcome.IsAccepted);
            Assert.True(outcome.Errors.Count >= 2);
            Assert.DoesNotContain(world.Accounts.Accounts.Values, account => account.Email == "anna@example.com");
            Assert.Single(await world.Accounts.ListOpenInvitationsAsync(MaintainerInvitationFlowsTests.Now));
        }

        // An account made for the address since it was invited (by hand) - the code is not a second
        // way into it.
        [Fact]
        public async Task A_code_for_an_address_that_has_since_got_an_account_makes_no_second_one()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            string code = world.Mailer.ExtractTokenFromLastMail();

            await world.Accounts.CreateAccountAsync(new NewAccount(
                "anna@example.com", "anna@example.com", "hash", "Anna", MaintainerInvitationFlowsTests.Now));

            InvitationAcceptOutcome outcome = await MaintainerInvitationFlowsTests.AcceptAsync(world, code);

            Assert.Equal(MaintainerInvitationRules.AlreadyHasAccountMessage, outcome.Message);
            Assert.Single(world.Accounts.Accounts.Values, account => account.Email == "anna@example.com");
            Assert.Empty(world.Accounts.Maintainers);
        }

        [Fact]
        public void The_accepted_message_names_every_system()
        {
            Assert.Equal(
                "Your account is ready, and you now maintain A/B/C, D/E/F and G/H/I. Sign in with your email address and the password you chose.",
                MaintainerInvitationRules.AcceptedMessage(["A/B/C", "D/E/F", "G/H/I"]));
        }

        // -----------------------------------------------------------------------------------
        // On the Systems screen
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // The open invitations ride on a system's detail - for the ADMINISTRATOR only, the one person
        // who can act on them. A maintainer reading the same system gets none.
        // ###########################################################################################
        [Fact]
        public async Task A_systems_detail_carries_its_open_invitations_for_the_administrator_only()
        {
            World world = await MaintainerInvitationFlowsTests.SetUpAsync();
            await MaintainerInvitationFlowsTests.InviteAsync(world, "anna@example.com");
            await MaintainerInvitationFlowsTests.InviteAsync(world, "bo@example.com", MaintainerInvitationFlowsTests.C128);

            SystemOverviewOutcome forAdmin = await SystemOverviewFlow.DetailAsync(
                world.Admin, MaintainerInvitationFlowsTests.C64, MaintainerInvitationFlowsTests.Tree, null,
                world.Submissions, world.Accounts, now: MaintainerInvitationFlowsTests.Now);

            MaintainerInvitationEntry invitation = Assert.Single(forAdmin.Detail!.Invitations!);
            Assert.Equal("anna@example.com", invitation.Email);
            Assert.Equal(MaintainerInvitationFlowsTests.Now + MaintainerInvitationFlows.InvitationLifetime, invitation.ExpiresUtc);

            long maintainerId = await world.Accounts.CreateAccountAsync(new NewAccount(
                "m@example.com", "m@example.com", "hash", "M", MaintainerInvitationFlowsTests.Now));
            world.Accounts.Accounts[maintainerId] = world.Accounts.Accounts[maintainerId] with { IsVerified = true };

            SystemOverviewOutcome forMaintainer = await SystemOverviewFlow.DetailAsync(
                ReviewAccess.For(world.Accounts.Accounts[maintainerId], [MaintainerInvitationFlowsTests.C64]),
                MaintainerInvitationFlowsTests.C64, MaintainerInvitationFlowsTests.Tree, null,
                world.Submissions, world.Accounts, now: MaintainerInvitationFlowsTests.Now);

            Assert.Null(forMaintainer.Detail!.Invitations);
        }
    }
}
