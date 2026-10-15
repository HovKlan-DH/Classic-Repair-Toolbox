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
    // Covers NewBoardMaintainerFlow - a NEW board's creator becomes its first maintainer when the
    // board is published (owner request, 2026-10-13: "make the contributor of a NEW board then
    // first/initial board maintainer also").
    //
    // The three ways it can go are driven as they happen: an account is put straight in the pool;
    // no account gets the administrator's own kind of invitation, whose code - taken OUT OF THE MAIL
    // - makes the maintainer when accepted; an account that may not maintain is left alone, and the
    // approver is told.
    // ###########################################################################################
    public sealed class NewBoardMaintainerFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 13, 12, 0, 0, TimeSpan.Zero);

        private const string NewBoard = "Commodore/C64/250469";

        private sealed record World(FakeAccountStore Accounts, FakeSubmissionStore Submissions, FakeEmailSender Mailer, ReviewAccess Admin, long AdminId);

        private static async Task<World> SetUpAsync()
        {
            var accounts = new FakeAccountStore();

            long adminId = await accounts.CreateAccountAsync(new NewAccount(
                "admin@example.com", "admin@example.com", "hash", "Dennis", NewBoardMaintainerFlowTests.Now));
            accounts.Accounts[adminId] = accounts.Accounts[adminId] with { IsAdministrator = true, IsVerified = true };

            var submissions = new FakeSubmissionStore();

            // The board's row, which the submission created.
            await submissions.EnsureBoardAsync(
                NewBoardMaintainerFlowTests.NewBoard, "Commodore", "C64", "250469",
                BoardDescriptorRules.BoardOrigin.Contributed, NewBoardMaintainerFlowTests.Now);

            return new World(accounts, submissions, new FakeEmailSender(), ReviewAccess.For(accounts.Accounts[adminId]), adminId);
        }

        private static async Task<long> AccountAsync(World world, string email, string name, bool verified = true, bool locked = false)
        {
            long id = await world.Accounts.CreateAccountAsync(new NewAccount(email, email.ToLowerInvariant(), "hash", name, NewBoardMaintainerFlowTests.Now));
            world.Accounts.Accounts[id] = world.Accounts.Accounts[id] with { IsVerified = verified, IsLocked = locked };
            return id;
        }

        private static SubmissionRecord Submission(long? accountId, string? contactEmail) =>
            new(
                Id: 7,
                BoardId: NewBoardMaintainerFlowTests.NewBoard,
                AccountId: accountId,
                ContactEmail: contactEmail,
                UploadTokenHash: null,
                BaseRevision: string.Empty,
                State: SubmissionState.Merged,
                Summary: "A new board",
                FormatVersion: SubmissionFormat.CurrentVersion,
                CreatedUtc: NewBoardMaintainerFlowTests.Now.AddDays(-1),
                ExpiresUtc: null,
                DecidedUtc: NewBoardMaintainerFlowTests.Now);

        private static Task<string> RunAsync(World world, SubmissionRecord submission) =>
            NewBoardMaintainerFlow.MakeCreatorMaintainerAsync(
                submission, world.Admin, world.Accounts, world.Submissions, world.Mailer, NewBoardMaintainerFlowTests.Now);

        // -----------------------------------------------------------------------------------
        // An account
        // -----------------------------------------------------------------------------------

        [Fact]
        public async Task A_creator_who_sent_it_signed_in_is_made_its_maintainer_and_told_so()
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();
            long anna = await NewBoardMaintainerFlowTests.AccountAsync(world, "anna@example.com", "Anna");

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(anna, contactEmail: null));

            Assert.Equal("Anna (anna@example.com), who created this new board, is now its maintainer.", note);
            Assert.Contains((NewBoardMaintainerFlowTests.NewBoard, anna), world.Accounts.Maintainers);

            // She can now publish to that board.
            ReviewAccess access = ReviewAccess.For(world.Accounts.Accounts[anna], await world.Accounts.GetReviewedBoardIdsAsync(anna));
            Assert.True(ReviewAuthority.CanPublish(access, NewBoardMaintainerFlowTests.NewBoard));

            // The mail: greeted by name, the board named, how to get to the work.
            EmailMessage mail = Assert.Single(world.Mailer.Sent);
            Assert.Equal("anna@example.com", mail.ToAddress);
            Assert.Equal("For CRT you are now the maintainer of the [Commodore / C64 / 250469] board", mail.Subject);
            Assert.StartsWith("Hi Anna,", mail.Body, StringComparison.Ordinal);
            Assert.Contains("as you agreed when you created it, you are now its maintainer", mail.Body, StringComparison.Ordinal);
            Assert.Contains("\"Enable Maintainer tab\"", mail.Body, StringComparison.Ordinal);

            // The board's history reads it as the administrator's own grant does - by the approver.
            AuditEntry granted = Assert.Single(world.Accounts.Audit);
            Assert.Equal(MaintainerAssignmentFlows.GrantedAction, granted.Action);
            Assert.Equal(NewBoardMaintainerFlowTests.NewBoard, granted.Subject);
            Assert.Equal(world.AdminId, granted.ActorAccountId);
            Assert.Equal($"account {anna} (anna@example.com)", granted.Detail);
        }

        // ###########################################################################################
        // *** A TYPED ADDRESS ALONE MAKES NOBODY MAINTAINER (code review, 2026-10-13). *** Anybody can
        // type anybody's address into the Submit dialog; nothing proves the sender owns it, and the
        // invitation that would prove it is refused for an address with an account. Nobody is added
        // or mailed, and the approver is told - whatever the capitals.
        // ###########################################################################################
        [Fact]
        public async Task A_typed_address_that_has_an_account_makes_nobody_maintainer_and_the_approver_is_told()
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();
            long anna = await NewBoardMaintainerFlowTests.AccountAsync(world, "anna@example.com", "Anna");

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(null, " Anna@Example.com "));

            Assert.Equal(NewBoardMaintainerWording.TypedAddressHasAccount("Anna@Example.com"), note);
            Assert.StartsWith("This new board was sent without signing in, giving Anna@Example.com", note, StringComparison.Ordinal);
            Assert.DoesNotContain((NewBoardMaintainerFlowTests.NewBoard, anna), world.Accounts.Maintainers);
            Assert.Empty(world.Accounts.Invitations);
            Assert.Empty(world.Mailer.Sent);
            Assert.Empty(world.Accounts.Audit);
        }

        // A board pushed back out of BETA and approved again creates it again: nothing new happens.
        [Fact]
        public async Task A_creator_who_already_maintains_it_is_not_granted_or_mailed_again()
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();
            long anna = await NewBoardMaintainerFlowTests.AccountAsync(world, "anna@example.com", "Anna");
            await world.Accounts.AddMaintainerAsync(NewBoardMaintainerFlowTests.NewBoard, anna, world.AdminId, NewBoardMaintainerFlowTests.Now);

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(anna, null));

            Assert.Equal("Anna (anna@example.com), who created this new board, already maintains it.", note);
            Assert.Empty(world.Mailer.Sent);
            Assert.Empty(world.Accounts.Audit);
        }

        // MaintainerAssignmentRules' own refusals: an account that cannot review is not put in a pool.
        [Theory]
        [InlineData(false, false, "anna@example.com created this new board, but that account's address is not verified yet, so it was not made the board's maintainer. Add it under \"Account\", \"Maintainers\", once it is.")]
        [InlineData(true, true, "anna@example.com created this new board, but that account is locked, so it was not made the board's maintainer.")]
        public async Task An_account_that_may_not_maintain_is_left_alone_and_the_approver_is_told(bool verified, bool locked, string expected)
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();
            long anna = await NewBoardMaintainerFlowTests.AccountAsync(world, "anna@example.com", "Anna", verified, locked);

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(anna, null));

            Assert.Equal(expected, note);
            Assert.DoesNotContain((NewBoardMaintainerFlowTests.NewBoard, anna), world.Accounts.Maintainers);
            Assert.Empty(world.Mailer.Sent);
            Assert.Empty(world.Accounts.Invitations);
        }

        // -----------------------------------------------------------------------------------
        // No account
        // -----------------------------------------------------------------------------------

        // The round trip: the invitation, worded for the board's creator, and its code - out of the
        // mail - makes the maintainer.
        [Fact]
        public async Task A_creator_with_no_account_is_invited_and_maintains_the_board_once_the_code_is_used()
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(null, "Bob@Example.com"));

            Assert.Equal("Bob@Example.com, who created this new board, has no account, so an invitation to become its maintainer was mailed there.", note);
            Assert.Empty(world.Accounts.Maintainers);

            EmailMessage mail = Assert.Single(world.Mailer.Sent);
            Assert.Equal("Bob@Example.com", mail.ToAddress);
            Assert.Equal("For CRT you are invited to maintain the [Commodore / C64 / 250469] board", mail.Subject);
            Assert.Contains("The new board you created, ", mail.Body, StringComparison.Ordinal);
            Assert.Contains("As you agreed when you created it, you are invited to be its maintainer", mail.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("has invited you", mail.Body, StringComparison.Ordinal);
            Assert.True(mail.CarriesCredential);

            AuditEntry invited = Assert.Single(world.Accounts.Audit);
            Assert.Equal(MaintainerInvitationFlows.InvitedAction, invited.Action);
            Assert.Equal(NewBoardMaintainerFlowTests.NewBoard, invited.Subject);
            Assert.Equal("Bob@Example.com", invited.Detail);

            InvitationAcceptOutcome accepted = await MaintainerInvitationFlows.AcceptAsync(
                world.Mailer.ExtractTokenFromLastMail(), "Bob", "correct horse battery staple", world.Accounts,
                new Argon2PasswordHasher(new ServerOptions { Argon2MemoryKib = 8192, Argon2Iterations = 1, Argon2Parallelism = 1 }),
                NewBoardMaintainerFlowTests.Now.AddHours(1));

            Assert.True(accepted.IsAccepted, accepted.Message);
            Assert.Equal([NewBoardMaintainerFlowTests.NewBoard], accepted.Answer!.BoardIds);
        }

        // Approved again (pushed back and re-approved): only the newest code works.
        [Fact]
        public async Task A_second_invitation_for_the_board_replaces_the_first()
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();
            SubmissionRecord submission = NewBoardMaintainerFlowTests.Submission(null, "bob@example.com");

            await NewBoardMaintainerFlowTests.RunAsync(world, submission);
            await NewBoardMaintainerFlowTests.RunAsync(world, submission);

            Assert.Single(await world.Accounts.ListOpenInvitationsAsync(NewBoardMaintainerFlowTests.Now));
            Assert.Equal(2, world.Mailer.Sent.Count);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("not an address")]
        public async Task A_submission_with_no_address_makes_nobody_maintainer(string? contactEmail)
        {
            World world = await NewBoardMaintainerFlowTests.SetUpAsync();

            string note = await NewBoardMaintainerFlowTests.RunAsync(world, NewBoardMaintainerFlowTests.Submission(null, contactEmail));

            Assert.Equal(NewBoardMaintainerWording.NoAddress, note);
            Assert.Empty(world.Mailer.Sent);
            Assert.Empty(world.Accounts.Invitations);
            Assert.Empty(world.Accounts.Maintainers);
        }

        // -----------------------------------------------------------------------------------
        // The sentences
        // -----------------------------------------------------------------------------------

        // The approver reads these about a contributor: the person is named, never "their", "them"
        // or "they" (owner rule).
        [Fact]
        public void No_sentence_calls_the_creator_their_them_or_they()
        {
            string[] sentences =
            [
                NewBoardMaintainerWording.Added("Anna", "anna@example.com"),
                NewBoardMaintainerWording.AlreadyMaintains("", "anna@example.com"),
                NewBoardMaintainerWording.Invited("bob@example.com"),
                NewBoardMaintainerWording.Locked("bob@example.com"),
                NewBoardMaintainerWording.Unverified("bob@example.com"),
                NewBoardMaintainerWording.NoAddress,
                NewBoardMaintainerWording.Failed,
                NewBoardMaintainerWording.TypedAddressHasAccount("bob@example.com"),
            ];

            foreach (string sentence in sentences)
            {
                foreach (string word in sentence.Split([' ', ',', '.', '-', '"', '(', ')'], StringSplitOptions.RemoveEmptyEntries))
                    Assert.DoesNotContain(word.ToLowerInvariant(), new[] { "their", "them", "they" });
            }
        }

        // An account without a name is named by its address alone.
        [Fact]
        public void An_account_without_a_name_is_named_by_its_address()
        {
            Assert.Equal(
                "anna@example.com, who created this new board, is now its maintainer.",
                NewBoardMaintainerWording.Added(" ", "anna@example.com"));
        }
    }
}
