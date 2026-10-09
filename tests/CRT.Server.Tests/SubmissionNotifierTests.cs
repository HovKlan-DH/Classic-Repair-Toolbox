using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Telling the contributor what a maintainer decided (owner request, 2026-09-23).
    //
    // *** WHY THIS IS WORTH TESTING RATHER THAN EYEBALLING. *** The contributor has no account, so
    // the mail and the "My submissions" window are the only two channels that exist - and the mail
    // is the only one that reaches somebody who is not sitting in front of CRT. Sending the wrong
    // outcome, or sending nothing, is invisible from the server's own side: the decision is
    // recorded either way and the maintainer sees a success.
    //
    // Three properties are pinned throughout:
    //   1. WHICH mail a state produces, because sending "not accepted" to somebody whose work was
    //      published is worse than sending nothing at all.
    //   2. That a state with no mail sends NOTHING, so a future outcome added server-side stays
    //      silent rather than guessing.
    //   3. That a failure to send never escapes, because by then the decision is already durable -
    //      and for a publish, the data tree is already overwritten.
    // ###########################################################################################
    public sealed class SubmissionNotifierTests
    {
        private const string Contributor = "someone@example.com";
        private const string BoardDisplayName = "Commodore/C64/250407";

        // Addresses with no names - what most of these tests are about.
        private static IReadOnlyList<MailRecipient> Recipients(params string?[] addresses) =>
            addresses.Select(address => new MailRecipient(address!)).ToList();

        private static SubmissionNotifier Notifier(IEmailSender mailer) =>
            new(mailer, NullLogger<SubmissionNotifier>.Instance);

        // ###########################################################################################
        // The state-to-mail mapping is internal, and CRT.Server has no InternalsVisibleTo for this
        // project - so it is reached by reflection, exactly as ReviewPublishedFilesTests reaches
        // ReviewEndpoints.PublishedFilePaths. Widening it to public for a test would put it on the
        // server's API surface for no caller.
        // ###########################################################################################
        private static EmailMessage? BuildMessage(string state, string? comment)
        {
            var method = typeof(SubmissionNotifier).GetMethod(
                "BuildMessage",
                System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.NonPublic);

            Assert.NotNull(method);

            return (EmailMessage?)method!.Invoke(
                null,
                // amendedByMaintainer (2026-09-25) - false, a submission nobody changed; then the
                // contributor's name (2026-10-03) - none, a contributor without an account.
                [SubmissionNotifierTests.Contributor, SubmissionNotifierTests.BoardDisplayName, state, comment, false, null]);
        }

        // ---------------------------------------------------------------- which mail, if any

        // ###########################################################################################
        // *** TWO STAGES, TWO MAILS (2026-09-25). *** "merged" is the approval, which writes the
        // BETA data - "published to the BETA source". "published" is its board reaching production
        // - "published to the stable source". The first must say BETA: their own data (on the stable
        // source) will not have it for days.
        // ###########################################################################################
        [Fact]
        public async Task An_approved_submission_is_told_it_was_published_to_the_BETA_SOURCE()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "merged",
                maintainerComment: null);

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Equal(SubmissionNotifierTests.Contributor, message.ToAddress);
            Assert.Contains("BETA source", message.Subject, StringComparison.Ordinal);
            Assert.Contains("BETA source", message.Body, StringComparison.Ordinal);

            // The board has to be named: a contributor who sent something three weeks ago cannot
            // act on "your contribution was accepted".
            Assert.Contains(SubmissionNotifierTests.BoardDisplayName, message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_submission_whose_board_reached_PRODUCTION_is_told_it_is_live()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "published",
                maintainerComment: null);

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Contains("published to the stable source", message.Subject, StringComparison.Ordinal);
            Assert.DoesNotContain("BETA", message.Subject, StringComparison.Ordinal);
            Assert.Contains(SubmissionNotifierTests.BoardDisplayName, message.Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** A ROLLED-BACK CONTRIBUTOR IS ACTUALLY TOLD (owner decision, 2026-09-27). *** The
        // first version of the rollback sent NotifyDecisionAsync with state "pending", which
        // BuildMessage answers with NULL - so the rollback completed, the submission went back to
        // the queue, and NOBODY was mailed. The suite was green because nothing asked. This asks.
        //
        // The mail must say the data was IN BETA and now is not: this contributor was already told
        // "published to BETA", and a mail reading like an ordinary change request would leave them
        // believing their work was still live.
        // ###########################################################################################
        [Fact]
        public async Task A_contributor_whose_board_was_rolled_back_is_told_it_left_BETA_and_why()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyReturnedToQueueAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "The U8 pinout is wrong.");

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Equal(SubmissionNotifierTests.Contributor, message.ToAddress);
            Assert.Contains("taken back out of BETA", message.Subject, StringComparison.Ordinal);

            // The board is named - a contributor who sent something weeks ago cannot act on "your
            // contribution was taken back".
            Assert.Contains(SubmissionNotifierTests.BoardDisplayName, message.Body, StringComparison.Ordinal);
            Assert.Contains("had been accepted into the BETA source", message.Body, StringComparison.Ordinal);
            Assert.Contains("no longer holds it", message.Body, StringComparison.Ordinal);
            Assert.Contains("The U8 pinout is wrong.", message.Body, StringComparison.Ordinal);

            // What IS true: the contribution is in the queue, and it can be corrected.
            Assert.Contains("is not lost", message.Body, StringComparison.Ordinal);
            Assert.Contains("make the change in the \"Drafts\" tab", message.Body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // Beta > Prod's "Reject" (owner request, 2026-09-28): the contributor gets the queue's own
        // rejection mail, with the reason - not "back in the queue", which it is not.
        // ###########################################################################################
        [Fact]
        public async Task A_board_rejected_out_of_BETA_mails_a_rejection_and_one_pushed_back_the_queue_mail()
        {
            var rejectedMailer = new FakeEmailSender();
            var returnedMailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(rejectedMailer).NotifyTakenOutOfBetaAsync(
                SubmissionNotifierTests.Contributor, SubmissionNotifierTests.BoardDisplayName, "Not for this board.", rejected: true);

            await SubmissionNotifierTests.Notifier(returnedMailer).NotifyTakenOutOfBetaAsync(
                SubmissionNotifierTests.Contributor, SubmissionNotifierTests.BoardDisplayName, "Not ready.", rejected: false);

            EmailMessage rejection = Assert.Single(rejectedMailer.Sent);
            Assert.Contains("will not be going", rejection.Body, StringComparison.Ordinal);
            Assert.Contains("Not for this board.", rejection.Body, StringComparison.Ordinal);
            Assert.DoesNotContain("taken back out of BETA", rejection.Subject, StringComparison.Ordinal);

            Assert.Contains("taken back out of BETA", Assert.Single(returnedMailer.Sent).Subject, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** IT PROMISES NOTHING THE BOARD DOES NOT GUARANTEE (code review, 2026-09-27). *** The
        // first version told the contributor their draft was "still on your own computer, exactly
        // as you left it" - but CRT deletes a draft once the published board matches it, which can
        // happen while the work sits in BETA - and that a new submission "takes this one's place",
        // which SubmissionReplacementRules does not do for one a maintainer amended.
        // ###########################################################################################
        [Fact]
        public async Task The_returned_to_queue_mail_does_not_promise_a_draft_or_a_replacement()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyReturnedToQueueAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "Needs a revision date.");

            string body = Assert.Single(mailer.Sent).Body;

            Assert.DoesNotContain("exactly as you left it", body, StringComparison.Ordinal);
            Assert.DoesNotContain("place in the queue", body, StringComparison.Ordinal);

            // It says what to do when the draft IS gone, rather than assuming it is there.
            Assert.Contains("If you have discarded your draft", body, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // *** "pending" IS NOT A DECISION MAIL, and this pins the trap. *** It is also the state a
        // brand-new submission arrives in, so BuildMessage deliberately has no wording for it. A
        // caller that reaches for NotifyDecisionAsync(Pending) to say "returned to the queue" gets
        // silence - use NotifyReturnedToQueueAsync. Fails if someone adds a "pending" arm, which
        // would make every new arrival mail its own contributor a decision.
        // ###########################################################################################
        [Fact]
        public void A_pending_state_is_not_a_decision_and_builds_no_mail()
        {
            Assert.Null(SubmissionNotifier.BuildMessage(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "pending",
                maintainerComment: "Anything."));
        }

        [Fact]
        public async Task The_administrators_are_told_when_a_maintainer_publishes_to_production()
        {
            // The stand-in for the administrator feed: with no second factor on a maintainer's
            // account, an unexpected production publish must be noticed.
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyProductionPublishAsync(
                SubmissionNotifierTests.Recipients("admin@example.com", "Admin@example.com", null),
                SubmissionNotifierTests.BoardDisplayName,
                "Anna (anna@example.com)",
                "2026-September-25",
                3);

            EmailMessage message = Assert.Single(mailer.Sent);
            Assert.Equal("admin@example.com", message.ToAddress);
            Assert.Contains("Anna (anna@example.com)", message.Body, StringComparison.Ordinal);
            Assert.Contains("stable source", message.Subject, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("3 file(s)", message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_rejection_carries_the_reason_the_maintainer_gave()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "rejected",
                "The pin numbering does not match the datasheet.");

            EmailMessage message = Assert.Single(mailer.Sent);

            // *** THE REASON IS THE ENTIRE POINT. *** A rejection with no explanation is
            // indistinguishable from being ignored, which is why the server refuses to record one
            // without a comment in the first place.
            Assert.Contains(
                "The pin numbering does not match the datasheet.",
                message.Body,
                StringComparison.Ordinal);

            // And it must say the work is not gone, or this reads as "your afternoon was wasted".
            Assert.Contains("still on your own computer", message.Body, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task A_change_request_says_what_to_do_next()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "changes_requested",
                "Please add the PAL region to U8.");

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Contains("Please add the PAL region to U8.", message.Body, StringComparison.Ordinal);

            // This is the one outcome where the contributor has to act, so the mail has to point
            // them at where they act.
            Assert.Contains("\"Drafts\" tab", message.Body, StringComparison.Ordinal);
        }

        // ---------------------------------------------------------------- telling the maintainers

        [Fact]
        public async Task Every_maintainer_is_told_once_and_a_blank_or_repeated_address_is_dropped()
        {
            // A maintainer who is also listed twice - or an administrator who is both - hears once.
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyMaintainersAsync(
                SubmissionNotifierTests.Recipients("anna@example.com", " ", null, "Anna@Example.com", "bob@example.com"),
                SubmissionNotifierTests.BoardDisplayName,
                42,
                "Corrected R12.");

            Assert.Equal(["anna@example.com", "bob@example.com"], mailer.Sent.Select(message => message.ToAddress));

            EmailMessage first = mailer.Sent[0];
            Assert.Contains("#42", first.Body, StringComparison.Ordinal);
            Assert.Contains("Corrected R12.", first.Body, StringComparison.Ordinal);
            Assert.Contains(SubmissionNotifierTests.BoardDisplayName, first.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_mailer_that_throws_does_not_stop_the_other_maintainers_being_told()
        {
            // The submission is already queued; one dead address must not silence the rest, and
            // nothing may escape to the contributor's finalise request.
            var mailer = new ThrowingOnceEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyMaintainersAsync(
                SubmissionNotifierTests.Recipients("first@example.com", "second@example.com"),
                SubmissionNotifierTests.BoardDisplayName,
                7,
                null);

            Assert.Equal(["first@example.com", "second@example.com"], mailer.Attempted);
        }

        private sealed class ThrowingOnceEmailSender : IEmailSender
        {
            public List<string> Attempted { get; } = [];

            public Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
            {
                this.Attempted.Add(message.ToAddress);

                if (this.Attempted.Count == 1)
                    throw new InvalidOperationException("SMTP is down.");

                return Task.FromResult(true);
            }
        }

        // ###########################################################################################
        // *** STATES THAT MUST SEND NOTHING. ***
        //
        // "pending" and "uploading" are passed through on the way in, not decided by anybody -
        // mailing somebody that their upload completed is noise that teaches them to ignore the
        // mails that matter. "approved" means published-is-next rather than published, and would
        // be followed minutes later by the real one.
        //
        // An unknown state sends nothing rather than guessing: a new outcome added server-side
        // should be silent until somebody writes its mail.
        // ###########################################################################################
        [Theory]
        [InlineData("pending")]
        [InlineData("uploading")]
        [InlineData("approved")]
        [InlineData("accepted")]
        [InlineData("withdrawn")]
        [InlineData("abandoned")]
        [InlineData("some_future_state")]
        [InlineData("")]
        public async Task A_state_that_is_not_a_maintainer_decision_sends_nothing(string state)
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                state,
                maintainerComment: "ignored");

            Assert.Empty(mailer.Sent);
        }

        // ---------------------------------------------------------------- no address

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("   ")]
        public async Task No_contact_address_means_no_mail_and_no_error(string? address)
        {
            // Normal, not exceptional: a submission from a signed-in maintainer carries no contact
            // address at all. The outcome is still in "My submissions".
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                address,
                SubmissionNotifierTests.BoardDisplayName,
                "published",
                maintainerComment: null);

            Assert.Empty(mailer.Sent);
        }

        // ---------------------------------------------------------------- failure is contained

        // ###########################################################################################
        // *** THE MOST IMPORTANT TEST HERE. ***
        //
        // By the time this runs the decision is recorded, and for a publish the board has already
        // been overwritten - the one irreversible operation in the board. An exception escaping
        // would be shown to the maintainer as a failure, and they would quite reasonably repeat a
        // decision that has in fact already been made.
        //
        // Proved with a mailer that throws, which is what an unreachable SMTP host looks like.
        // ###########################################################################################
        [Fact]
        public async Task A_mailer_that_THROWS_does_not_fail_the_decision()
        {
            var notifier = SubmissionNotifierTests.Notifier(new ThrowingEmailSender());

            // No assertion beyond "this returns" - the absence of an exception IS the behaviour.
            await notifier.NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.BoardDisplayName,
                "published",
                maintainerComment: null);
        }

        // ---------------------------------------------------------------- whom a submission's mail goes to

        private static SubmissionRecord Submission(long? account, string? contactEmail) =>
            new(1, SubmissionNotifierTests.BoardDisplayName, account, contactEmail, "hash", "", "merged", null, 1,
                new DateTimeOffset(2026, 10, 1, 12, 0, 0, TimeSpan.Zero), null, null);

        // ###########################################################################################
        // *** A SUBMISSION SENT SIGNED IN IS WRITTEN TO ITS ACCOUNT (owner request, 2026-10-01). ***
        // It carries no contact address of its own, and the decision mails read only that - so a
        // signed-in maintainer's own submission was approved, rejected or sent back in silence. Fails
        // against the version that handed the endpoints' ContactEmail straight through.
        // ###########################################################################################
        [Fact]
        public async Task A_signed_in_submissions_decision_mail_goes_to_its_accounts_address()
        {
            var mailer = new FakeEmailSender();
            var accounts = new FakeAccountStore();

            accounts.Accounts[7] = new CRT.Server.Handlers.Accounts.AccountRecord(
                7, "dh@example.com", "dh@example.com", "hash", "Dennis", IsVerified: true, IsAdministrator: false, IsLocked: false,
                new DateTimeOffset(2026, 9, 2, 12, 0, 0, TimeSpan.Zero), null);

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Submission(account: 7, contactEmail: null),
                accounts,
                "rejected",
                maintainerComment: "Not this board.");

            Assert.Equal("dh@example.com", Assert.Single(mailer.Sent).ToAddress);
        }

        // Without an account it is the address typed when sending, as it always was.
        [Fact]
        public async Task A_submission_sent_without_an_account_is_written_to_its_contact_address()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Submission(account: null, contactEmail: SubmissionNotifierTests.Contributor),
                new FakeAccountStore(),
                "merged",
                maintainerComment: null);

            Assert.Equal(SubmissionNotifierTests.Contributor, Assert.Single(mailer.Sent).ToAddress);
        }

        // ###########################################################################################
        // The greeting (owner request, 2026-10-03: "if {name} is known, use that, otherwise just
        // "Hi there.""): the account's name for a submission sent signed in - the only case a name is
        // known - and "Hi there," for one sent with a typed address.
        // ###########################################################################################
        [Fact]
        public async Task A_signed_in_contributor_is_greeted_by_the_accounts_name_and_anybody_else_as_there()
        {
            var mailer = new FakeEmailSender();
            var accounts = new FakeAccountStore();

            accounts.Accounts[7] = new CRT.Server.Handlers.Accounts.AccountRecord(
                7, "dh@example.com", "dh@example.com", "hash", "Dennis", IsVerified: true, IsAdministrator: false, IsLocked: false,
                new DateTimeOffset(2026, 9, 2, 12, 0, 0, TimeSpan.Zero), null);

            SubmissionNotifier notifier = SubmissionNotifierTests.Notifier(mailer);

            await notifier.NotifyDecisionAsync(SubmissionNotifierTests.Submission(account: 7, contactEmail: null), accounts, "merged", null);
            await notifier.NotifyDecisionAsync(SubmissionNotifierTests.Submission(account: null, contactEmail: SubmissionNotifierTests.Contributor), accounts, "merged", null);

            Assert.StartsWith("Hi Dennis,", mailer.Sent[0].Body, StringComparison.Ordinal);
            Assert.StartsWith("Hi there,", mailer.Sent[1].Body, StringComparison.Ordinal);
        }

        // Each maintainer by the name on their account, and the mail says whether it is a new
        // board (owner request, 2026-10-03).
        [Fact]
        public async Task Each_maintainer_is_greeted_by_name_and_told_whether_it_is_a_new_board()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyMaintainersAsync(
                [new MailRecipient("anna@example.com", "Anna"), new MailRecipient("bob@example.com", "Bob")],
                SubmissionNotifierTests.BoardDisplayName,
                18,
                "A whole new board.",
                isNewBoard: true);

            Assert.StartsWith("Hi Anna,", mailer.Sent[0].Body, StringComparison.Ordinal);
            Assert.StartsWith("Hi Bob,", mailer.Sent[1].Body, StringComparison.Ordinal);
            Assert.All(mailer.Sent, message => Assert.Contains("a completely new board", message.Body, StringComparison.Ordinal));
        }

        // ---------------------------------------------------------------- the mapping itself

        [Fact]
        public void The_three_decisions_produce_three_DIFFERENT_mails()
        {
            // Pinned together so that a copy-paste between templates cannot leave two outcomes
            // telling the contributor the same thing - the failure this mapping most invites.
            EmailMessage? published = SubmissionNotifierTests.BuildMessage("merged", null);
            EmailMessage? rejected = SubmissionNotifierTests.BuildMessage("rejected", "no");
            EmailMessage? changes = SubmissionNotifierTests.BuildMessage("changes_requested", "fix");

            Assert.NotNull(published);
            Assert.NotNull(rejected);
            Assert.NotNull(changes);

            Assert.Equal(3, new HashSet<string>(
                [published!.Subject, rejected!.Subject, changes!.Subject],
                StringComparer.Ordinal).Count);

            Assert.Equal(3, new HashSet<string>(
                [published.Body, rejected.Body, changes.Body],
                StringComparer.Ordinal).Count);
        }

        [Fact]
        public void The_state_is_matched_case_insensitively_and_trimmed()
        {
            // The value arrives from the database and from SubmissionState constants; a stray
            // space or a capital must not silently send nothing.
            Assert.NotNull(SubmissionNotifierTests.BuildMessage("  Merged ", null));
        }

        private sealed class ThrowingEmailSender : IEmailSender
        {
            public Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
                => throw new InvalidOperationException("SMTP is unreachable.");
        }
    }
}
