using CRT.Server.Handlers.Email;
using CRT.Server.Handlers.Submissions;
using CRT.Server.Tests.Fakes;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Telling the contributor what a reviewer decided (maintainer request, 2026-09-23).
    //
    // *** WHY THIS IS WORTH TESTING RATHER THAN EYEBALLING. *** The contributor has no account, so
    // the mail and the "My submissions" window are the only two channels that exist - and the mail
    // is the only one that reaches somebody who is not sitting in front of CRT. Sending the wrong
    // outcome, or sending nothing, is invisible from the server's own side: the decision is
    // recorded either way and the reviewer sees a success.
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
        private const string SystemName = "Commodore/C64/250407";

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
                [SubmissionNotifierTests.Contributor, SubmissionNotifierTests.SystemName, state, comment]);
        }

        // ---------------------------------------------------------------- which mail, if any

        [Theory]
        [InlineData("published")]
        [InlineData("merged")]
        public async Task A_published_submission_is_told_it_is_in(string state)
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.SystemName,
                state,
                reviewerComment: null);

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Equal(SubmissionNotifierTests.Contributor, message.ToAddress);
            Assert.Contains("published", message.Subject, StringComparison.OrdinalIgnoreCase);

            // The board has to be named: a contributor who sent something three weeks ago cannot
            // act on "your contribution was accepted".
            Assert.Contains(SubmissionNotifierTests.SystemName, message.Body, StringComparison.Ordinal);
        }

        [Fact]
        public async Task A_rejection_carries_the_reason_the_reviewer_gave()
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.SystemName,
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
                SubmissionNotifierTests.SystemName,
                "changes_requested",
                "Please add the PAL region to U8.");

            EmailMessage message = Assert.Single(mailer.Sent);

            Assert.Contains("Please add the PAL region to U8.", message.Body, StringComparison.Ordinal);

            // This is the one outcome where the contributor has to act, so the mail has to point
            // them at where they act.
            Assert.Contains("Drafts tab", message.Body, StringComparison.OrdinalIgnoreCase);
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
        public async Task A_state_that_is_not_a_reviewer_decision_sends_nothing(string state)
        {
            var mailer = new FakeEmailSender();

            await SubmissionNotifierTests.Notifier(mailer).NotifyDecisionAsync(
                SubmissionNotifierTests.Contributor,
                SubmissionNotifierTests.SystemName,
                state,
                reviewerComment: "ignored");

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
                SubmissionNotifierTests.SystemName,
                "published",
                reviewerComment: null);

            Assert.Empty(mailer.Sent);
        }

        // ---------------------------------------------------------------- failure is contained

        // ###########################################################################################
        // *** THE MOST IMPORTANT TEST HERE. ***
        //
        // By the time this runs the decision is recorded, and for a publish the board has already
        // been overwritten - the one irreversible operation in the system. An exception escaping
        // would be shown to the reviewer as a failure, and they would quite reasonably repeat a
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
                SubmissionNotifierTests.SystemName,
                "published",
                reviewerComment: null);
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
            public Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
                => throw new InvalidOperationException("SMTP is unreachable.");
        }
    }
}
