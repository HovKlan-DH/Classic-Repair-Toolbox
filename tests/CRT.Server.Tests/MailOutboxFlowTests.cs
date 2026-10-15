using CRT.Server.Handlers.Email;
using CRT.Server.Tests.Fakes;
using Microsoft.Extensions.Logging.Abstractions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the mail outbox (code review, 2026-10-13): a mail the mail system refuses is KEPT and
    // sent again, rather than logged and lost - except one carrying a working credential, which is
    // only counted (only token hashes are stored, ever), and one whose sender is told to try again
    // (feedback). No network: the mail system is FakeEmailSender, the table FakeMailOutboxStore.
    // ###########################################################################################
    public sealed class MailOutboxFlowTests
    {
        private static readonly DateTimeOffset Now = new(2026, 10, 13, 12, 0, 0, TimeSpan.Zero);

        private static readonly EmailMessage Decision = new(
            "contributor@example.com", "Your CRT contribution is accepted", "The words.", "<p>The words.</p>");

        private sealed class FixedTime(DateTimeOffset now) : TimeProvider
        {
            public override DateTimeOffset GetUtcNow() => now;
        }

        private static OutboxEmailSender Sender(FakeEmailSender mailSystem, FakeMailOutboxStore outbox) =>
            new(mailSystem, outbox, NullLogger.Instance, new FixedTime(MailOutboxFlowTests.Now));

        [Fact]
        public async Task A_mail_the_mail_system_takes_is_not_kept()
        {
            var mailSystem = new FakeEmailSender();
            var outbox = new FakeMailOutboxStore();

            Assert.True(await MailOutboxFlowTests.Sender(mailSystem, outbox).SendAsync(MailOutboxFlowTests.Decision));
            Assert.Empty(outbox.Mails);
        }

        // ###########################################################################################
        // *** A REFUSED MAIL IS KEPT, WORDS AND ALL. *** It used to be logged and gone - a decision
        // mail sent while postfix was down reached nobody. The caller is still told it did not go.
        // ###########################################################################################
        [Fact]
        public async Task A_refused_mail_is_kept_with_both_bodies_and_the_caller_is_told_it_did_not_go()
        {
            var mailSystem = new FakeEmailSender { Delivers = false };
            var outbox = new FakeMailOutboxStore();

            Assert.False(await MailOutboxFlowTests.Sender(mailSystem, outbox).SendAsync(MailOutboxFlowTests.Decision));

            OutboxMail kept = Assert.Single(outbox.Mails);
            Assert.Equal("contributor@example.com", kept.ToAddress);
            Assert.Equal("The words.", kept.Body);
            Assert.Equal("<p>The words.</p>", kept.HtmlBody);
            Assert.Equal(1, kept.Attempts);
            Assert.True(kept.NextAttemptUtc > MailOutboxFlowTests.Now);
        }

        // ###########################################################################################
        // *** A CREDENTIAL IS NEVER STORED. *** A verification link, a reset or address-change code,
        // an invitation code: only their hashes may be in the database, so such a mail is kept
        // WITHOUT its words - counted, never sent again.
        // ###########################################################################################
        [Fact]
        public async Task A_refused_mail_carrying_a_credential_is_counted_but_kept_without_its_words()
        {
            var mailSystem = new FakeEmailSender { Delivers = false };
            var outbox = new FakeMailOutboxStore();

            EmailMessage invitation = EmailTemplates.MaintainerInvitation(
                "new@example.com", "Commodore / C64 / 250407", "Dennis", "ABCDEFGHIJKLMNOPQRSTUV", 72);

            Assert.True(invitation.CarriesCredential);

            await MailOutboxFlowTests.Sender(mailSystem, outbox).SendAsync(invitation);

            OutboxMail kept = Assert.Single(outbox.Mails);
            Assert.Null(kept.Body);
            Assert.Null(kept.HtmlBody);
            Assert.DoesNotContain("ABCDEFGHIJKLMNOPQRSTUV", kept.Subject, StringComparison.Ordinal);
        }

        // Feedback: CRT is told it failed and sends it again, so a later copy would arrive twice.
        [Fact]
        public async Task A_refused_mail_its_sender_will_send_again_is_not_kept()
        {
            var mailSystem = new FakeEmailSender { Delivers = false };
            var outbox = new FakeMailOutboxStore();

            await MailOutboxFlowTests.Sender(mailSystem, outbox).SendAsync(MailOutboxFlowTests.Decision with { KeepUntilSent = false });

            Assert.Empty(outbox.Mails);
        }

        // An outbox that cannot be written either never throws back at the operation the mail
        // accompanied (IEmailSender's header).
        [Fact]
        public async Task An_outbox_that_cannot_be_written_does_not_throw()
        {
            var mailSystem = new FakeEmailSender { Delivers = false };
            var outbox = new FakeMailOutboxStore { FailAdd = true };

            Assert.False(await MailOutboxFlowTests.Sender(mailSystem, outbox).SendAsync(MailOutboxFlowTests.Decision));
        }

        [Fact]
        public async Task A_kept_mail_is_sent_again_once_due_and_then_leaves_the_outbox()
        {
            var outbox = new FakeMailOutboxStore();
            await outbox.AddAsync(MailOutboxFlow.ForRefused(MailOutboxFlowTests.Decision, MailOutboxFlowTests.Now, "down"));

            var mailSystem = new FakeEmailSender();

            // Not yet due: nothing is tried.
            Assert.Equal(0, await MailOutboxFlow.SendDueAsync(outbox, mailSystem, MailOutboxFlowTests.Now, NullLogger.Instance));
            Assert.Empty(mailSystem.Sent);

            int sent = await MailOutboxFlow.SendDueAsync(outbox, mailSystem, MailOutboxFlowTests.Now.AddHours(1), NullLogger.Instance);

            Assert.Equal(1, sent);
            Assert.Equal("The words.", Assert.Single(mailSystem.Sent).Body);
            Assert.Equal("<p>The words.</p>", mailSystem.Sent[0].HtmlBody);
            Assert.Empty(outbox.Mails);
        }

        [Fact]
        public async Task A_mail_refused_again_waits_longer_before_the_next_try()
        {
            var outbox = new FakeMailOutboxStore();
            await outbox.AddAsync(MailOutboxFlow.ForRefused(MailOutboxFlowTests.Decision, MailOutboxFlowTests.Now, "down"));

            DateTimeOffset later = MailOutboxFlowTests.Now.AddHours(1);

            await MailOutboxFlow.SendDueAsync(outbox, new FakeEmailSender { Delivers = false }, later, NullLogger.Instance);

            OutboxMail kept = Assert.Single(outbox.Mails);
            Assert.Equal(2, kept.Attempts);
            Assert.Equal(later + MailOutboxFlow.RetryDelay(2), kept.NextAttemptUtc);
        }

        [Fact]
        public async Task A_mail_refused_for_longer_than_the_give_up_age_is_given_up()
        {
            var outbox = new FakeMailOutboxStore();
            await outbox.AddAsync(MailOutboxFlow.ForRefused(MailOutboxFlowTests.Decision, MailOutboxFlowTests.Now, "down"));

            await MailOutboxFlow.SendDueAsync(
                outbox, new FakeEmailSender { Delivers = false }, MailOutboxFlowTests.Now + MailOutboxFlow.GiveUpAfter + TimeSpan.FromMinutes(1), NullLogger.Instance);

            Assert.Empty(outbox.Mails);
        }

        // A mail kept without its words is never "sent" as an empty one.
        [Fact]
        public async Task A_mail_kept_without_its_words_is_never_sent()
        {
            var outbox = new FakeMailOutboxStore();
            await outbox.AddAsync(MailOutboxFlow.ForRefused(MailOutboxFlowTests.Decision with { CarriesCredential = true }, MailOutboxFlowTests.Now, "down"));

            var mailSystem = new FakeEmailSender();

            await MailOutboxFlow.SendDueAsync(outbox, mailSystem, MailOutboxFlowTests.Now.AddDays(1), NullLogger.Instance);

            Assert.Empty(mailSystem.Sent);
            Assert.Single(outbox.Mails);
        }

        [Fact]
        public void The_wait_between_tries_grows_and_stops_growing_at_four_hours()
        {
            Assert.Equal(TimeSpan.FromMinutes(5), MailOutboxFlow.RetryDelay(1));
            Assert.Equal(TimeSpan.FromMinutes(10), MailOutboxFlow.RetryDelay(2));
            Assert.Equal(TimeSpan.FromHours(4), MailOutboxFlow.RetryDelay(50));
        }

        // Every template carrying a working link or code says so - the rule the outbox keeps by.
        [Fact]
        public void Every_mail_with_a_working_link_or_code_is_marked_as_carrying_a_credential()
        {
            Assert.True(EmailTemplates.Verification("a@example.com", "A", "https://example.com/verify?token=x", 24).CarriesCredential);
            Assert.True(EmailTemplates.PasswordReset("a@example.com", "A", "CODECODECODECODECODE", 1).CarriesCredential);
            Assert.True(EmailTemplates.AlreadyRegistered("a@example.com", "A", "CODECODECODECODECODE", 1).CarriesCredential);
            Assert.True(EmailTemplates.EmailChangeCode("a@example.com", "A", "CODECODECODECODECODE", 1).CarriesCredential);
            Assert.False(EmailTemplates.PasswordChanged("a@example.com", "A").CarriesCredential);
        }
    }
}
