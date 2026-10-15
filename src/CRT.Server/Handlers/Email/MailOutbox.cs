namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // MAIL THE MAIL SYSTEM REFUSED IS KEPT AND SENT AGAIN (code review, 2026-10-13; migration 0022).
    //
    // Every mail was fire-and-forget: SmtpEmailSender logged a refusal and the message was gone, so
    // with postfix down for an hour every decision, "waiting for your review" and approval-needed
    // mail of that hour reached nobody, and nothing on any screen said so. Now:
    //
    //   - OutboxEmailSender, the IEmailSender every flow is handed, writes a refused mail into the
    //     outbox (and still answers false, so a caller that cares - feedback - knows);
    //   - MailOutboxSender tries the outbox again every few minutes, waiting longer after each
    //     failure (RetryDelay), and gives a mail up after GiveUpAfter;
    //   - Account > "Server version" shows how many are waiting (ServerStatusAnswer.UnsentMails).
    //
    // *** A MAIL CARRYING A WORKING CREDENTIAL IS KEPT WITHOUT ITS WORDS. *** Only token hashes are
    // stored, ever (CLAUDE.md) - so a verification, password reset, address-change or invitation mail
    // is recorded with no body: counted, never retried, given up with the rest. The person asks for
    // another (or the administrator invites again); the count tells the administrator to look.
    //
    // The words are MailBody's as they were rendered - HTML and plain text both - so a mail sent
    // from the outbox is the mail that was written.
    // ###########################################################################################
    public static class MailOutboxFlow
    {
        // How long a refused mail is tried before it is given up - long enough to outlast a weekend
        // outage, short enough that nobody receives a "your submission is accepted" a month late.
        public static readonly TimeSpan GiveUpAfter = TimeSpan.FromDays(3);

        // How often the outbox is looked at.
        public static readonly TimeSpan Interval = TimeSpan.FromMinutes(5);

        // At most this many mails per pass - a long outage's backlog goes out over a few passes
        // rather than in one burst at the mail system that has only just come back.
        public const int BatchSize = 50;

        // ###########################################################################################
        // How long to wait before trying a mail again after `attempts` tries: 5 minutes, then
        // doubling, never more than 4 hours.
        // ###########################################################################################
        public static TimeSpan RetryDelay(int attempts)
        {
            int doublings = Math.Clamp(attempts - 1, 0, 10);
            double minutes = Math.Min(5 * Math.Pow(2, doublings), 240);
            return TimeSpan.FromMinutes(minutes);
        }

        // ###########################################################################################
        // The outbox row for a mail the mail system refused - with its words, or without them for
        // one carrying a credential (see the header).
        // ###########################################################################################
        public static OutboxMail ForRefused(EmailMessage message, DateTimeOffset nowUtc, string? error)
        {
            ArgumentNullException.ThrowIfNull(message);

            bool keepWords = !message.CarriesCredential;

            return new OutboxMail(
                0,
                message.ToAddress,
                message.Subject,
                keepWords ? message.Body : null,
                keepWords ? message.HtmlBody : null,
                message.ReplyToAddress,
                nowUtc,
                1,
                nowUtc + MailOutboxFlow.RetryDelay(1),
                MailOutboxFlow.Shorten(error));
        }

        // ###########################################################################################
        // One pass: every due mail with words is handed to the mail system again - sent ones leave
        // the outbox, refused ones wait longer - and everything older than GiveUpAfter is given up.
        // Returns how many went. Never throws for a single mail; the caller guards the pass.
        // ###########################################################################################
        public static async Task<int> SendDueAsync(
            IMailOutboxStore outbox,
            IEmailSender mailSystem,
            DateTimeOffset nowUtc,
            ILogger logger,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(outbox);
            ArgumentNullException.ThrowIfNull(mailSystem);
            ArgumentNullException.ThrowIfNull(logger);

            int sent = 0;

            foreach (OutboxMail mail in await outbox.GetDueAsync(nowUtc, MailOutboxFlow.BatchSize, cancellationToken).ConfigureAwait(false))
            {
                if (mail.Body is null)
                    continue;

                bool went = await mailSystem.SendAsync(
                    new EmailMessage(mail.ToAddress, mail.Subject, mail.Body, mail.HtmlBody, mail.ReplyToAddress),
                    cancellationToken).ConfigureAwait(false);

                if (went)
                {
                    await outbox.RemoveAsync(mail.Id, cancellationToken).ConfigureAwait(false);
                    sent++;
                    continue;
                }

                int attempts = mail.Attempts + 1;

                await outbox.RecordFailureAsync(
                    mail.Id,
                    attempts,
                    nowUtc + MailOutboxFlow.RetryDelay(attempts),
                    "The mail system refused it again.",
                    cancellationToken).ConfigureAwait(false);
            }

            int givenUp = await outbox.RemoveCreatedBeforeAsync(nowUtc - MailOutboxFlow.GiveUpAfter, cancellationToken).ConfigureAwait(false);

            if (sent > 0)
                logger.LogInformation("Sent {Count} mail(s) from the outbox.", sent);

            if (givenUp > 0)
            {
                logger.LogWarning(
                    "Gave up {Count} mail(s) the mail system refused for {Days} days.",
                    givenUp,
                    MailOutboxFlow.GiveUpAfter.TotalDays);
            }

            return sent;
        }

        private static string? Shorten(string? text) =>
            string.IsNullOrWhiteSpace(text) ? null : text.Length <= 500 ? text : text[..500];
    }

    // ###########################################################################################
    // One mail in the outbox. Body is null for a mail kept without its words (a credential) - never
    // sent from here.
    // ###########################################################################################
    public sealed record OutboxMail(
        long Id,
        string ToAddress,
        string Subject,
        string? Body,
        string? HtmlBody,
        string? ReplyToAddress,
        DateTimeOffset CreatedUtc,
        int Attempts,
        DateTimeOffset NextAttemptUtc,
        string? LastError);

    // The outbox's seam - the MySQL store and a test fake.
    public interface IMailOutboxStore
    {
        Task AddAsync(OutboxMail mail, CancellationToken cancellationToken = default);

        // Mails due a try at `nowUtc`, oldest first, at most `limit`.
        Task<IReadOnlyList<OutboxMail>> GetDueAsync(DateTimeOffset nowUtc, int limit, CancellationToken cancellationToken = default);

        Task RemoveAsync(long id, CancellationToken cancellationToken = default);

        Task RecordFailureAsync(long id, int attempts, DateTimeOffset nextAttemptUtc, string? error, CancellationToken cancellationToken = default);

        // Gives up every mail first refused before `createdBefore`; how many.
        Task<int> RemoveCreatedBeforeAsync(DateTimeOffset createdBefore, CancellationToken cancellationToken = default);

        // How many mails wait - Account > "Server version".
        Task<int> CountAsync(CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // The IEmailSender every flow is handed: the mail system first, the outbox when it refuses
    // (MailOutboxFlow). Answers what the mail system answered - false still means "it did not go
    // now". A mail that cannot even be kept is logged; nothing here throws back at the operation it
    // accompanies (IEmailSender's header).
    // ###########################################################################################
    public sealed class OutboxEmailSender : IEmailSender
    {
        private readonly IEmailSender thisMailSystem;
        private readonly IMailOutboxStore thisOutbox;
        private readonly ILogger thisLogger;
        private readonly TimeProvider thisTime;

        public OutboxEmailSender(IEmailSender mailSystem, IMailOutboxStore outbox, ILogger<OutboxEmailSender> logger)
            : this(mailSystem, outbox, logger, TimeProvider.System)
        {
        }

        internal OutboxEmailSender(IEmailSender mailSystem, IMailOutboxStore outbox, ILogger logger, TimeProvider time)
        {
            ArgumentNullException.ThrowIfNull(mailSystem);
            ArgumentNullException.ThrowIfNull(outbox);
            ArgumentNullException.ThrowIfNull(logger);
            ArgumentNullException.ThrowIfNull(time);

            this.thisMailSystem = mailSystem;
            this.thisOutbox = outbox;
            this.thisLogger = logger;
            this.thisTime = time;
        }

        public async Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(message);

            if (await this.thisMailSystem.SendAsync(message, cancellationToken).ConfigureAwait(false))
                return true;

            if (!message.KeepUntilSent)
                return false;

            try
            {
                // Not the request's token: the mail is already refused, and keeping it is the point.
                await this.thisOutbox
                    .AddAsync(MailOutboxFlow.ForRefused(message, this.thisTime.GetUtcNow(), "The mail system refused it."), CancellationToken.None)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogError(ex, "A refused mail to {Recipient} ({Subject}) could not be kept in the outbox either.", message.ToAddress, message.Subject);
            }

            return false;
        }
    }

    // ###########################################################################################
    // Runs MailOutboxFlow.SendDueAsync every MailOutboxFlow.Interval for as long as the service runs
    // - through the MAIL SYSTEM itself, never the outbox sender, so a refusal is not queued twice. A
    // failed pass is logged and the next one tries again: from .NET 8 an exception escaping a
    // BackgroundService stops the host (AbandonedUploadSweeper's header).
    // ###########################################################################################
    public sealed class MailOutboxSender : BackgroundService
    {
        private readonly IMailOutboxStore thisOutbox;
        private readonly SmtpEmailSender thisMailSystem;
        private readonly ILogger<MailOutboxSender> thisLogger;

        public MailOutboxSender(IMailOutboxStore outbox, SmtpEmailSender mailSystem, ILogger<MailOutboxSender> logger)
        {
            this.thisOutbox = outbox;
            this.thisMailSystem = mailSystem;
            this.thisLogger = logger;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            using var timer = new PeriodicTimer(MailOutboxFlow.Interval);

            try
            {
                while (await timer.WaitForNextTickAsync(stoppingToken))
                {
                    try
                    {
                        await MailOutboxFlow.SendDueAsync(this.thisOutbox, this.thisMailSystem, DateTimeOffset.UtcNow, this.thisLogger, stoppingToken);
                    }
                    catch (Exception ex) when (ex is not OperationCanceledException)
                    {
                        this.thisLogger.LogWarning(ex, "The mail outbox could not be sent; it is tried again in {Interval}.", MailOutboxFlow.Interval);
                    }
                }
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                // The service is stopping.
            }
        }
    }
}
