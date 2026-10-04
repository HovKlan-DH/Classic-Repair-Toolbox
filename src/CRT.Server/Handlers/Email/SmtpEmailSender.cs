using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using CRT.Server.Configuration;

namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // The real IEmailSender: hands the message to the local postfix on 127.0.0.1:25.
    //
    // THIS IS AN UNTESTED I/O BOUNDARY, on purpose and by the same rule as ScopeScpiClient and
    // MiniproProcessRunner in CRT.App. It contains no decisions - which mail, to whom, saying what
    // is all decided by EmailTemplates, which is pure and fully tested. What is left here is a
    // library call, and a test of it would be a test of System.Net.Mail.
    //
    // NO AUTHENTICATION AND NO TLS, deliberately. The connection is to 127.0.0.1 and never leaves
    // the box; postfix handles authenticated, encrypted delivery onward. Adding credentials here
    // would mean another secret in appsettings.Production.json to protect, guarding a hop that
    // cannot be observed without already having root on the machine.
    //
    // SmtpClient IS OBSOLETE in .NET and that is accepted. The replacement Microsoft points at
    // (MailKit) is a substantial dependency whose features - OAuth, TLS negotiation, IMAP - are
    // all for talking to a REMOTE server. For handing a plain message to a local postfix over
    // loopback, the obsolete class is exactly adequate. Revisit if this ever needs to talk to
    // anything but localhost.
    //
    // A FAILURE TO SEND MUST NOT FAIL THE OPERATION IT ACCOMPANIES - see IEmailSender's header.
    // Everything is caught and logged.
    // ###########################################################################################
    public sealed class SmtpEmailSender : IEmailSender
    {
        private readonly ServerOptions thisOptions;
        private readonly ILogger<SmtpEmailSender> thisLogger;

        public SmtpEmailSender(ServerOptions options, ILogger<SmtpEmailSender> logger)
        {
            ArgumentNullException.ThrowIfNull(options);
            ArgumentNullException.ThrowIfNull(logger);

            this.thisOptions = options;
            this.thisLogger = logger;
        }

        public async Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(message);

            try
            {
#pragma warning disable SYSLIB0014 // SmtpClient is obsolete - see the header for why it is used.
                using var client = new SmtpClient(this.thisOptions.SmtpHost, this.thisOptions.SmtpPort)
                {
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    EnableSsl = false,
                    Timeout = 15_000
                };
#pragma warning restore SYSLIB0014

                using var mail = new MailMessage
                {
                    From = new MailAddress(
                        this.thisOptions.MailFromAddress!,
                        this.thisOptions.MailFromDisplayName),
                    Subject = message.Subject,
                    SubjectEncoding = Encoding.UTF8,
                    Body = message.Body,
                    BodyEncoding = Encoding.UTF8,
                    IsBodyHtml = false
                };

                // ###########################################################################################
                // THE HTML IS SENT AS THE PREFERRED ALTERNATIVE (2026-10-03). multipart/alternative
                // with the plain text first and the HTML last - the order RFC 2046 gives for "the
                // last one is the best one" - so every client that can show HTML shows it, and the
                // plain part (which spam filters expect beside HTML) is only a fallback.
                // ###########################################################################################
                if (!string.IsNullOrEmpty(message.HtmlBody))
                {
                    mail.AlternateViews.Add(AlternateView.CreateAlternateViewFromString(
                        message.HtmlBody, Encoding.UTF8, MediaTypeNames.Text.Html));
                }

                mail.To.Add(message.ToAddress);

                if (!string.IsNullOrWhiteSpace(message.ReplyToAddress))
                    mail.ReplyToList.Add(message.ReplyToAddress);

                await client.SendMailAsync(mail, cancellationToken);

                // The SUBJECT is logged, never the body: a verification or reset body contains a
                // working credential, and the journal is readable by anyone who can read logs.
                this.thisLogger.LogInformation(
                    "Sent mail to {Recipient}: {Subject}", message.ToAddress, message.Subject);

                return true;
            }
            catch (Exception ex) when (ex is SmtpException or InvalidOperationException or IOException or FormatException)
            {
                // Logged at Warning rather than Error: a single undeliverable mail is not a
                // service fault, and the user's recovery path (ask for another verification mail)
                // is already there. A repeated pattern of these is what the project owner should notice,
                // which is what the log is for.
                this.thisLogger.LogWarning(
                    ex,
                    "Could not send mail to {Recipient} ({Subject}). The operation it accompanied " +
                    "still succeeded; the user can request another.",
                    message.ToAddress,
                    message.Subject);

                return false;
            }
        }
    }
}
