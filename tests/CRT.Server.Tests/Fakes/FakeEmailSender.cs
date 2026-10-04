using CRT.Server.Handlers.Email;

namespace CRT.Server.Tests.Fakes
{
    // ###########################################################################################
    // Collects the mail a flow would have sent, so a test can assert on it without a network call
    // (CLAUDE.md test rule 6). The same idea as MockMiniproRunner.
    //
    // WHICH mail is sent is the anti-enumeration design - a known address gets "already
    // registered", an unknown reset address gets nothing at all - so being able to assert "exactly
    // one mail, to this address, containing this link" is the point of this class.
    // ###########################################################################################
    public sealed class FakeEmailSender : IEmailSender
    {
        public List<EmailMessage> Sent { get; } = [];

        public EmailMessage? Last => this.Sent.Count == 0 ? null : this.Sent[^1];

        // False makes every send report that postfix refused it - the feedback route's failure.
        public bool Delivers { get; set; } = true;

        public Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default)
        {
            this.Sent.Add(message);
            return Task.FromResult(this.Delivers);
        }

        // ###########################################################################################
        // Pulls the token out of the most recent mail, so a test can redeem it exactly as the
        // recipient would - rather than reaching into the store for a value the real flow would
        // never have handed them.
        //
        // *** TWO SHAPES, BECAUSE THE TWO MAILS DIFFER ON PURPOSE. *** Verification carries a
        // clickable "...?token=..." link, because GET /verify is mapped and a click is all it
        // needs. A reset carries a bare CODE on its own line, because completing a reset needs a
        // new password and so cannot be a link - the URL that used to be printed there pointed at
        // an unmapped path and 404'd for every user who clicked it (2026-09-22).
        // ###########################################################################################
        public string ExtractTokenFromLastMail()
        {
            EmailMessage message = this.Last
                ?? throw new InvalidOperationException("No mail was sent.");

            int index = message.Body.IndexOf("token=", StringComparison.Ordinal);

            if (index >= 0)
            {
                string linkTail = message.Body[(index + "token=".Length)..];

                // The link ends at the first whitespace or line break.
                int linkEnd = linkTail.AsSpan().IndexOfAny(" \r\n\t");

                return Uri.UnescapeDataString(linkEnd < 0 ? linkTail : linkTail[..linkEnd]);
            }

            return FakeEmailSender.ExtractBareCode(message.Body);
        }

        // ###########################################################################################
        // The reset code: the one line that is nothing but an opaque token.
        //
        // Matched by SHAPE rather than by the sentence above it, so rewording the mail cannot
        // silently break every reset test. SecureToken.Create produces base64url, so the line is
        // long and made only of url-safe characters - no spaces, and nothing a sentence contains.
        // ###########################################################################################
        private static string ExtractBareCode(string body)
        {
            foreach (string line in body.Split('\n'))
            {
                string candidate = line.Trim();

                if (candidate.Length >= 20 &&
                    candidate.All(character =>
                        char.IsAsciiLetterOrDigit(character) || character is '-' or '_' or '='))
                {
                    return candidate;
                }
            }

            throw new InvalidOperationException($"No token in the mail body: {body}");
        }
    }
}
