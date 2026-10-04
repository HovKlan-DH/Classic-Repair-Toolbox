namespace CRT.Server.Handlers.Email
{
    // ###########################################################################################
    // The "send mail" seam. Exactly the same idea as IMiniproRunner and IScopeClient in CRT.App:
    // the interface is what the logic depends on, the real implementation is an untested I/O
    // boundary, and a fake stands in for it under test.
    //
    // This exists because CLAUDE.md test rule 6 forbids a test that needs a network call. Without
    // the seam, testing "registering sends exactly one verification mail" would mean either a real
    // SMTP connection or not testing it at all - and "which mail does this flow send, to whom,
    // containing what link" is precisely the logic most worth pinning down, because getting it
    // wrong either strands a user or mails a working credential to the wrong person.
    //
    // SENDING MUST NOT FAIL THE OPERATION IT ACCOMPANIES. If a verification mail cannot be sent,
    // the account has still been created, and answering with an error would tell the caller
    // something about the state of the mail system while leaving them unable to retry cleanly.
    // The implementation logs and swallows; "resend verification" is the recovery path.
    //
    // *** IT SAYS WHETHER THE MAIL WENT (2026-10-03). *** Feedback is the one case where the mail IS
    // the operation: a feedback whose mail never left reached nobody, and the sender must be told
    // so and try again (the old PHP page answered "Mail sending failed"). True means postfix took
    // the message; every other caller ignores it, as before.
    // ###########################################################################################
    public interface IEmailSender
    {
        Task<bool> SendAsync(EmailMessage message, CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // One outbound message: the HTML a mail client shows, and the same words as plain text.
    //
    // *** HTML SINCE 2026-10-03 (owner request: "All mails should be sent in HTML format"). *** It
    // was plain text only, to avoid a second body to keep in step and an escaping surface. Both are
    // answered by MailBody: a template writes its blocks ONCE and both bodies are rendered from
    // them, every typed value HTML-encoded on the way out. Body (the plain text) is kept - it is the
    // alternative part sent beside the HTML, which spam filters expect, and what tests read.
    // HtmlBody is null only for a message built by hand without MailBody; the sender then sends the
    // plain text alone.
    //
    // ReplyToAddress: where "Reply" goes, when it is not the From address - feedback, which comes
    // FROM the service (a mail claiming to be from the user's own address fails SPF at their
    // provider and lands in spam) but is answered to the user (2026-10-03).
    // ###########################################################################################
    public sealed record EmailMessage(string ToAddress, string Subject, string Body, string? HtmlBody = null, string? ReplyToAddress = null);

    // ###########################################################################################
    // Somebody to write to, and the name to greet them by (2026-10-03: "Hi {name}"). The name is
    // the one on their ACCOUNT - a maintainer's or administrator's, or a contributor who sent while
    // signed in - and blank for anybody without one, who is greeted "Hi there,".
    // ###########################################################################################
    public sealed record MailRecipient(string Email, string Name = "");
}
