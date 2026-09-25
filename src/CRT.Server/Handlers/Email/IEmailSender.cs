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
    // ###########################################################################################
    public interface IEmailSender
    {
        Task SendAsync(EmailMessage message, CancellationToken cancellationToken = default);
    }

    // ###########################################################################################
    // One outbound message. Plain text only, deliberately.
    //
    // NO HTML BODY. Every mail this service sends is a sentence and a link. HTML would add a
    // second body to keep in step with the first, an escaping surface where user-supplied display
    // names meet markup, and a reason for mail clients to mangle or hide the URL - and it buys
    // nothing a plain-text link does not already do. A plain-text mail also renders identically
    // everywhere, including in the terminal mail clients some of this audience genuinely use.
    // ###########################################################################################
    public sealed record EmailMessage(string ToAddress, string Subject, string Body);
}
