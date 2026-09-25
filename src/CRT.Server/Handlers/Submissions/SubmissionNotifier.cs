using CRT.Server.Handlers.Email;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // TELLS THE CONTRIBUTOR WHAT A REVIEWER DECIDED (maintainer request, 2026-09-23).
    //
    // *** UNTIL NOW NOTHING DID. *** The Submit dialog asked for an email address and said it was
    // "used only to tell you whether your contribution was accepted"; the address was stored,
    // shown to the reviewer, and never sent to. EmailTemplates carried four account mails and
    // nothing for a submission. So the one channel that works when the contributor is not sitting
    // in front of CRT did not exist, and three user-visible strings promised it anyway.
    //
    // *** SENDING MUST NEVER FAIL THE DECISION. *** By the time this runs the decision is already
    // recorded - and for a publish, the data tree is already written. An exception here would turn
    // a completed, irreversible operation into an error the reviewer sees, and they would quite
    // reasonably try again. So every failure is caught and logged, exactly as IEmailSender's own
    // header says ("sending must not fail the operation it accompanies").
    //
    // *** A MISSING ADDRESS IS NORMAL, NOT AN ERROR. *** A submission from a signed-in maintainer
    // carries no ContactEmail, and an older receipt may have none either. No address simply means
    // no mail; the outcome is still in "My submissions" either way.
    //
    // THE APP IS STILL THE PRIMARY CHANNEL. This is a notification, not a replacement: the mail
    // names "My submissions" rather than trying to reproduce the review screen in text.
    // ###########################################################################################
    public sealed class SubmissionNotifier
    {
        private readonly IEmailSender thisMailer;
        private readonly ILogger<SubmissionNotifier> thisLogger;

        public SubmissionNotifier(IEmailSender mailer, ILogger<SubmissionNotifier> logger)
        {
            this.thisMailer = mailer;
            this.thisLogger = logger;
        }

        // ###########################################################################################
        // Sends the mail for one decision, or does nothing at all.
        //
        // `state` is the state just written, so this maps one-to-one onto what the reviewer did
        // rather than re-deriving it. An unrecognised state sends NOTHING: a new outcome added
        // server-side should be silent until somebody writes its mail, which is far better than
        // guessing and sending a contributor the wrong news about their own work.
        // ###########################################################################################
        public async Task NotifyDecisionAsync(
            string? contactEmail,
            string? systemName,
            string state,
            string? reviewerComment,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(contactEmail))
            {
                return;
            }

            try
            {
                // *** BUILDING IS INSIDE THE GUARD TOO, not just sending. *** It is only string
                // formatting today, so nothing here is expected to throw - but this whole method
                // runs after a decision that is already recorded, and for a publish after the data
                // tree has already been overwritten. "Expected not to throw" is exactly the
                // reasoning that put DataChecksumManifest's scan outside its own try and answered
                // the reviewer a 500 on a publish that had in fact succeeded (2026-09-23). There is
                // no fault here whose right answer is an exception reaching the endpoint.
                EmailMessage? message = SubmissionNotifier.BuildMessage(
                    contactEmail.Trim(),
                    systemName,
                    state,
                    reviewerComment);

                if (message is null)
                {
                    return;
                }

                await this.thisMailer.SendAsync(message, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Logged with the submission's system so a contributor asking "I never heard
                // anything" can be answered from the log rather than guessed at.
                this.thisLogger.LogWarning(
                    ex,
                    "Could not send the review-outcome mail for [{System}] (state [{State}]).",
                    systemName,
                    state);
            }
        }

        // ###########################################################################################
        // Which mail a state deserves, or null for one that deserves none.
        //
        // *** ONLY THE THREE DECISIONS A REVIEWER MAKES. *** "pending" and "uploading" are states a
        // submission passes through on its own, and mailing somebody that their upload finished is
        // noise that teaches them to ignore the ones that matter. "approved" is deliberately
        // absent too: it means published-is-next, not published, and the contributor learns
        // nothing actionable from a mail that would be followed minutes later by the real one.
        //
        // Static and pure, so the mapping is unit tested without a mailer.
        // ###########################################################################################
        internal static EmailMessage? BuildMessage(
            string toAddress,
            string? systemName,
            string? state,
            string? reviewerComment)
        {
            return (state ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                // Both mean "it is in the published tree" - see DraftRetirement.IsPublishedState,
                // which draws the same line for the same reason.
                "published" or "merged" => EmailTemplates.SubmissionPublished(
                    toAddress, systemName ?? string.Empty, reviewerComment),

                "changes_requested" => EmailTemplates.SubmissionChangesRequested(
                    toAddress, systemName ?? string.Empty, reviewerComment ?? string.Empty),

                "rejected" => EmailTemplates.SubmissionRejected(
                    toAddress, systemName ?? string.Empty, reviewerComment ?? string.Empty),

                _ => null
            };
        }
    }
}
