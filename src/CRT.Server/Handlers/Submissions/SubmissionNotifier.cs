using CRT.Server.Handlers.Email;
using Microsoft.Extensions.Logging;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // TELLS THE CONTRIBUTOR WHAT A MAINTAINER DECIDED (owner request, 2026-09-23).
    //
    // *** UNTIL NOW NOTHING DID. *** The Submit dialog asked for an email address and said it was
    // "used only to tell you whether your contribution was accepted"; the address was stored,
    // shown to the maintainer, and never sent to. EmailTemplates carried four account mails and
    // nothing for a submission. So the one channel that works when the contributor is not sitting
    // in front of CRT did not exist, and three user-visible strings promised it anyway.
    //
    // *** SENDING MUST NEVER FAIL THE DECISION. *** By the time this runs the decision is already
    // recorded - and for a publish, the data tree is already written. An exception here would turn
    // a completed, irreversible operation into an error the maintainer sees, and they would quite
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
        // `state` is the state just written, so this maps one-to-one onto what the maintainer did
        // rather than re-deriving it. An unrecognised state sends NOTHING: a new outcome added
        // server-side should be silent until somebody writes its mail, which is far better than
        // guessing and sending a contributor the wrong news about their own work.
        // ###########################################################################################
        public async Task NotifyDecisionAsync(
            string? contactEmail,
            string? systemName,
            string state,
            string? maintainerComment,
            bool amendedByMaintainer = false,
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
                // the maintainer a 500 on a publish that had in fact succeeded (2026-09-23). There is
                // no fault here whose right answer is an exception reaching the endpoint.
                EmailMessage? message = SubmissionNotifier.BuildMessage(
                    contactEmail.Trim(),
                    systemName,
                    state,
                    maintainerComment,
                    amendedByMaintainer);

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
        // Tells the people who can decide a submission that one is waiting (Phase 6 task 11).
        //
        // One mail per address, each failure logged and swallowed - the submission is already
        // queued, and a mailer that is down must not turn that into an error the contributor
        // sees. Duplicate and blank addresses are dropped so a maintainer who is also an
        // administrator hears once.
        // ###########################################################################################
        public async Task NotifyMaintainersAsync(
            IEnumerable<string?> maintainerAddresses,
            string? systemName,
            long submissionId,
            string? contributorSummary,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(maintainerAddresses);

            IEnumerable<string> addresses = maintainerAddresses
                .Where(address => !string.IsNullOrWhiteSpace(address))
                .Select(address => address!.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase);

            foreach (string address in addresses)
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(
                            EmailTemplates.SubmissionWaiting(address, systemName ?? string.Empty, submissionId, contributorSummary),
                            cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(
                        ex,
                        "Could not tell {Address} that submission {SubmissionId} for [{System}] is waiting.",
                        address, submissionId, systemName);
                }
            }
        }

        // ###########################################################################################
        // Tells the other half of a two-person approval that it is their turn (2026-09-25): a
        // shared-file change has one approval and waits for theirs. Nothing escapes.
        // ###########################################################################################
        public async Task NotifyApprovalNeededAsync(
            IEnumerable<string?> addresses,
            string? systemName,
            string what,
            string approvedBy,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(addresses);

            foreach (string address in addresses
                         .Where(address => !string.IsNullOrWhiteSpace(address))
                         .Select(address => address!.Trim())
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(EmailTemplates.ApprovalNeeded(address, systemName ?? string.Empty, what, approvedBy), cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(ex, "Could not tell {Address} that [{System}] needs their approval.", address, systemName);
                }
            }
        }

        // ###########################################################################################
        // Tells the administrators that a maintainer published a system to production. Same
        // contract as the rest of this class: nothing escapes.
        // ###########################################################################################
        public async Task NotifyProductionPublishAsync(
            IEnumerable<string?> administratorAddresses,
            string? systemName,
            string actor,
            string? revision,
            int filesCopied,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(administratorAddresses);

            foreach (string address in administratorAddresses
                         .Where(address => !string.IsNullOrWhiteSpace(address))
                         .Select(address => address!.Trim())
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(
                            EmailTemplates.PublishedToProduction(address, systemName ?? string.Empty, actor, revision, filesCopied),
                            cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(
                        ex, "Could not tell {Address} that [{System}] was published to production.", address, systemName);
                }
            }
        }

        // ###########################################################################################
        // Which mail a state deserves, or null for one that deserves none.
        //
        // *** ONLY THE THREE DECISIONS A MAINTAINER MAKES. *** "pending" and "uploading" are states a
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
            string? maintainerComment,
            bool amendedByMaintainer = false)
        {
            return (state ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                // *** TWO STAGES, TWO MAILS (2026-09-25). *** "merged" is the maintainer's approval,
                // which writes the BETA data; "published" is its board going out to production -
                // ProductionPromotionFlow's contributors, and the word ProductionPromotionRules
                // reports to a contributor once it has happened.
                "merged" => EmailTemplates.SubmissionPublishedToBeta(
                    toAddress, systemName ?? string.Empty, maintainerComment, amendedByMaintainer),

                "published" => EmailTemplates.SubmissionPublishedToSource(
                    toAddress, systemName ?? string.Empty),

                "changes_requested" => EmailTemplates.SubmissionChangesRequested(
                    toAddress, systemName ?? string.Empty, maintainerComment ?? string.Empty),

                "rejected" => EmailTemplates.SubmissionRejected(
                    toAddress, systemName ?? string.Empty, maintainerComment ?? string.Empty),

                _ => null
            };
        }
    }
}
