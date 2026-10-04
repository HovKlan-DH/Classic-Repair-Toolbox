using CRT.Server.Handlers.Accounts;
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
    // *** A MISSING ADDRESS IS NORMAL, NOT AN ERROR. *** An older receipt may carry none. No
    // address simply means no mail; the outcome is still in "My submissions" either way. A
    // submission from a signed-in maintainer carries no ContactEmail either, which is why the
    // decision mails go through the SubmissionRecord overload, which writes to the account.
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
            CancellationToken cancellationToken = default,
            string? contributorName = null)
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
                    amendedByMaintainer,
                    contributorName);

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
        // The same mail for one SUBMISSION, written to its contributor wherever they are reached -
        // the account's address for a submission sent signed in, the typed one otherwise
        // (ContributorAddresses, the one rule).
        //
        // *** THE DECISION MAILS READ THE TYPED ADDRESS ALONE (owner request, 2026-10-01). *** A
        // signed-in submission carries none of its own, so its contributor heard nothing about an
        // approval, a rejection or a request for changes - unseen while CRT never sent one signed
        // in, and reached the moment CRT began sending a signed-in maintainer's submissions with
        // their account. The lookup sits inside the guard like everything else here: the decision
        // is already recorded, so a database fault costs the account's address (the typed one is
        // used, which may be none), never the request.
        // ###########################################################################################
        public async Task NotifyDecisionAsync(
            SubmissionRecord submission,
            IAccountStore accounts,
            string state,
            string? maintainerComment,
            bool amendedByMaintainer = false,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(accounts);

            MailRecipient recipient;

            try
            {
                AccountRecord? account = submission.AccountId is long accountId
                    ? await accounts.FindByIdAsync(accountId, cancellationToken).ConfigureAwait(false)
                    : null;

                recipient = ContributorAddresses.RecipientOf(submission, account);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(
                    ex,
                    "Could not read the account of submission {SubmissionId} for its review-outcome mail; using its contact address.",
                    submission.Id);

                recipient = new MailRecipient(submission.ContactEmail ?? string.Empty);
            }

            await this.NotifyDecisionAsync(
                recipient.Email,
                submission.SystemId,
                state,
                maintainerComment,
                amendedByMaintainer,
                cancellationToken,
                recipient.Name).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Tells a contributor their merged submission was rolled back out of BETA and is in the
        // queue again (owner decision, 2026-09-27). Its OWN method rather than a "pending" arm in
        // BuildMessage: "pending" is also the state a brand-new submission arrives in, and a state
        // word that names two different events is a trap for the next caller. The first version of
        // the rollback DID send NotifyDecisionAsync(Pending), which BuildMessage answered with null
        // - so no contributor would ever have been told, defeating the feature's whole purpose.
        //
        // Same guard shape as NotifyDecisionAsync: the rollback is already done, so a mail that
        // cannot be sent is logged, never thrown.
        // ###########################################################################################
        // ###########################################################################################
        // A board taken out of BETA (2026-09-28): the contributor is told their submission is back
        // in the queue - or, for Beta > Prod's "Reject", that it was rejected, in the queue's own
        // rejection mail, since that is what it now is.
        // ###########################################################################################
        public Task NotifyTakenOutOfBetaAsync(
            string? contactEmail,
            string? systemName,
            string maintainerComment,
            bool rejected,
            CancellationToken cancellationToken = default,
            string? contributorName = null) =>
            rejected
                ? this.NotifyDecisionAsync(contactEmail, systemName, SubmissionState.Rejected, maintainerComment, cancellationToken: cancellationToken, contributorName: contributorName)
                : this.NotifyReturnedToQueueAsync(contactEmail, systemName, maintainerComment, cancellationToken, contributorName);

        public async Task NotifyReturnedToQueueAsync(
            string? contactEmail,
            string? systemName,
            string maintainerComment,
            CancellationToken cancellationToken = default,
            string? contributorName = null)
        {
            if (string.IsNullOrWhiteSpace(contactEmail))
            {
                return;
            }

            try
            {
                EmailMessage message = EmailTemplates.SubmissionReturnedToQueue(
                    contactEmail.Trim(),
                    systemName ?? string.Empty,
                    maintainerComment,
                    contributorName);

                await this.thisMailer.SendAsync(message, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                this.thisLogger.LogWarning(
                    ex,
                    "Could not send the returned-to-queue mail for [{System}].",
                    systemName);
            }
        }

        // ###########################################################################################
        // Tells the people who can decide a submission that one is waiting (Phase 6 task 11) -
        // saying whether it is a completely NEW system or an update to an existing one, and greeting
        // each by the name on their account (owner request, 2026-10-03).
        //
        // One mail per address, each failure logged and swallowed - the submission is already
        // queued, and a mailer that is down must not turn that into an error the contributor
        // sees. Duplicate and blank addresses are dropped so a maintainer who is also an
        // administrator hears once.
        // ###########################################################################################
        public async Task NotifyMaintainersAsync(
            IEnumerable<MailRecipient> recipients,
            string? systemName,
            long submissionId,
            string? contributorSummary,
            bool isNewSystem = false,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(recipients);

            foreach (MailRecipient recipient in SubmissionNotifier.Distinct(recipients))
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(
                            EmailTemplates.SubmissionWaiting(
                                recipient.Email, systemName ?? string.Empty, submissionId, contributorSummary, isNewSystem, recipient.Name),
                            cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(
                        ex,
                        "Could not tell {Address} that submission {SubmissionId} for [{System}] is waiting.",
                        recipient.Email, submissionId, systemName);
                }
            }
        }

        // ###########################################################################################
        // Tells the other half of a two-person approval that it is their turn (2026-09-25): a
        // shared-file change has one approval and waits for theirs. Nothing escapes.
        // ###########################################################################################
        public async Task NotifyApprovalNeededAsync(
            IEnumerable<MailRecipient> recipients,
            string? systemName,
            string what,
            string approvedBy,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(recipients);

            foreach (MailRecipient recipient in SubmissionNotifier.Distinct(recipients))
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(
                            EmailTemplates.ApprovalNeeded(recipient.Email, systemName ?? string.Empty, what, approvedBy, recipient.Name),
                            cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(ex, "Could not tell {Address} that [{System}] needs their approval.", recipient.Email, systemName);
                }
            }
        }

        // ###########################################################################################
        // Tells the administrators that a maintainer published a system to production. Same
        // contract as the rest of this class: nothing escapes.
        // ###########################################################################################
        public async Task NotifyProductionPublishAsync(
            IEnumerable<MailRecipient> administrators,
            string? systemName,
            string actor,
            string? revision,
            int filesCopied,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(administrators);

            foreach (MailRecipient recipient in SubmissionNotifier.Distinct(administrators))
            {
                try
                {
                    await this.thisMailer
                        .SendAsync(
                            EmailTemplates.PublishedToProduction(recipient.Email, systemName ?? string.Empty, actor, revision, filesCopied, recipient.Name),
                            cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(
                        ex, "Could not tell {Address} that [{System}] was published to production.", recipient.Email, systemName);
                }
            }
        }

        // ###########################################################################################
        // Tells the contributors of a deleted system's open submissions that the system - and their
        // submission - is gone (owner decision, 2026-10-03). The deletion is already done, so
        // nothing escapes. Answers how many mails the mailer took, for the administrator's answer.
        // ###########################################################################################
        public async Task<int> NotifySystemDeletedAsync(
            IEnumerable<MailRecipient> contributors,
            string? systemName,
            string reason,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(contributors);

            int sent = 0;

            foreach (MailRecipient recipient in SubmissionNotifier.Distinct(contributors))
            {
                try
                {
                    bool went = await this.thisMailer
                        .SendAsync(
                            EmailTemplates.SystemDeleted(recipient.Email, systemName ?? string.Empty, reason, recipient.Name),
                            cancellationToken)
                        .ConfigureAwait(false);

                    if (went)
                        sent++;
                }
                catch (Exception ex)
                {
                    this.thisLogger.LogWarning(ex, "Could not tell {Address} that [{System}] was deleted.", recipient.Email, systemName);
                }
            }

            return sent;
        }

        // Blank addresses dropped, each address once (the first name given for it), trimmed.
        private static IEnumerable<MailRecipient> Distinct(IEnumerable<MailRecipient> recipients) =>
            recipients
                .Where(recipient => !string.IsNullOrWhiteSpace(recipient?.Email))
                .Select(recipient => recipient with { Email = recipient.Email.Trim() })
                .GroupBy(recipient => recipient.Email, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First());

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
            bool amendedByMaintainer = false,
            string? contributorName = null)
        {
            return (state ?? string.Empty).Trim().ToLowerInvariant() switch
            {
                // *** TWO STAGES, TWO MAILS (2026-09-25). *** "merged" is the maintainer's approval,
                // which writes the BETA data; "published" is its board going out to production -
                // ProductionPromotionFlow's contributors, and the word ProductionPromotionRules
                // reports to a contributor once it has happened.
                "merged" => EmailTemplates.SubmissionPublishedToBeta(
                    toAddress, systemName ?? string.Empty, maintainerComment, amendedByMaintainer, contributorName),

                "published" => EmailTemplates.SubmissionPublishedToSource(
                    toAddress, systemName ?? string.Empty, contributorName),

                "changes_requested" => EmailTemplates.SubmissionChangesRequested(
                    toAddress, systemName ?? string.Empty, maintainerComment ?? string.Empty, contributorName),

                "rejected" => EmailTemplates.SubmissionRejected(
                    toAddress, systemName ?? string.Empty, maintainerComment ?? string.Empty, contributorName),

                _ => null
            };
        }
    }
}
