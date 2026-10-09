using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT (owner request, 2026-09-28) - see CRT.Data's
    // DraftDiscardContract for why a maintainer is told.
    //
    // The endpoint has already proved the request holds the submission's capability token; this
    // records the notice and, the first time only, writes the audit row that puts it in the board's
    // history on the Boards screen (BoardHistoryRules shows DraftDiscarded audited under "#{id}").
    //
    // *** ANY STATE IS RECORDED. *** CRT only reports submissions still open, but a notice that
    // arrives late - the submission published meanwhile - is still true, and the screens decide what
    // it means for a state. Refusing it would only make CRT keep trying.
    //
    // Pure but for its two stores, like every flow here, so it tests against the fakes.
    // ###########################################################################################
    public static class DraftDiscardFlow
    {
        // True when this call recorded it; false when an earlier notice already had.
        public static async Task<bool> RecordAsync(
            SubmissionRecord submission,
            ISubmissionStore store,
            IAccountStore accounts,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(submission);
            ArgumentNullException.ThrowIfNull(store);
            ArgumentNullException.ThrowIfNull(accounts);

            bool recorded = await store.RecordDraftDiscardedAsync(submission.Id, nowUtc, cancellationToken);

            if (!recorded)
                return false;

            // Who did it, in the words the history names a contributor by: their address. Contributing
            // needs no account, so the address is usually all there is.
            string who = string.IsNullOrWhiteSpace(submission.ContactEmail)
                ? "the contributor"
                : submission.ContactEmail.Trim();

            await accounts.WriteAuditAsync(
                new AuditEntry(
                    submission.AccountId,
                    who,
                    BoardHistoryEvents.DraftDiscarded,
                    FormattableString.Invariant($"#{submission.Id}"),
                    null,
                    nowUtc),
                cancellationToken);

            return true;
        }
    }
}
