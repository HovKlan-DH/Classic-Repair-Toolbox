using System;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // "THE CONTRIBUTOR DISCARDED THEIR OWN DRAFT" as the maintainer reads it (owner request,
    // 2026-09-28: "it should be clearly visible on the system and for the maintainer(s) - both in the
    // BETA to PROD queue, but also in the normal queue. In this scenario the maintainer most likely
    // should push system back to normal queue, and clarify with contributor").
    //
    // One mark for a list row, and a sentence where the maintainer DECIDES - above the submission's
    // table, and on "Beta > Prod" - saying what it means and what to do: ask the contributor first.
    // CRT.Data's DraftDiscardContract explains why CRT reports it at all.
    //
    // *** IT NEVER SAYS "WITHDRAWN". *** Discarding a draft deletes the contributor's working copy; it
    // does not take back what was sent, and the submission can still be published. The sentence
    // says both, so nobody rejects good work on a guess.
    //
    // *** NEVER "THEIR" OR "THEM" FOR THE CONTRIBUTOR (owner request, 2026-10-01). *** "The
    // contributor", "the draft" - no pronoun at all.
    //
    // Dates in the one format every submission line uses (SubmissionReceiptPresenter.FormatDate).
    // ###########################################################################################
    public static class DraftDiscardWording
    {
        // The BETA list row's mark - the row is a system, which may carry the work of more than one.
        public const string ListMark = "Contributor discarded the draft";

        // The short mark on a queue row and a Systems submission.
        public static string Mark(DateTimeOffset discardedUtc) =>
            $"Contributor discarded the draft on {SubmissionReceiptPresenter.FormatDate(discardedUtc)}";

        // ###########################################################################################
        // Above an opened submission's table - before the maintainer decides on it.
        // ###########################################################################################
        public static string SubmissionWarning(DateTimeOffset discardedUtc, string? contactEmail) =>
            $"The contributor discarded the draft of this board on {SubmissionReceiptPresenter.FormatDate(discardedUtc)}. " +
            "That does not withdraw the submission, but it may no longer be wanted - " +
            $"check with the contributor{DraftDiscardWording.Address(contactEmail)} before approving it.";

        // ###########################################################################################
        // On "Beta > Prod", when the BETA state it would publish carries such a submission - the
        // owner's own advice: push it back to the queue and ask.
        // ###########################################################################################
        public static string BetaWarning(CarriedSubmission submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            string who = string.IsNullOrWhiteSpace(submission.ContactEmail) ? "The contributor" : submission.ContactEmail.Trim();
            string when = submission.DraftDiscardedUtc is DateTimeOffset at
                ? $" on {SubmissionReceiptPresenter.FormatDate(at)}"
                : string.Empty;

            return $"{who} discarded the draft of this board{when}, after it was accepted into BETA. " +
                "Consider pushing it back to the queue and checking with the contributor before publishing it to the stable source.";
        }

        // The Systems screen's history line for the event.
        public static string HistoryWhat(string submissionNumber) =>
            $"{submissionNumber} - the contributor discarded the draft";

        private static string Address(string? contactEmail) =>
            string.IsNullOrWhiteSpace(contactEmail) ? string.Empty : $" ({contactEmail.Trim()})";
    }
}
