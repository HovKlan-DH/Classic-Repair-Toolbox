using System;
using System.Collections.Generic;
using System.Globalization;
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
        // The BETA list row's mark - the row is a board, which may carry the work of more than one.
        public const string ListMark = "Contributor discarded the draft";

        // The short mark on a queue row and a Boards submission.
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
        //
        // In the amber panel under the board's name, apart from the carried submissions' lines - so
        // it says WHICH one (code review, 2026-10-11): its number, and the description its carried
        // line shows. The contributor is in bold, as on that line (PersonRuns); with no address it
        // is "the contributor", plain.
        // ###########################################################################################
        public static IReadOnlyList<ReviewNoteRun> BetaWarningRuns(CarriedSubmission submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            string which = string.IsNullOrWhiteSpace(submission.Comment)
                ? $"#{submission.Id.ToString(CultureInfo.InvariantCulture)} - "
                : $"#{submission.Id.ToString(CultureInfo.InvariantCulture)} ({submission.Comment.Trim()}) - ";

            string when = submission.DraftDiscardedUtc is DateTimeOffset at
                ? $" on {SubmissionReceiptPresenter.FormatDate(at)}"
                : string.Empty;

            string rest =
                $" discarded the draft of this board{when}, after it was accepted into BETA. " +
                "Consider pushing it back to the queue and checking with the contributor before publishing it to the stable source.";

            return
            [
                PersonRuns.Plain(which),
                PersonRuns.PersonOr(submission.ContactEmail, "the contributor"),
                PersonRuns.Plain(rest)
            ];
        }

        // What BetaWarningRuns says, as one text.
        public static string BetaWarning(CarriedSubmission submission) =>
            PersonRuns.Text(DraftDiscardWording.BetaWarningRuns(submission));

        // The Boards screen's history line for the event.
        public static string HistoryWhat(string submissionNumber) =>
            $"{submissionNumber} - the contributor discarded the draft";

        private static string Address(string? contactEmail) =>
            string.IsNullOrWhiteSpace(contactEmail) ? string.Empty : $" ({contactEmail.Trim()})";
    }
}
