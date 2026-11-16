using System;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A SUBMISSION SAYS IT CHANGED - its summary, which the maintainers read in the queue, in the
    // mail telling them of it ("The contributor wrote:") and in the board's history, and the contributor
    // in "My submissions".
    //
    // *** NOBODY TYPES IT ANY MORE (owner request, 2026-11-15: "we do have the "Main contribution" and the
    // board maintainer can see the details elsewhere, so there is no need to ALSO specify the detailed
    // changes in this field - so remove it"). *** The Submit dialog's "What did you change?" box went: the
    // summary is the credit's main contribution in plain words - what the contributor writes for the
    // board's history anyway. A draft whose work deserves no credit (the credits alone, the KiCad data, a
    // calibration, an order - CreditSuggestions) has no main contribution; its summary is then what the
    // draft changes, in the words "Review changes" uses. Never empty while there is either: the server
    // answers an empty summary with the warning "summary.missing", asking the contributor for a sentence
    // there is no box for.
    //
    // At most SubmissionFormat.MaximumSummaryLength, the server's limit (summary.too_long is an error),
    // cut between two characters - never inside one - with "..." saying it was cut.
    //
    // PURE: words in, words out.
    // ###########################################################################################
    public static class SubmissionSummaryText
    {
        // How many of the changes' lines stand in for a main contribution - the last then counting the rest.
        public const int ChangeLinesAtMost = 8;

        private const string Cut = "...";

        // ###########################################################################################
        // The summary to send: `mainContribution` (a credit's, as RichNote markup) in plain words, else
        // `changes` line by line, else nothing.
        // ###########################################################################################
        public static string For(string? mainContribution, DraftChangeSummary? changes)
        {
            string text = RichNote.PlainOf(mainContribution).Trim();

            if (text.Length == 0 && changes is { IsEmpty: false })
            {
                text = string.Join("\n", changes.AsLines(SubmissionSummaryText.ChangeLinesAtMost));
            }

            return SubmissionSummaryText.AtMost(text, SubmissionFormat.MaximumSummaryLength);
        }

        private static string AtMost(string text, int maxLength)
        {
            if (text.Length <= maxLength)
            {
                return text;
            }

            int cut = Math.Max(0, maxLength - SubmissionSummaryText.Cut.Length);

            // Not between the two halves of one character.
            if (cut > 0 && char.IsHighSurrogate(text[cut - 1]))
            {
                cut--;
            }

            return text[..cut].TrimEnd() + SubmissionSummaryText.Cut;
        }
    }
}
