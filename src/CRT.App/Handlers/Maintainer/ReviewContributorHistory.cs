using System;
using System.Collections.Generic;
using System.Globalization;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // THE CONTRIBUTOR VIEW, in words (owner request, 2026-09-30: ""Contributor info" would show all
    // information about the contributor, to get an honest opinion if this person can be trusted, so
    // all data we can see for whatever he/she has contributed, and which has been accepted and
    // rejected etc. - the trustworthiness history of the contributor").
    //
    // Top to bottom:
    //
    //     Dennis (dh@hinet.dk)
    //     Sent with an account made 2026-September-2 - the address is verified.
    //     [6] submissions in total whereof [4] published to stable, [1] waiting and [1] rejected
    //
    //     Other submissions, newest first
    //       Fixed the U8 pinout                                          <- what was written
    //       #12 - Commodore / C64 / 250407 - Published to the BETA source - sent 2026-September-25
    //       Told the contributor: ...                                    <- only when told anything
    //
    // The facts are the server's (ContributorHistory, CRT.Data's ReviewContributorFacts); the state
    // words are CRT's own (SubmissionReceiptPresenter), so a submission reads the same here, on the
    // Boards screen and in the contributor's own Drafts tab. Pure, so the view is tested.
    //
    // *** THE ACCOUNT LINE IS THE ONE FACT NO COUNT GIVES. *** A submission sent without an account
    // carries an address somebody TYPED - anybody could have typed it - so its record belongs to
    // whoever uses that address, not provably to one person. Said plainly, since it is exactly what
    // an "honest opinion" of a contributor must weigh.
    // ###########################################################################################
    public static class ReviewContributorHistory
    {
        public static string Heading(ReviewContributorFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            return ReviewContributorLine.Who(facts);
        }

        // Null from a server that does not say (older than 3.7.0).
        public static string? AccountLine(ReviewContributorFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            return facts.SignedIn switch
            {
                true when facts.AccountCreatedUtc is DateTimeOffset since =>
                    $"Sent with an account made {SubmissionReceiptPresenter.FormatDate(since)} - the address is verified.",
                true => "Sent with an account - the address is verified.",
                false => "Sent without an account - the address was typed in when sending, and nothing checked it.",
                _ => null
            };
        }

        // "[6] submissions in total whereof [4] published to stable, ..." - the contributor line's
        // counts, without the "From ..." the heading already says.
        public static IReadOnlyList<ReviewNoteRun> Counts(ReviewContributorFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            var runs = new List<ReviewNoteRun>(ReviewContributorLine.Counts(facts));

            // It starts a line here, not after "From ... - ": "No other submissions".
            if (runs.Count > 0 && runs[0].Text.Length > 0)
                runs[0] = runs[0] with { Text = char.ToUpperInvariant(runs[0].Text[0]) + runs[0].Text[1..] };

            return runs;
        }

        // ###########################################################################################
        // The heading over the list. `Listed` is what the server sent; the counts are over ALL of
        // them, so when the server stopped at its limit the heading says it shows the newest ones.
        // Null list (an older server) says that nothing more can be shown.
        //
        // *** NULL - NO HEADING - WHEN THERE IS NOTHING AT ALL (code review, 2026-10-01). *** The
        // counts line just above already reads "No other submissions", and the heading said
        // "No other submissions." again under it. An empty list beside counts that are NOT zero
        // (the server counted what it did not list) still says so.
        // ###########################################################################################
        public static string? SubmissionsHeading(ReviewContributorFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            // *** NEVER "THEIR" (owner request, 2026-10-01: "never refer to a contributor as
            // "their"... like in "Their other submissions""). *** Each line names what it lists
            // without a pronoun for the contributor - "this contributor", or nothing at all.
            if (facts.Submissions is not IReadOnlyList<ContributorSubmissionEntry> listed)
                return "The server does not list this contributor's earlier submissions yet.";

            int all = facts.Published + facts.Waiting + facts.ChangesRequested + facts.Rejected;

            if (listed.Count == 0)
                return all == 0 ? null : "None of the other submissions could be listed.";

            return listed.Count < all
                ? $"The newest {listed.Count.ToString(CultureInfo.InvariantCulture)} of {all.ToString(CultureInfo.InvariantCulture)} other submissions"
                : "Other submissions, newest first";
        }

        // What the contributor wrote about it, or that they wrote nothing.
        public static string Title(ContributorSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.IsNullOrWhiteSpace(submission.Summary) ? "(no description given)" : submission.Summary.Trim();
        }

        // "#12 - Commodore / C64 / 250407 - Published to the BETA source - sent 2026-September-25".
        public static string Footer(ContributorSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            var parts = new List<string> { $"#{submission.Id.ToString(CultureInfo.InvariantCulture)}" };

            if (!string.IsNullOrWhiteSpace(submission.BoardId))
                parts.Add(submission.BoardId.Trim().Replace("/", " / ", StringComparison.Ordinal));

            parts.Add(SubmissionReceiptPresenter.DescribeState(submission.State));

            if (submission.CreatedUtc != default)
                parts.Add($"sent {SubmissionReceiptPresenter.FormatDate(submission.CreatedUtc)}");

            return string.Join(" - ", parts);
        }

        // What the contributor was told - the Boards screen's words - or null for nothing.
        public static string? Comment(ContributorSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.IsNullOrWhiteSpace(submission.DecisionComment)
                ? null
                : $"Told the contributor: {submission.DecisionComment.Trim()}";
        }

        // Whether a row's state is a turn-down - drawn in the failure colour, so a record with
        // rejections in it is seen at a glance rather than read line by line.
        public static bool IsTurnedDown(ContributorSubmissionEntry submission)
        {
            ArgumentNullException.ThrowIfNull(submission);

            return string.Equals(submission.State?.Trim(), "rejected", StringComparison.OrdinalIgnoreCase);
        }
    }
}
