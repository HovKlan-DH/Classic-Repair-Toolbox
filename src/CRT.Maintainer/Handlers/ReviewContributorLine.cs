using System;
using System.Collections.Generic;
using System.Globalization;
using Handlers.DataHandling;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // THE CONTRIBUTOR LINE above a submission's table (owner request, 2026-09-26): who sent it, and
    // how their other submissions went -
    //
    //     From dh@hinet.dk - [6] other submissions: [4] published, [1] waiting, [1] rejected
    //     From Dennis (dh@hinet.dk) - no other submissions
    //
    // The facts are the server's (CRT.Data's ReviewContributorFacts, built by ContributorHistory);
    // this only words them. A kind with no submissions is left out, so the line stays short. The
    // numbers are bold in brackets, like the lines under it - hence ReviewNoteRun pieces.
    // ###########################################################################################
    public static class ReviewContributorLine
    {
        // Null when the server said nothing (an older server): no line rather than a wrong one.
        public static ReviewNoteLine? For(ReviewContributorFacts? facts)
        {
            if (facts is null)
                return null;

            var runs = new List<ReviewNoteRun> { new($"From {ReviewContributorLine.Who(facts)} - ", IsCount: false) };

            int others = facts.Published + facts.Waiting + facts.ChangesRequested + facts.Rejected;

            if (others == 0)
            {
                runs.Add(new ReviewNoteRun("no other submissions", IsCount: false));
                return ReviewNoteLine.FromRuns(runs, ReviewNoteKind.Change);
            }

            ReviewContributorLine.AddCount(runs, others, others == 1 ? " other submission: " : " other submissions: ");

            bool first = true;

            foreach ((int count, string word) in new[]
            {
                (facts.Published, " published"),
                (facts.Waiting, " waiting"),
                (facts.ChangesRequested, " changes requested"),
                (facts.Rejected, " rejected")
            })
            {
                if (count == 0)
                    continue;

                if (!first)
                    runs.Add(new ReviewNoteRun(", ", IsCount: false));

                ReviewContributorLine.AddCount(runs, count, word);
                first = false;
            }

            return ReviewNoteLine.FromRuns(runs, ReviewNoteKind.Change);
        }

        // "Dennis (dh@hinet.dk)" for a signed-in contributor, the address alone for the ordinary one.
        private static string Who(ReviewContributorFacts facts)
        {
            string email = facts.Email?.Trim() ?? string.Empty;
            string name = facts.Name?.Trim() ?? string.Empty;

            if (email.Length == 0)
                return name.Length == 0 ? "an unknown contributor" : name;

            return name.Length == 0 ? email : $"{name} ({email})";
        }

        // "[N]" then its words, N bold.
        private static void AddCount(List<ReviewNoteRun> runs, int count, string after)
        {
            runs.Add(new ReviewNoteRun("[", IsCount: false));
            runs.Add(new ReviewNoteRun(count.ToString(CultureInfo.InvariantCulture), IsCount: true));
            runs.Add(new ReviewNoteRun("]" + after, IsCount: false));
        }
    }
}
