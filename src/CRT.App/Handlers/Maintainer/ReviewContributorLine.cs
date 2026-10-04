using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHO SENT A SUBMISSION, AND HOW THE CONTRIBUTOR'S OTHER SUBMISSIONS WENT, in words (owner
    // request, 2026-09-26) - the name and the counts the Contributor view opens with
    // (ReviewContributorHistory):
    //
    //     dh@hinet.dk / Dennis (dh@hinet.dk)
    //     [6] submissions in total whereof [3] published to stable, [1] published to BETA,
    //         [1] waiting and [1] rejected
    //
    // It was a one-line "From ... - [6] other submissions: ..." above the submission's table until
    // 2026-09-30, when the owner asked for it to go ("this should be available under
    // "Contributor""), and "[6] other submissions: [4] published, ..." until 2026-10-01 (owner
    // request: "submissions in total whereof [1] published to stable and [4] rejected"). The
    // facts are the server's (CRT.Data's ReviewContributorFacts, built by ContributorHistory); this
    // only words them. A kind with no submissions is left out, so the line stays short. The numbers
    // are bold in brackets - hence ReviewNoteRun pieces.
    //
    // *** THE TOTAL IS STILL THE OTHER SUBMISSIONS - the one being reviewed is not in it. *** It
    // is what the list under it shows, and what the kinds add up to; the submission on screen has
    // no outcome yet.
    // ###########################################################################################
    public static class ReviewContributorLine
    {
        // ###########################################################################################
        // "[6] submissions in total whereof [4] published to stable, [1] waiting and [1] rejected",
        // or "no other submissions" - the Contributor view starts a line with it
        // (ReviewContributorHistory.Counts).
        //
        // Published is split into stable and BETA when the server says how many reached stable
        // (PublishedToStable, server 3.9.0); an older server's count stays plain "published" -
        // calling a BETA-only one "published to stable" would be the one wrong thing to say.
        // ###########################################################################################
        public static IReadOnlyList<ReviewNoteRun> Counts(ReviewContributorFacts facts)
        {
            ArgumentNullException.ThrowIfNull(facts);

            var runs = new List<ReviewNoteRun>();

            int others = facts.Published + facts.Waiting + facts.ChangesRequested + facts.Rejected;

            if (others == 0)
            {
                runs.Add(new ReviewNoteRun("no other submissions", IsCount: false));
                return runs;
            }

            ReviewContributorLine.AddCount(runs, others, others == 1 ? " submission in total whereof " : " submissions in total whereof ");

            var kinds = new List<(int Count, string Word)>();

            if (facts.PublishedToStable is int stable)
            {
                int toStable = Math.Clamp(stable, 0, facts.Published);

                kinds.Add((toStable, " published to stable"));
                kinds.Add((facts.Published - toStable, " published to BETA"));
            }
            else
            {
                kinds.Add((facts.Published, " published"));
            }

            kinds.Add((facts.Waiting, " waiting"));
            kinds.Add((facts.ChangesRequested, " changes requested"));
            kinds.Add((facts.Rejected, " rejected"));

            List<(int Count, string Word)> shown = kinds.Where(kind => kind.Count > 0).ToList();

            for (int i = 0; i < shown.Count; i++)
            {
                // "a, b and c" - the last one joined by "and".
                if (i > 0)
                    runs.Add(new ReviewNoteRun(i == shown.Count - 1 ? " and " : ", ", IsCount: false));

                ReviewContributorLine.AddCount(runs, shown[i].Count, shown[i].Word);
            }

            return runs;
        }

        // "Dennis (dh@hinet.dk)" for a signed-in contributor, the address alone for the ordinary one.
        // The Contributor view's heading (ReviewContributorHistory.Heading).
        public static string Who(ReviewContributorFacts facts)
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
