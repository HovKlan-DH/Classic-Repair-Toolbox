using System;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // What a review decision SAYS - before it is sent, and after it lands (Phase 5, task 5).
    //
    // Pure, and separate from the window for the reason every other presenter in this app is: a
    // maintainer acts on these sentences, and approving cannot be undone. Wording that is tested is
    // wording that cannot quietly drift into saying something untrue.
    //
    // *** THE COMMENT RULE HERE IS A CONVENIENCE, NOT THE RULE. *** The server owns it
    // (ReviewDecisionRules.IsUsableReason) and refuses regardless of what this says. This exists
    // only so a maintainer is told to write more BEFORE a round trip, rather than after one. The
    // two are kept deliberately identical in effect, and a test pins the minimum so they cannot
    // drift into the client being the STRICTER of the two - which would refuse something the
    // server would have accepted, with no way for the maintainer to get past it.
    // ###########################################################################################
    public static class ReviewDecisionWording
    {
        // Matches ReviewDecisionRules.MinimumReasonLength on the server. Deliberately low: it
        // rejects thoughtlessness ("no"), not brevity.
        public const int MinimumCommentLength = 10;

        // The open submission left the queue while it was on screen - decided by another maintainer
        // or the administrator, or replaced by the contributor's newer submission of the same board
        // (the server's SubmissionReplacementRules, 2026-09-26) - see MaintainerMain.QueueRefresh.cs.
        // Its decision buttons are off. The queue check cannot tell which, so both are named.
        public const string DecidedElsewhere =
            "No longer in the queue - decided by someone else, or replaced by a newer submission from the same contributor.";

        // ###########################################################################################
        // Is this comment worth sending to the person who did the work?
        //
        // Trimmed before measuring, so leading whitespace cannot satisfy the length check while
        // saying nothing.
        // ###########################################################################################
        public static bool IsUsableComment(string? comment, out string problem)
        {
            problem = string.Empty;

            string trimmed = comment?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                problem = "Say why. This is the only message the contributor receives.";
                return false;
            }

            if (trimmed.Length < ReviewDecisionWording.MinimumCommentLength)
            {
                problem =
                    $"Write at least {ReviewDecisionWording.MinimumCommentLength} characters - " +
                    "this is the only message the contributor receives.";

                return false;
            }

            return true;
        }

        // ###########################################################################################
        // What happened, once the server has answered.
        //
        // *** AN APPROVAL NAMES THE PUBLISHED REVISION. *** That value is what every contributor's
        // next submission will diff against, and seeing it confirmed is how a maintainer knows the
        // publish reached the tree rather than merely being accepted.
        //
        // The state comes from the SERVER's answer rather than from the button that was pressed,
        // so the message describes what actually happened instead of what was asked for.
        // ###########################################################################################
        public static string Describe(ReviewDecisionKind kind, ReviewDecisionResult result)
        {
            ArgumentNullException.ThrowIfNull(result);

            if (kind == ReviewDecisionKind.Approve && string.Equals(result.State, "approved", StringComparison.Ordinal))
            {
                // The first of the two approvals a shared-file change needs: nothing published.
                return ApprovalWording.Recorded(result.WaitingFor, "BETA");
            }

            if (kind == ReviewDecisionKind.Approve)
            {
                // *** BETA, AND WHAT COMES NEXT (2026-09-25). *** An approval writes the BETA data
                // only; everyone gets it from "Production" once somebody has checked it there.
                // Saying just "Published" read as done.
                string published = string.IsNullOrWhiteSpace(result.Revision)
                    ? "Published to BETA."
                    : $"Published to BETA at revision {result.Revision}.";

                return published + FileRemovalWording.Done(result.RemovedFiles) +
                    " Check it in CRT with the BETA data, then publish it from Production.";
            }

            return kind == ReviewDecisionKind.Reject
                ? "Rejected. The contributor has been told why."
                : "Returned to the contributor for changes.";
        }
    }

    // ###########################################################################################
    // The three outcomes, as the window names them.
    //
    // An enum rather than three booleans or a string: the window branches on it three times, and
    // a typo in a string would send a decision to the wrong route - which on Approve is a publish
    // nobody asked for.
    // ###########################################################################################
    public enum ReviewDecisionKind
    {
        Approve,
        Reject,
        RequestChanges
    }
}
