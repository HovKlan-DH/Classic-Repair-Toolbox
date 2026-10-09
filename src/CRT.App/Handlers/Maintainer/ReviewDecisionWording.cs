using System;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // What a review decision SAYS - before it is sent, and after it lands (Phase 5, task 5).
    //
    // Pure, and separate from the tab for the reason every other presenter behind it is: a
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
        // (the server's SubmissionReplacementRules, 2026-09-26) - see TabMaintainer.QueueRefresh.cs.
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
        // The "please wait" over the whole window while a decision is on its way (owner report,
        // 2026-09-28: a small green "Working..." under the table left the screen looking hung while
        // an approval published). The same overlay and the same shape of sentence as a push-back or
        // a publish to production - see ProductionDisplay.PushingBackWait.
        //
        // *** AN APPROVAL SAYS WHETHER IT PUBLISHES, from the same ApprovalStatus the button's
        // own text is written from (ApprovalWording.ApproveButton). *** The first of two approvals
        // publishes nothing, and "Publishing to BETA" over it would promise what is not happening.
        // ###########################################################################################
        public static string Waiting(ReviewDecisionKind kind, ApprovalStatus? approval, string? boardId)
        {
            string board = string.IsNullOrWhiteSpace(boardId) ? "this board" : boardId;

            return kind switch
            {
                ReviewDecisionKind.Approve when approval is null || approval.ApprovalPublishes =>
                    $"Publishing {board} to BETA. The submission is being written into the BETA data - please wait until it is done.",

                ReviewDecisionKind.Approve =>
                    $"Recording your approval of this submission to {board} - please wait until it is done.",

                ReviewDecisionKind.Reject =>
                    $"Rejecting this submission to {board} and telling the contributor why - please wait until it is done.",

                _ =>
                    $"Sending your request for changes on {board} to the contributor - please wait until it is done."
            };
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
                    " Check it in CRT with the BETA data, then publish it from " + MaintainerScreenWording.BetaQueueQuoted + ".";
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
