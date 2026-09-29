namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHETHER a review decision may be made at all (NewContributeStrategy.md Phase 5, task 5;
    // per-system authority from Phase 6, 2026-09-25).
    //
    // *** APPROVE PUBLISHES, AND PUBLISHING CANNOT BE UNDONE. *** Task 7 was struck by the
    // project owner, so no publish history is retained: a published file is overwritten in place and
    // a bad merge is fixed only by publishing a correction. That single fact is why these rules
    // live in their own unit-tested class rather than as a few `if`s inside an endpoint - each one
    // is the last thing between a wrong request and a data tree that cannot be restored.
    //
    // *** ALL THREE OUTCOMES NEED THE SAME AUTHORITY NOW: a maintainer OF THIS SYSTEM, or an
    // administrator. *** ReviewAuthority answers that, against the submission's own system and
    // whether it touches shared files. The four-role plan gave a recommend-only Reviewer the two
    // cheap outcomes and withheld Approve; the project owner collapsed the roles, so a maintainer
    // assigned to a system decides everything about it. What stays true is that rejecting and
    // returning never need MORE authority than approving - if they did, the cheap outcome would be
    // the harder one to reach and maintainers would reject things that could have been a
    // conversation.
    //
    // *** ONE METHOD PER QUESTION, rather than one "is this allowed" taking an outcome. *** Three
    // named methods cannot be called for the wrong outcome by accident, and a later role that
    // splits the outcomes again changes one method rather than an enum switch.
    //
    // Every refusal hands back a REASON. The Maintainer tab shows it rather than silently not drawing
    // a button: a maintainer whose account is not in this system's pool needs telling that, and a
    // submission somebody else already decided needs saying so rather than appearing broken.
    // ###########################################################################################
    public static class ReviewDecisionRules
    {
        // ###########################################################################################
        // A rejection reason must be long enough to tell the contributor something.
        //
        // *** THIS IS THE ENTIRE FEEDBACK CHANNEL. *** Contributors have no account and no other
        // way of learning what happened - the contact email and this message are all there is. A
        // rejection with no reason is indistinguishable from being ignored, and it is the thing
        // most likely to stop somebody contributing again.
        //
        // The floor is deliberately LOW. It rejects thoughtlessness ("no"), not brevity - a short
        // accurate sentence is a fine reason and this must not push anybody into padding.
        // ###########################################################################################
        public const int MinimumReasonLength = 10;

        // A database column and an email body. Unbounded text arriving in a request is a storage
        // problem, and refusing is cheaper than a truncation nobody sees happen.
        public const int MaximumReasonLength = 4000;

        // ###########################################################################################
        // May this account APPROVE - and therefore publish - this submission?
        // ###########################################################################################
        public static bool CanApprove(ReviewAccess? access, SubmissionRecord? submission, out string reason) =>
            ReviewDecisionRules.CanDecide(access, submission, out reason);

        // ###########################################################################################
        // May this account REJECT this submission, with a reason?
        // ###########################################################################################
        public static bool CanReject(ReviewAccess? access, SubmissionRecord? submission, out string reason) =>
            ReviewDecisionRules.CanDecide(access, submission, out reason);

        // ###########################################################################################
        // May this account RETURN this submission to its contributor for changes?
        // ###########################################################################################
        public static bool CanRequestChanges(ReviewAccess? access, SubmissionRecord? submission, out string reason) =>
            ReviewDecisionRules.CanDecide(access, submission, out reason);

        // ###########################################################################################
        // The one rule the three share: authority over THIS submission's system, then the state
        // interlock. Authority first, so an account that may not act is never told anything about
        // the submission's state.
        // ###########################################################################################
        private static bool CanDecide(ReviewAccess? access, SubmissionRecord? submission, out string reason)
        {
            if (submission is null)
            {
                reason = "No such submission.";
                return false;
            }

            if (!ReviewAuthority.CanPublish(access, submission))
            {
                reason = ReviewAuthority.DescribeRefusal(access, submission);
                return false;
            }

            return ReviewDecisionRules.IsUndecided(submission.State, out reason);
        }

        // ###########################################################################################
        // Is this submission still waiting for a decision?
        //
        // *** THE DOUBLE-DECISION INTERLOCK, and it matters most for Approve. *** Two maintainers
        // with the queue open both press it, or one presses twice over a slow link. Without this
        // the second publish rewrites the tree again - and with no revision history, re-running a
        // publish whose blobs have since been collected is not a harmless no-op.
        //
        // `approved` IS still actionable, which is the one non-obvious case. It means a maintainer
        // accepted the submission but publishing has not happened or did not finish; refusing it
        // would strand something already agreed to with no way forward. `merged` is the terminal
        // success state and is refused.
        // ###########################################################################################
        private static bool IsUndecided(string? state, out string reason)
        {
            reason = string.Empty;

            if (string.IsNullOrWhiteSpace(state))
            {
                reason = "This submission has no state and cannot be decided.";
                return false;
            }

            if (state == SubmissionState.Pending || state == SubmissionState.Approved)
                return true;

            if (state == SubmissionState.Uploading)
            {
                reason = "This submission is still uploading and is not ready to be reviewed.";
                return false;
            }

            reason = state == SubmissionState.Merged
                ? "This submission has already been published."
                : $"This submission has already been decided ({state}).";

            return false;
        }

        // ###########################################################################################
        // Is this a reason worth sending to the person who wrote the contribution?
        //
        // TRIMMED BEFORE MEASURING, so leading whitespace cannot satisfy the length check while
        // saying nothing.
        // ###########################################################################################
        public static bool IsUsableReason(string? text, out string reason)
        {
            reason = string.Empty;

            string trimmed = text?.Trim() ?? string.Empty;

            if (trimmed.Length == 0)
            {
                reason = "Say why, so the contributor knows what to do next.";
                return false;
            }

            if (trimmed.Length < ReviewDecisionRules.MinimumReasonLength)
            {
                reason =
                    $"Write at least {ReviewDecisionRules.MinimumReasonLength} characters - " +
                    "this is the only message the contributor receives.";

                return false;
            }

            if (trimmed.Length > ReviewDecisionRules.MaximumReasonLength)
            {
                reason = $"Keep this under {ReviewDecisionRules.MaximumReasonLength} characters.";
                return false;
            }

            return true;
        }
    }
}
