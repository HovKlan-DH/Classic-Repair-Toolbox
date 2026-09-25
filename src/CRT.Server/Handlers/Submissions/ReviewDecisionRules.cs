using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // WHETHER a review decision may be made at all (NewContributeStrategy.md Phase 5, task 5).
    //
    // *** APPROVE PUBLISHES, AND PUBLISHING CANNOT BE UNDONE. *** Task 7 was struck by the
    // maintainer, so no publish history is retained: a published file is overwritten in place and
    // a bad merge is fixed only by publishing a correction. That single fact is why these rules
    // live in their own unit-tested class rather than as a few `if`s inside an endpoint - each one
    // is the last thing between a wrong request and a data tree that cannot be restored.
    //
    // *** THE THREE OUTCOMES ARE NOT EQUALLY DANGEROUS, and the rules say so. ***
    //
    //   APPROVE          - writes the published tree. ADMINISTRATOR ONLY. Irreversible.
    //   REJECT           - ends the submission with a reason. A reviewer may. Recoverable: the
    //                      contributor still holds their draft locally, which is the entire point
    //                      of the local-first design.
    //   REQUEST CHANGES  - returns it as an editable draft with a comment. A reviewer may. The
    //                      cheapest outcome and the one the strategy says explicitly not to skip.
    //
    // A REVIEWER MAY REJECT AND RETURN BUT NEVER APPROVE. That is what makes Phase 6's role table
    // honest when it gives Reviewer a blast radius of "none - no published data can change": both
    // of the outcomes a reviewer can reach leave the published tree untouched. A role that could
    // only look would not reduce the administrator's workload at all, which is why the role exists.
    //
    // *** ONE METHOD PER QUESTION, rather than one "is this allowed" taking an outcome. *** The
    // answer genuinely differs per outcome, and a single method would force every caller to pass
    // an outcome enum and remember which arguments matter for which value. Three named methods
    // cannot be called for the wrong outcome by accident.
    //
    // Every refusal hands back a REASON. The review app shows it rather than silently not drawing
    // a button: a reviewer whose account lacks the role needs telling that, and a submission
    // somebody else already decided needs saying so rather than appearing broken.
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
        //
        // Administrator only. See the class header for why a Reviewer is refused here and nowhere
        // else.
        // ###########################################################################################
        public static bool CanApprove(AccountRecord? account, string? state, out string reason)
        {
            if (!ReviewAuthority.CanPublish(account))
            {
                reason = "This account is not allowed to publish. Ask an administrator to approve it.";
                return false;
            }

            return ReviewDecisionRules.IsUndecided(state, out reason);
        }

        // ###########################################################################################
        // May this account REJECT this submission, with a reason?
        //
        // A reviewer may: nothing published changes, and the contributor keeps their local draft.
        // ###########################################################################################
        public static bool CanReject(AccountRecord? account, string? state, out string reason)
        {
            if (!ReviewAuthority.CanReview(account))
            {
                reason = "This account is not allowed to review submissions.";
                return false;
            }

            return ReviewDecisionRules.IsUndecided(state, out reason);
        }

        // ###########################################################################################
        // May this account RETURN this submission to its contributor for changes?
        //
        // The same authority as rejecting, deliberately. If returning needed a higher role than
        // rejecting, the cheap outcome would be the harder one to reach and reviewers would reject
        // things that could have been a conversation - the exact failure the strategy warns about.
        // ###########################################################################################
        public static bool CanRequestChanges(AccountRecord? account, string? state, out string reason)
        {
            if (!ReviewAuthority.CanReview(account))
            {
                reason = "This account is not allowed to review submissions.";
                return false;
            }

            return ReviewDecisionRules.IsUndecided(state, out reason);
        }

        // ###########################################################################################
        // Is this submission still waiting for a decision?
        //
        // *** THE DOUBLE-DECISION INTERLOCK, and it matters most for Approve. *** Two reviewers
        // with the queue open both press it, or one presses twice over a slow link. Without this
        // the second publish rewrites the tree again - and with no revision history, re-running a
        // publish whose blobs have since been collected is not a harmless no-op.
        //
        // `approved` IS still actionable, which is the one non-obvious case. It means a reviewer
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
