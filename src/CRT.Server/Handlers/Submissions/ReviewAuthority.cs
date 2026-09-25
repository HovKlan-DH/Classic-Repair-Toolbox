using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Who may review, and who may publish (NewContributeStrategy.md Phase 6, as decided by the
    // maintainer on 2026-09-25: TWO roles).
    //
    // *** THE AUTHORITY QUESTION IS ANSWERED IN ONE PLACE, ON PURPOSE. *** Phase 6's own traps
    // say it outright: "Do not implement the administrator as a special case at each call site.
    // Compute authority once and use it everywhere. Special-casing invites a site that forgets the
    // check." This is that one place.
    //
    // *** THE QUESTION IS ALWAYS ABOUT ONE SYSTEM. *** "Is this account an administrator, OR in
    // THIS system's pool" - the exact phrasing the plan insists on. There is no "may review
    // somewhere" that grants anything about a particular submission; CanReviewAnything below
    // exists only so the queue can answer 403 to an account with no role at all rather than an
    // empty list.
    //
    // *** A SHARED-FILE CHANGE NEEDS TWO PEOPLE, NOT A DIFFERENT PERSON (maintainer decision,
    // 2026-09-25). *** A board's reviewer sees and approves a submission that changes a shared
    // file like any other - but their approval alone does not publish it; the administrator must
    // approve too. WHO must approve is ApprovalRules (CRT.Data); WHO may take part is this class.
    // It used to hide such submissions from reviewers altogether, with the administrator deciding
    // alone; the maintainer asked for both.
    //
    // *** NO SECOND FACTOR (recorded risk, 2026-09-25). *** Threat 2 asks for TOTP before any
    // account may publish. The maintainer chose to defer it and open publishing to reviewers
    // without it; the accepted risk is written into NewContributeStrategy.md's security review.
    // ###########################################################################################
    public static class ReviewAuthority
    {
        // ###########################################################################################
        // May this account open the review application at all - see the queue, and be told
        // "nothing waiting" rather than "not allowed"?
        //
        // Administrators, and anyone in at least one pool. An account in no pool is refused the
        // queue outright, which is what tells the review app to say the account lacks the role.
        // ###########################################################################################
        public static bool CanReviewAnything(ReviewAccess? access) =>
            ReviewAuthority.IsUsable(access) &&
            (access!.Account.IsAdministrator || access.ReviewerOf.Count > 0);

        // ###########################################################################################
        // May this account SEE and DECIDE this submission (or this system)? Deciding includes
        // giving one of two approvals a shared-file change needs - ApprovalRules says whether the
        // approval publishes.
        // ###########################################################################################
        public static bool CanReview(ReviewAccess? access, SubmissionRecord? submission) =>
            submission is not null && ReviewAuthority.CanReview(access, submission.SystemId);

        public static bool CanReview(ReviewAccess? access, string? systemId) =>
            ReviewAuthority.RoleIn(access, systemId) is not null;

        // ###########################################################################################
        // May this account APPROVE a publish of this submission or system - on its own, or as one
        // of two? The same answer as CanReview; kept as its own question because opening something
        // and writing the tree are different acts, and a later read-only role would split them.
        // ###########################################################################################
        public static bool CanPublish(ReviewAccess? access, SubmissionRecord? submission) =>
            ReviewAuthority.CanReview(access, submission);

        public static bool CanPublish(ReviewAccess? access, string? systemId) =>
            ReviewAuthority.CanReview(access, systemId);

        // ###########################################################################################
        // The role this account approves as, for this system: Administrator for an administrator
        // (in every pool by definition), Reviewer for a member of the system's pool, null for
        // anyone else. What ApprovalRules counts.
        // ###########################################################################################
        public static ApproverRole? RoleIn(ReviewAccess? access, string? systemId)
        {
            if (!ReviewAuthority.IsUsable(access))
                return null;

            if (access!.Account.IsAdministrator)
                return ApproverRole.Administrator;

            return !string.IsNullOrWhiteSpace(systemId) && access.ReviewerOf.Contains(systemId)
                ? ApproverRole.Reviewer
                : null;
        }

        // ###########################################################################################
        // Can this pool member give the REVIEWER half of a two-person approval (code review,
        // 2026-09-25)?
        //
        // Not every row in a pool can. An account granted a pool and later made administrator BY
        // HAND (the documented SQL step) keeps its row, but approves as the administrator
        // (RoleIn); a locked or unverified account cannot approve at all. Counting such a row as
        // "the board has a reviewer" made a shared-file change demand a reviewer approval nobody
        // could give, and the submission waited in 'approved' for ever. The same preconditions as
        // IsUsable, plus "not an administrator".
        // ###########################################################################################
        public static bool CanGiveReviewerApproval(ReviewerRecord? reviewer) =>
            reviewer is not null && reviewer.IsVerified && !reviewer.IsLocked && !reviewer.IsAdministrator;

        // ###########################################################################################
        // May this account manage reviewer pools? Administrators only, and there is deliberately
        // no way for anyone else to become one - see DEPLOYMENT.md on granting the first.
        // ###########################################################################################
        public static bool CanAdminister(ReviewAccess? access) =>
            ReviewAuthority.IsUsable(access) && access!.Account.IsAdministrator;

        // ###########################################################################################
        // WHY a submission was refused to this account, written for the person reading it. One
        // place, so the endpoint, the publish flow and the decision rules cannot each explain the
        // same refusal differently.
        // ###########################################################################################
        public static string DescribeRefusal(ReviewAccess? access, SubmissionRecord? submission)
        {
            if (!ReviewAuthority.CanReviewAnything(access))
                return "This account is not allowed to review submissions.";

            string system = string.IsNullOrWhiteSpace(submission?.SystemId) ? "this system" : submission!.SystemId;

            return $"This account is not a reviewer of {system}.";
        }

        // ###########################################################################################
        // The preconditions any authority rests on, whatever the role.
        //
        // UNVERIFIED IS REFUSED because an unverified address is an unproven one - the account may
        // belong to somebody who never asked for it. LOCKED is refused because locking is how
        // access is withdrawn, and Phase 6 requires that to bite on the very next request rather
        // than at next login.
        // ###########################################################################################
        private static bool IsUsable(ReviewAccess? access) =>
            access is not null && access.Account.IsVerified && !access.Account.IsLocked;
    }
}
