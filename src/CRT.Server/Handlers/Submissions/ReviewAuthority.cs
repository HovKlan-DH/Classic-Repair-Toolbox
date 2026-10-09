using CRT.Server.Handlers.Accounts;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Who may review, and who may publish (NewContributeStrategy.md Phase 6, as decided by the
    // project owner on 2026-09-25: TWO roles).
    //
    // *** THE AUTHORITY QUESTION IS ANSWERED IN ONE PLACE, ON PURPOSE. *** Phase 6's own traps
    // say it outright: "Do not implement the administrator as a special case at each call site.
    // Compute authority once and use it everywhere. Special-casing invites a site that forgets the
    // check." This is that one place.
    //
    // *** THE QUESTION IS ALWAYS ABOUT ONE BOARD. *** "Is this account an administrator, OR in
    // THIS board's pool" - the exact phrasing the plan insists on. There is no "may review
    // somewhere" that grants anything about a particular submission; CanReviewAnything below
    // exists only so the queue can answer 403 to an account with no role at all rather than an
    // empty list.
    //
    // *** A SHARED-FILE CHANGE NEEDS TWO PEOPLE, NOT A DIFFERENT PERSON (owner decision,
    // 2026-09-25). *** A board's maintainer sees and approves a submission that changes a shared
    // file like any other - but their approval alone does not publish it; the administrator must
    // approve too. WHO must approve is ApprovalRules (CRT.Data); WHO may take part is this class.
    // It used to hide such submissions from maintainers altogether, with the administrator deciding
    // alone; the project owner asked for both.
    //
    // *** NO SECOND FACTOR (recorded risk, 2026-09-25). *** Threat 2 asks for TOTP before any
    // account may publish. The project owner chose to defer it and open publishing to maintainers
    // without it; the accepted risk is written into NewContributeStrategy.md's security review.
    // ###########################################################################################
    public static class ReviewAuthority
    {
        // ###########################################################################################
        // May this account open the Maintainer tab at all - see the queue, and be told
        // "nothing waiting" rather than "not allowed"?
        //
        // Administrators, and anyone in at least one pool. An account in no pool is refused the
        // queue outright, which is what tells the Maintainer tab to say the account lacks the role.
        // ###########################################################################################
        public static bool CanReviewAnything(ReviewAccess? access) =>
            ReviewAuthority.IsUsable(access) &&
            (access!.Account.IsAdministrator || access.MaintainerOf.Count > 0);

        // ###########################################################################################
        // May this account SEE and DECIDE this submission (or this board)? Deciding includes
        // giving one of two approvals a shared-file change needs - ApprovalRules says whether the
        // approval publishes.
        // ###########################################################################################
        public static bool CanReview(ReviewAccess? access, SubmissionRecord? submission) =>
            submission is not null && ReviewAuthority.CanReview(access, submission.BoardId);

        public static bool CanReview(ReviewAccess? access, string? boardId) =>
            ReviewAuthority.RoleIn(access, boardId) is not null;

        // ###########################################################################################
        // May this account APPROVE a publish of this submission or board - on its own, or as one
        // of two? The same answer as CanReview; kept as its own question because opening something
        // and writing the tree are different acts, and a later read-only role would split them.
        // ###########################################################################################
        public static bool CanPublish(ReviewAccess? access, SubmissionRecord? submission) =>
            ReviewAuthority.CanReview(access, submission);

        public static bool CanPublish(ReviewAccess? access, string? boardId) =>
            ReviewAuthority.CanReview(access, boardId);

        // ###########################################################################################
        // May this account PUBLISH THIS BOARD FROM BETA TO THE STABLE SOURCE? CanPublish - unless
        // the server lets only administrators do that (ServerOptions.ProductionPublishingAdministratorsOnly,
        // owner request, 2026-10-05). Seeing the board in the BETA queue, reading its plan, and
        // pushing it back or rejecting it are not this question: they stay CanPublish's.
        // ###########################################################################################
        public static bool CanPublishToProduction(ReviewAccess? access, string? boardId, bool administratorsOnly) =>
            ReviewAuthority.CanPublish(access, boardId) && (!administratorsOnly || access!.Account.IsAdministrator);

        // ###########################################################################################
        // May this account see the EMAIL ADDRESSES of this board's people - its maintainers, its
        // contributors, and whoever its history names (owner request, 2026-10-05: "I do not think that
        // normal maintainer should be able to see other email addresses if they are not set as
        // maintainer for that system. They should be able to see all mail addresses for their own
        // system(s)")? The administrator and the board's own maintainers: CanReview. Everybody else
        // on the Boards screen sees the same board with names and no addresses (BoardOverviewFlow).
        // ###########################################################################################
        public static bool CanSeeAddressesOf(ReviewAccess? access, string? boardId) =>
            ReviewAuthority.CanReview(access, boardId);

        // ###########################################################################################
        // The role this account approves as, for this board: Administrator for an administrator
        // (in every pool by definition), Maintainer for a member of the board's pool, null for
        // anyone else. What ApprovalRules counts.
        // ###########################################################################################
        public static ApproverRole? RoleIn(ReviewAccess? access, string? boardId)
        {
            if (!ReviewAuthority.IsUsable(access))
                return null;

            if (access!.Account.IsAdministrator)
                return ApproverRole.Administrator;

            return !string.IsNullOrWhiteSpace(boardId) && access.MaintainerOf.Contains(boardId)
                ? ApproverRole.Maintainer
                : null;
        }

        // ###########################################################################################
        // Can this pool member give the MAINTAINER half of a two-person approval (code review,
        // 2026-09-25)?
        //
        // Not every row in a pool can. An ADMINISTRATOR's row - granted directly since 2026-10-05, so
        // others see who maintains a board, or kept by an account made administrator by hand later
        // (the documented SQL step) - approves as the administrator (RoleIn); a locked or unverified
        // account cannot approve at all. Counting such a row as
        // "the board has a maintainer" made a shared-file change demand a maintainer approval nobody
        // could give, and the submission waited in 'approved' for ever. The same preconditions as
        // IsUsable, plus "not an administrator".
        // ###########################################################################################
        public static bool CanGiveMaintainerApproval(MaintainerRecord? maintainer) =>
            maintainer is not null && maintainer.IsVerified && !maintainer.IsLocked && !maintainer.IsAdministrator;

        // ###########################################################################################
        // May this account manage maintainer pools? Administrators only, and there is deliberately
        // no way for anyone else to become one - see INSTALLING.md on making the first.
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

            string board = string.IsNullOrWhiteSpace(submission?.BoardId) ? "this board" : submission!.BoardId;

            return $"This account is not a maintainer of {board}.";
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
