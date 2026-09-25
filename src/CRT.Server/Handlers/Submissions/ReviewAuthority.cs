using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // Who may review, and who may publish (NewContributeStrategy.md Phase 5, task 2; Phase 6
    // tasks 4 and 6).
    //
    // *** THE AUTHORITY QUESTION IS ANSWERED IN ONE PLACE, ON PURPOSE. *** Phase 6's own traps
    // say it outright: "Do not implement the administrator as a special case at each call site.
    // Compute authority once and use it everywhere. Special-casing invites a site that forgets the
    // check." This is that one place, written at Phase 5 rather than retrofitted at Phase 6,
    // because a rule introduced after the call sites exist has to find them all.
    //
    // *** A REVIEWER MAY NOT PUBLISH. EVER. *** That is the single most important line here.
    // Phase 6's role table gives Reviewer a blast radius of "none - no published data can change",
    // and the plan says to resist any later request to let Reviewers "just publish the easy ones",
    // because it turns a zero-blast-radius role into a global publisher. The separation is
    // structural - two different questions with two different answers - rather than a UI that
    // hides a button.
    //
    // *** PHASE 5 IS SINGLE-REVIEWER, AND THAT IS WHY PER-SYSTEM SCOPE IS ABSENT RATHER THAN
    // FORGOTTEN. *** Phase 5's goal says "one reviewer only; roles come in Phase 6", so today
    // publishing authority means administrator. Phase 6 adds the per-system maintainer pool, and
    // when it does it changes THIS FILE - CanPublish gains the system id and the pool lookup, and
    // every call site inherits it. See the note on CanPublish itself.
    // ###########################################################################################
    public static class ReviewAuthority
    {
        // ###########################################################################################
        // May this account SEE the queue and open a submission?
        //
        // Administrators and reviewers both may: examining a submission is what the Reviewer role
        // exists for. A locked or unverified account may not - a locked account is one whose
        // access has been withdrawn, and acting on that must not wait for a token to expire
        // (Phase 6's definition of done requires removal to take effect immediately).
        // ###########################################################################################
        public static bool CanReview(AccountRecord? account) =>
            ReviewAuthority.IsUsable(account) &&
            (account!.IsAdministrator || account.IsReviewer);

        // ###########################################################################################
        // May this account PUBLISH - write the data tree that every user of a system downloads?
        //
        // Administrator only, today. A Reviewer is deliberately refused: see the class header.
        //
        // *** WHEN PHASE 6 ADDS MAINTAINERS, THIS METHOD GAINS THE SYSTEM ID *** and answers "is
        // this account in this system's maintainer pool, OR an administrator" - the exact phrasing
        // Phase 6's traps insist on. It must NOT grow a second method beside this one for the
        // maintainer case; that is the special-casing the plan warns against, and the signature
        // change is what forces every call site to be revisited rather than quietly keeping the
        // old, wider answer.
        // ###########################################################################################
        public static bool CanPublish(AccountRecord? account) =>
            ReviewAuthority.IsUsable(account) && account!.IsAdministrator;

        // ###########################################################################################
        // The preconditions any authority rests on, whatever the role.
        //
        // UNVERIFIED IS REFUSED because an unverified address is an unproven one - the account may
        // belong to somebody who never asked for it. LOCKED is refused because locking is how
        // access is withdrawn, and Phase 6 requires that to bite on the very next request rather
        // than at next login.
        // ###########################################################################################
        private static bool IsUsable(AccountRecord? account) =>
            account is not null && account.IsVerified && !account.IsLocked;
    }
}
