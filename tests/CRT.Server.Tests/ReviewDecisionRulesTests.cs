using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewDecisionRules - whether a review decision may be made at all (Phase 5, task 5).
    //
    // *** APPROVE PUBLISHES, AND PUBLISHING CANNOT BE UNDONE. *** Task 7 was struck by the
    // project owner, so no publish history is retained: a published file is overwritten in place and
    // a bad merge is fixed only by publishing a correction. That single fact is why these rules
    // are a separate, unit-tested class rather than a few `if`s in an endpoint - every one of them
    // is the last thing standing between a wrong request and a data tree that cannot be restored.
    //
    // Three outcomes, and they are NOT equally dangerous:
    //
    //   APPROVE          - writes the published tree. Irreversible.
    //   REJECT           - ends the submission with a reason. Recoverable: the contributor still
    //                      has their draft locally, which is the whole point of the local-first
    //                      design.
    //   REQUEST CHANGES  - returns it to the contributor as an editable draft with a comment.
    //                      The cheapest outcome and the one the strategy says not to skip.
    //
    // Since Phase 6 (2026-09-25) all three need the same authority - a maintainer of the
    // submission's own system, or an administrator - and the per-system half of that is
    // ReviewAuthorityTests' subject. What is pinned here is that each outcome asks it, and the
    // state interlock on top.
    //
    // The rules are expressed as one method per question rather than a single "is this allowed",
    // so a caller cannot reach the wrong rule for the outcome it named.
    // ###########################################################################################
    public sealed class ReviewDecisionRulesTests
    {
        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";

        private static AccountRecord Account(
            bool administrator = false,
            bool verified = true,
            bool locked = false) =>
            new(
                Id: 1,
                Email: "someone@example.com",
                NormalisedEmail: "someone@example.com",
                PasswordHash: "hash",
                DisplayName: "Someone",
                IsVerified: verified,
                IsAdministrator: administrator,
                IsLocked: locked,
                CreatedUtc: DateTimeOffset.UnixEpoch,
                LastLoginUtc: null);

        private static ReviewAccess Admin(bool verified = true, bool locked = false) =>
            ReviewAccess.For(ReviewDecisionRulesTests.Account(administrator: true, verified: verified, locked: locked));

        // A maintainer of the C64 board the fixture submission names.
        private static ReviewAccess Maintainer() =>
            ReviewAccess.For(ReviewDecisionRulesTests.Account(), [ReviewDecisionRulesTests.C64]);

        private static ReviewAccess MaintainerOfAnotherSystem() =>
            ReviewAccess.For(ReviewDecisionRulesTests.Account(), [ReviewDecisionRulesTests.C128]);

        private static ReviewAccess Ordinary() => ReviewAccess.For(ReviewDecisionRulesTests.Account());

        private static SubmissionRecord Submission(string state, bool touchesShared = false) =>
            new(
                Id: 1, SystemId: ReviewDecisionRulesTests.C64, AccountId: null, ContactEmail: "c@example.com",
                UploadTokenHash: "h", BaseRevision: "r1", State: state, Summary: "x",
                FormatVersion: 1, CreatedUtc: DateTimeOffset.UnixEpoch, ExpiresUtc: null, DecidedUtc: null,
                DecisionComment: null, TouchesSharedFiles: touchesShared);

        // -----------------------------------------------------------------------------------
        // APPROVE - the irreversible one.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_ADMINISTRATOR_may_approve_a_pending_submission()
        {
            Assert.True(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Pending),
                out _));
        }

        [Fact]
        public void A_MAINTAINER_of_the_system_may_approve_a_pending_submission()
        {
            // The project owner's model: a maintainer assigned to a system publishes to it.
            Assert.True(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Maintainer(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Pending),
                out _));
        }

        [Fact]
        public void A_maintainer_of_ANOTHER_system_may_NOT_approve()
        {
            // *** THE STRUCTURAL GUARANTEE. *** A maintainer's authority is bounded to their own
            // systems; on any other it is exactly an ordinary account.
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.MaintainerOfAnotherSystem(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Pending),
                out string reason));

            // And the refusal says WHY, so the app can explain rather than silently not drawing
            // a button - it names the system the maintainer is not assigned to.
            Assert.Contains(ReviewDecisionRulesTests.C64, reason, StringComparison.Ordinal);
        }

        [Fact]
        public void A_shared_files_submission_may_be_approved_by_BOTH_its_maintainer_and_the_administrator()
        {
            // Since 2026-09-25 both take part; ApprovalRules decides whether an approval publishes
            // or waits for the other - these rules only say either may act.
            SubmissionRecord shared = ReviewDecisionRulesTests.Submission(SubmissionState.Pending, touchesShared: true);

            Assert.True(ReviewDecisionRules.CanApprove(ReviewDecisionRulesTests.Maintainer(), shared, out _));
            Assert.True(ReviewDecisionRules.CanApprove(ReviewDecisionRulesTests.Admin(), shared, out _));
        }

        [Fact]
        public void A_half_approved_submission_can_still_be_decided()
        {
            // 'approved' is where the first of two approvals leaves it; the second approval, or a
            // rejection by either, must still be possible.
            SubmissionRecord half = ReviewDecisionRulesTests.Submission(SubmissionState.Approved, touchesShared: true);

            Assert.True(ReviewDecisionRules.CanApprove(ReviewDecisionRulesTests.Admin(), half, out _));
            Assert.True(ReviewDecisionRules.CanReject(ReviewDecisionRulesTests.Maintainer(), half, out _));
        }

        [Fact]
        public void A_missing_submission_cannot_be_approved()
        {
            Assert.False(ReviewDecisionRules.CanApprove(ReviewDecisionRulesTests.Admin(), null, out string reason));
            Assert.NotEmpty(reason);
        }

        [Theory]
        [InlineData(SubmissionState.Merged)]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.ChangesRequested)]
        [InlineData(SubmissionState.Withdrawn)]
        public void An_ALREADY_DECIDED_submission_cannot_be_approved_again(string state)
        {
            // *** THE DOUBLE-PUBLISH GUARD, and the reason it matters more here than it looks. ***
            // Two maintainers with the queue open both click Approve; or one clicks twice on a slow
            // link. Without this the second publish overwrites the tree again - and with no
            // revision history, re-running a publish whose blobs have since been garbage-collected
            // is not a no-op. The state is the interlock.
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(), ReviewDecisionRulesTests.Submission(state), out string reason));

            Assert.NotEmpty(reason);
        }

        [Fact]
        public void An_UPLOADING_submission_cannot_be_approved()
        {
            // It is not finished arriving. Publishing a half-uploaded submission would write a
            // board referencing blobs that were never sent.
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Uploading),
                out _));
        }

        [Fact]
        public void An_ABANDONED_submission_cannot_be_approved()
        {
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Abandoned),
                out _));
        }

        [Fact]
        public void An_APPROVED_but_unpublished_submission_CAN_still_be_approved()
        {
            // `approved` means a maintainer accepted it but publishing has not happened or did not
            // finish. That is precisely the state a retry must be allowed from - refusing it would
            // strand a submission that has been agreed to, with no way forward.
            Assert.True(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(),
                ReviewDecisionRulesTests.Submission(SubmissionState.Approved),
                out _));
        }

        [Fact]
        public void A_LOCKED_administrator_may_not_approve()
        {
            // Locking is how access is withdrawn, and it must bite on the very next request rather
            // than at next login - Phase 6's definition of done requires exactly that.
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(locked: true),
                ReviewDecisionRulesTests.Submission(SubmissionState.Pending),
                out _));
        }

        [Fact]
        public void An_UNVERIFIED_administrator_may_not_approve()
        {
            // An unverified address is an unproven one - the account may belong to somebody who
            // never asked for it.
            Assert.False(ReviewDecisionRules.CanApprove(
                ReviewDecisionRulesTests.Admin(verified: false),
                ReviewDecisionRulesTests.Submission(SubmissionState.Pending),
                out _));
        }

        [Fact]
        public void A_NULL_account_may_not_approve()
        {
            Assert.False(ReviewDecisionRules.CanApprove(null, ReviewDecisionRulesTests.Submission(SubmissionState.Pending), out _));
        }

        // -----------------------------------------------------------------------------------
        // REJECT and REQUEST CHANGES - the same per-system authority as approving.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_MAINTAINER_may_REJECT()
        {
            // Rejecting changes no published data - the contributor keeps their draft locally.
            Assert.True(ReviewDecisionRules.CanReject(
                ReviewDecisionRulesTests.Maintainer(), ReviewDecisionRulesTests.Submission(SubmissionState.Pending), out _));
        }

        [Fact]
        public void A_MAINTAINER_may_REQUEST_CHANGES()
        {
            // The outcome the strategy says explicitly not to skip: most imperfect contributions
            // are fixable by their author in a minute, and a reject that could have been a
            // conversation costs a contributor.
            Assert.True(ReviewDecisionRules.CanRequestChanges(
                ReviewDecisionRulesTests.Maintainer(), ReviewDecisionRulesTests.Submission(SubmissionState.Pending), out _));
        }

        [Fact]
        public void An_ADMINISTRATOR_may_do_both_as_well()
        {
            ReviewAccess admin = ReviewDecisionRulesTests.Admin();

            Assert.True(ReviewDecisionRules.CanReject(admin, ReviewDecisionRulesTests.Submission(SubmissionState.Pending), out _));
            Assert.True(ReviewDecisionRules.CanRequestChanges(admin, ReviewDecisionRulesTests.Submission(SubmissionState.Pending), out _));
        }

        [Fact]
        public void An_ORDINARY_account_may_do_NONE_of_the_three()
        {
            // Somebody with an account but no review role. They can contribute; they cannot judge.
            ReviewAccess ordinary = ReviewDecisionRulesTests.Ordinary();
            SubmissionRecord pending = ReviewDecisionRulesTests.Submission(SubmissionState.Pending);

            Assert.False(ReviewDecisionRules.CanApprove(ordinary, pending, out _));
            Assert.False(ReviewDecisionRules.CanReject(ordinary, pending, out _));
            Assert.False(ReviewDecisionRules.CanRequestChanges(ordinary, pending, out _));
        }

        [Fact]
        public void A_maintainer_of_ANOTHER_system_may_do_NONE_of_the_three_either()
        {
            // Rejecting a submission on a board you do not review is as out of bounds as
            // publishing to it - the cheap outcomes are scoped exactly like the expensive one.
            ReviewAccess other = ReviewDecisionRulesTests.MaintainerOfAnotherSystem();
            SubmissionRecord pending = ReviewDecisionRulesTests.Submission(SubmissionState.Pending);

            Assert.False(ReviewDecisionRules.CanApprove(other, pending, out _));
            Assert.False(ReviewDecisionRules.CanReject(other, pending, out _));
            Assert.False(ReviewDecisionRules.CanRequestChanges(other, pending, out _));
        }

        [Theory]
        [InlineData(SubmissionState.Merged)]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.ChangesRequested)]
        [InlineData(SubmissionState.Withdrawn)]
        public void An_ALREADY_DECIDED_submission_cannot_be_rejected_or_returned_either(string state)
        {
            // A decided submission is done. Rejecting an already-merged one would tell the
            // contributor their published work was refused.
            ReviewAccess admin = ReviewDecisionRulesTests.Admin();

            Assert.False(ReviewDecisionRules.CanReject(admin, ReviewDecisionRulesTests.Submission(state), out _));
            Assert.False(ReviewDecisionRules.CanRequestChanges(admin, ReviewDecisionRulesTests.Submission(state), out _));
        }

        // -----------------------------------------------------------------------------------
        // The REASON a contributor is given.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_REJECTION_MUST_carry_a_reason()
        {
            // *** THE WHOLE CHANNEL. *** Contributors have no account and no other way of finding
            // out what happened; the contact email and this message are the entire feedback loop.
            // A rejection with no reason is indistinguishable from being ignored, and it is the
            // thing most likely to make somebody never contribute again.
            Assert.False(ReviewDecisionRules.IsUsableReason("", out _));
            Assert.False(ReviewDecisionRules.IsUsableReason("   ", out _));
            Assert.False(ReviewDecisionRules.IsUsableReason(null, out _));
        }

        [Fact]
        public void A_reason_that_is_too_SHORT_to_be_useful_is_refused()
        {
            // "no" is a reason in the sense that a field is populated, and useless to the person
            // who spent an evening documenting a board. The floor is deliberately low - this
            // rejects thoughtlessness, not brevity.
            Assert.False(ReviewDecisionRules.IsUsableReason("no", out _));
        }

        [Fact]
        public void An_ORDINARY_reason_is_accepted()
        {
            Assert.True(ReviewDecisionRules.IsUsableReason(
                "The highlight coordinates for U8 look wrong - did you mean the other pin?", out _));
        }

        [Fact]
        public void A_reason_is_TRIMMED_so_leading_whitespace_cannot_pass_the_length_check()
        {
            // Otherwise "          x" satisfies a raw length test while saying nothing.
            Assert.False(ReviewDecisionRules.IsUsableReason("          x", out _));
        }

        [Fact]
        public void An_ABSURDLY_LONG_reason_is_refused_rather_than_stored()
        {
            // This goes into a database column and into an email. Unbounded text from a request
            // is a storage problem, and the refusal is cheaper than a truncation nobody sees.
            Assert.False(ReviewDecisionRules.IsUsableReason(new string('x', 10_000), out _));
        }

        [Fact]
        public void The_minimum_reason_length_is_the_value_the_REVIEW_APP_also_uses()
        {
            // *** A CONTRACT ACROSS TWO PROJECTS THAT DO NOT REFERENCE EACH OTHER. ***
            // ReviewDecisionWording.MinimumCommentLength in CRT.Maintainer checks the same rule
            // locally, purely so a maintainer is told to write more BEFORE a round trip. The server
            // owns the rule and refuses regardless.
            //
            // What must never happen is the CLIENT becoming the stricter of the two: it would
            // refuse something this method would have accepted, and the maintainer would have no way
            // past it. The maintainer app cannot reference CRT.Server, so this literal is the pin -
            // change one and this test names the other.
            Assert.Equal(10, ReviewDecisionRules.MinimumReasonLength);
        }

        [Fact]
        public void The_refusal_message_SAYS_what_is_wrong()
        {
            // Shown to the maintainer who typed it, so it has to be actionable rather than "invalid".
            ReviewDecisionRules.IsUsableReason("no", out string reason);

            Assert.NotEmpty(reason);
        }
    }
}
