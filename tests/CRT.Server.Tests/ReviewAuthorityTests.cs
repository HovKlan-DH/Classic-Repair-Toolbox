using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewAuthority - who may review, and who may publish, PER SYSTEM (Phase 6 roles as
    // the project owner decided them on 2026-09-25: Administrator, and Maintainer assigned to systems).
    //
    // *** THIS IS A SECURITY BOUNDARY, NOT A UI CONVENIENCE. *** Phase 6 task 6 is explicit that
    // the desktop app hiding a button is not enforcement, because the app is public source and an
    // attacker calls the API directly. These tests are the enforcement's own proof.
    //
    // THE ONES THAT MATTER MOST: a maintainer of one system gets NOTHING on another (threat 3's
    // "check the object, not just the verb"), and a submission changing SHARED FILES is refused to
    // every maintainer, because those files reach every board. The negative cases are deliberately
    // as thorough as the positive ones: an authority check is only worth what it REFUSES.
    // ###########################################################################################
    public sealed class ReviewAuthorityTests
    {
        private const string C64 = "Commodore/C64/250407";
        private const string C128 = "Commodore/C128/310378";

        private static AccountRecord Account(
            bool isAdministrator = false,
            bool isVerified = true,
            bool isLocked = false) =>
            new(
                Id: 1,
                Email: "someone@example.com",
                NormalisedEmail: "someone@example.com",
                PasswordHash: "hash",
                DisplayName: "Someone",
                IsVerified: isVerified,
                IsAdministrator: isAdministrator,
                IsLocked: isLocked,
                CreatedUtc: DateTimeOffset.UnixEpoch,
                LastLoginUtc: null);

        private static ReviewAccess Admin(bool verified = true, bool locked = false) =>
            ReviewAccess.For(ReviewAuthorityTests.Account(isAdministrator: true, isVerified: verified, isLocked: locked));

        private static ReviewAccess MaintainerOf(params string[] systems) =>
            ReviewAccess.For(ReviewAuthorityTests.Account(), systems);

        private static ReviewAccess Ordinary() => ReviewAccess.For(ReviewAuthorityTests.Account());

        private static SubmissionRecord Submission(string systemId = ReviewAuthorityTests.C64, bool touchesShared = false) =>
            new(
                Id: 1, SystemId: systemId, AccountId: null, ContactEmail: "c@example.com",
                UploadTokenHash: "h", BaseRevision: "r1", State: SubmissionState.Pending, Summary: "x",
                FormatVersion: 1, CreatedUtc: DateTimeOffset.UnixEpoch, ExpiresUtc: null, DecidedUtc: null,
                DecisionComment: null, TouchesSharedFiles: touchesShared);

        // -----------------------------------------------------------------------------------
        // Per system - the line that must not move
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_maintainer_of_a_system_may_review_and_publish_THAT_system()
        {
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64);

            Assert.True(ReviewAuthority.CanReview(maintainer, ReviewAuthorityTests.Submission()));
            Assert.True(ReviewAuthority.CanPublish(maintainer, ReviewAuthorityTests.Submission()));
        }

        [Fact]
        public void A_maintainer_of_ONE_system_gets_NOTHING_on_ANOTHER()
        {
            // *** THE MOST IMPORTANT ASSERTION IN THIS FILE. *** A maintainer's token must be
            // useless against systems they do not review - threat 2's "keep authority narrow".
            // Not seeing it, not rejecting it, not publishing it.
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64);
            SubmissionRecord other = ReviewAuthorityTests.Submission(ReviewAuthorityTests.C128);

            Assert.False(ReviewAuthority.CanReview(maintainer, other));
            Assert.False(ReviewAuthority.CanPublish(maintainer, other));
        }

        [Fact]
        public void A_maintainer_of_SEVERAL_systems_may_act_on_each_of_them()
        {
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64, ReviewAuthorityTests.C128);

            Assert.True(ReviewAuthority.CanPublish(maintainer, ReviewAuthorityTests.Submission(ReviewAuthorityTests.C64)));
            Assert.True(ReviewAuthority.CanPublish(maintainer, ReviewAuthorityTests.Submission(ReviewAuthorityTests.C128)));
        }

        [Fact]
        public void The_system_id_is_compared_EXACTLY_like_the_database_key()
        {
            // utf8mb4_bin since migration 0005: "commodore/c64/250407" is a different system,
            // and a case-folding comparison here would grant what the database refuses.
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf("commodore/c64/250407");

            Assert.False(ReviewAuthority.CanReview(maintainer, ReviewAuthorityTests.Submission(ReviewAuthorityTests.C64)));
        }

        // -----------------------------------------------------------------------------------
        // Shared files - the board's maintainer AND the administrator (2026-09-25)
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_shared_files_submission_is_OPEN_to_the_boards_maintainer_as_well_as_the_administrator()
        {
            // The project owner: "in case of changes to any shared file, then both the maintainer and
            // the admin should approve". So the maintainer takes part - it used to be hidden from
            // them. Whether one approval PUBLISHES is ApprovalRules', not this class's.
            SubmissionRecord shared = ReviewAuthorityTests.Submission(touchesShared: true);

            Assert.True(ReviewAuthority.CanReview(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64), shared));
            Assert.True(ReviewAuthority.CanReview(ReviewAuthorityTests.Admin(), shared));

            // ...and still not to a maintainer of another board.
            Assert.False(ReviewAuthority.CanReview(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C128), shared));
        }

        [Fact]
        public void The_role_an_account_approves_as_is_Administrator_Maintainer_or_none()
        {
            Assert.Equal(ApproverRole.Administrator, ReviewAuthority.RoleIn(ReviewAuthorityTests.Admin(), ReviewAuthorityTests.C64));
            Assert.Equal(ApproverRole.Maintainer, ReviewAuthority.RoleIn(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64), ReviewAuthorityTests.C64));
            Assert.Null(ReviewAuthority.RoleIn(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C128), ReviewAuthorityTests.C64));
            Assert.Null(ReviewAuthority.RoleIn(ReviewAuthorityTests.Admin(locked: true), ReviewAuthorityTests.C64));
        }

        // -----------------------------------------------------------------------------------
        // The administrator - in every pool by definition
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_administrator_may_review_and_publish_EVERY_system_without_being_in_any_pool()
        {
            // No "override" path and no rows: the administrator is a maintainer of every system by
            // definition, which is what Phase 6's traps mean by "compute authority once".
            ReviewAccess admin = ReviewAuthorityTests.Admin();

            Assert.Empty(admin.MaintainerOf);
            Assert.True(ReviewAuthority.CanPublish(admin, ReviewAuthorityTests.Submission(ReviewAuthorityTests.C64)));
            Assert.True(ReviewAuthority.CanPublish(admin, ReviewAuthorityTests.Submission(ReviewAuthorityTests.C128)));
        }

        [Fact]
        public void Only_a_verified_unlocked_non_administrator_pool_member_gives_the_MAINTAINER_half()
        {
            // An administrator in a pool approves AS the administrator (RoleIn), and a locked or
            // unverified account cannot approve at all - so none of them is "the board's maintainer"
            // a shared-file change waits for (code review, 2026-09-25).
            var maintainer = new MaintainerRecord(ReviewAuthorityTests.C64, 1, "Anna", "anna@example.com");

            Assert.True(ReviewAuthority.CanGiveMaintainerApproval(maintainer));
            Assert.False(ReviewAuthority.CanGiveMaintainerApproval(maintainer with { IsAdministrator = true }));
            Assert.False(ReviewAuthority.CanGiveMaintainerApproval(maintainer with { IsLocked = true }));
            Assert.False(ReviewAuthority.CanGiveMaintainerApproval(maintainer with { IsVerified = false }));
            Assert.False(ReviewAuthority.CanGiveMaintainerApproval(null));
        }

        [Fact]
        public void Only_an_administrator_may_administer()
        {
            Assert.True(ReviewAuthority.CanAdminister(ReviewAuthorityTests.Admin()));
            Assert.False(ReviewAuthority.CanAdminister(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64)));
            Assert.False(ReviewAuthority.CanAdminister(ReviewAuthorityTests.Ordinary()));
            Assert.False(ReviewAuthority.CanAdminister(null));
        }

        // ###########################################################################################
        // Publishing from BETA to the stable source (owner request, 2026-10-05: "for now ... Only
        // me, as admin, should be able to publish to stable"). While the server lets only
        // administrators, a maintainer of the very system is refused; switched off, it is
        // CanPublish exactly. Never wider than CanPublish either way - a locked administrator and a
        // maintainer of another system stay refused.
        // ###########################################################################################
        [Fact]
        public void While_only_administrators_publish_to_stable_a_maintainer_of_the_system_may_not()
        {
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64);

            Assert.False(ReviewAuthority.CanPublishToProduction(maintainer, ReviewAuthorityTests.C64, administratorsOnly: true));
            Assert.True(ReviewAuthority.CanPublishToProduction(maintainer, ReviewAuthorityTests.C64, administratorsOnly: false));

            Assert.True(ReviewAuthority.CanPublishToProduction(ReviewAuthorityTests.Admin(), ReviewAuthorityTests.C64, administratorsOnly: true));
            Assert.True(ReviewAuthority.CanPublishToProduction(ReviewAuthorityTests.Admin(), ReviewAuthorityTests.C64, administratorsOnly: false));

            foreach (bool administratorsOnly in new[] { true, false })
            {
                Assert.False(ReviewAuthority.CanPublishToProduction(maintainer, ReviewAuthorityTests.C128, administratorsOnly));
                Assert.False(ReviewAuthority.CanPublishToProduction(ReviewAuthorityTests.Admin(locked: true), ReviewAuthorityTests.C64, administratorsOnly));
                Assert.False(ReviewAuthority.CanPublishToProduction(ReviewAuthorityTests.Ordinary(), ReviewAuthorityTests.C64, administratorsOnly));
                Assert.False(ReviewAuthority.CanPublishToProduction(null, ReviewAuthorityTests.C64, administratorsOnly));
            }
        }

        // ###########################################################################################
        // Whose addresses an account sees on the Systems screen (owner request, 2026-10-05: "They
        // should be able to see all mail addresses for their own system(s)") - the system's own
        // maintainers and the administrator; a maintainer of ANOTHER system sees the system, not them.
        // ###########################################################################################
        [Fact]
        public void Only_a_systems_own_maintainers_and_the_administrator_see_its_addresses()
        {
            ReviewAccess maintainer = ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64);

            Assert.True(ReviewAuthority.CanSeeAddressesOf(maintainer, ReviewAuthorityTests.C64));
            Assert.False(ReviewAuthority.CanSeeAddressesOf(maintainer, ReviewAuthorityTests.C128));
            Assert.True(ReviewAuthority.CanSeeAddressesOf(ReviewAuthorityTests.Admin(), ReviewAuthorityTests.C128));
            Assert.False(ReviewAuthority.CanSeeAddressesOf(ReviewAuthorityTests.Admin(locked: true), ReviewAuthorityTests.C64));
            Assert.False(ReviewAuthority.CanSeeAddressesOf(null, ReviewAuthorityTests.C64));
        }

        // -----------------------------------------------------------------------------------
        // Nobody at all
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_ordinary_account_may_NOT_review_anything()
        {
            // Having an account does not make somebody a maintainer. Being in a pool does.
            ReviewAccess ordinary = ReviewAuthorityTests.Ordinary();

            Assert.False(ReviewAuthority.CanReviewAnything(ordinary));
            Assert.False(ReviewAuthority.CanReview(ordinary, ReviewAuthorityTests.Submission()));
            Assert.False(ReviewAuthority.CanPublish(ordinary, ReviewAuthorityTests.Submission()));
        }

        [Fact]
        public void No_account_at_all_may_do_nothing()
        {
            Assert.False(ReviewAuthority.CanReviewAnything(null));
            Assert.False(ReviewAuthority.CanReview(null, ReviewAuthorityTests.Submission()));
            Assert.False(ReviewAuthority.CanPublish(null, ReviewAuthorityTests.Submission()));
        }

        [Fact]
        public void A_missing_submission_is_never_reviewable()
        {
            Assert.False(ReviewAuthority.CanReview(ReviewAuthorityTests.Admin(), (SubmissionRecord?)null));
        }

        [Fact]
        public void CanReviewAnything_is_true_for_anyone_in_at_least_one_pool()
        {
            // What lets the queue answer 403 to an account with no role, and an empty list to a
            // maintainer whose systems simply have nothing waiting.
            Assert.True(ReviewAuthority.CanReviewAnything(ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C128)));
            Assert.True(ReviewAuthority.CanReviewAnything(ReviewAuthorityTests.Admin()));
        }

        // -----------------------------------------------------------------------------------
        // The preconditions underneath every role
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_LOCKED_administrator_may_neither_review_nor_publish()
        {
            // *** LOCKING MUST BITE ON THE VERY NEXT REQUEST. *** Phase 6's definition of done
            // requires removal to take effect immediately rather than at next login, and locking
            // is how access is withdrawn. An already-issued session token must stop working.
            ReviewAccess locked = ReviewAuthorityTests.Admin(locked: true);

            Assert.False(ReviewAuthority.CanReviewAnything(locked));
            Assert.False(ReviewAuthority.CanPublish(locked, ReviewAuthorityTests.Submission()));
            Assert.False(ReviewAuthority.CanAdminister(locked));
        }

        [Fact]
        public void A_LOCKED_maintainer_may_not_review_their_own_system()
        {
            ReviewAccess locked = ReviewAccess.For(
                ReviewAuthorityTests.Account(isLocked: true), [ReviewAuthorityTests.C64]);

            Assert.False(ReviewAuthority.CanReview(locked, ReviewAuthorityTests.Submission()));
        }

        [Fact]
        public void An_UNVERIFIED_account_may_do_nothing_whatever_its_pools()
        {
            // An unverified address is an unproven one - the account may belong to somebody who
            // never asked for it. Authority must not rest on an address nobody has confirmed.
            ReviewAccess unverified = ReviewAccess.For(
                ReviewAuthorityTests.Account(isVerified: false), [ReviewAuthorityTests.C64]);

            Assert.False(ReviewAuthority.CanReview(unverified, ReviewAuthorityTests.Submission()));
            Assert.False(ReviewAuthority.CanReviewAnything(ReviewAuthorityTests.Admin(verified: false)));
        }

        // -----------------------------------------------------------------------------------
        // The relationship between the two answers
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Reviewing_and_publishing_are_the_same_authority()
        {
            // The project owner's model: whoever may review a system may publish to it. Stated as a
            // property over every kind of caller and submission, so that a later change that
            // split them again does so deliberately.
            ReviewAccess?[] callers =
            [
                null,
                ReviewAuthorityTests.Ordinary(),
                ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C64),
                ReviewAuthorityTests.Admin(),
                ReviewAuthorityTests.Admin(locked: true)
            ];

            SubmissionRecord[] submissions =
            [
                ReviewAuthorityTests.Submission(ReviewAuthorityTests.C64),
                ReviewAuthorityTests.Submission(ReviewAuthorityTests.C128),
                ReviewAuthorityTests.Submission(touchesShared: true)
            ];

            foreach (ReviewAccess? caller in callers)
            foreach (SubmissionRecord submission in submissions)
            {
                Assert.Equal(
                    ReviewAuthority.CanReview(caller, submission),
                    ReviewAuthority.CanPublish(caller, submission));
            }
        }

        // -----------------------------------------------------------------------------------
        // The sentence a refused maintainer reads
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_refusal_names_the_system_for_a_maintainer_of_another_one()
        {
            string why = ReviewAuthority.DescribeRefusal(
                ReviewAuthorityTests.MaintainerOf(ReviewAuthorityTests.C128), ReviewAuthorityTests.Submission(ReviewAuthorityTests.C64));

            Assert.Contains(ReviewAuthorityTests.C64, why, StringComparison.Ordinal);
        }

        [Fact]
        public void The_refusal_for_an_account_with_no_role_matches_the_403_the_app_recognises()
        {
            // ReviewApiClient turns a 403 into this exact sentence; the server's own must agree.
            Assert.Equal(
                "This account is not allowed to review submissions.",
                ReviewAuthority.DescribeRefusal(ReviewAuthorityTests.Ordinary(), ReviewAuthorityTests.Submission()));
        }
    }
}
