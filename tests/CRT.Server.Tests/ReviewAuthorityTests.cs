using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ReviewAuthority - who may review, and who may publish.
    //
    // *** THIS IS A SECURITY BOUNDARY, NOT A UI CONVENIENCE. *** Phase 6 task 6 is explicit that
    // the desktop app hiding a button is not enforcement, because the app is public source and an
    // attacker calls the API directly. These tests are the enforcement's own proof.
    //
    // THE ONE THAT MATTERS MOST is that a Reviewer cannot publish. Phase 6's role table gives that
    // role a blast radius of "none - no published data can change", and the plan says outright to
    // resist any later request to let Reviewers "just publish the easy ones", because it converts
    // a zero-blast-radius role into a global publisher. If that test ever fails, the failure is
    // the point.
    //
    // The negative cases are deliberately as thorough as the positive ones: an authority check is
    // only worth what it REFUSES.
    // ###########################################################################################
    public sealed class ReviewAuthorityTests
    {
        private static AccountRecord Account(
            bool isAdministrator = false,
            bool isReviewer = false,
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
                IsReviewer: isReviewer,
                IsLocked: isLocked,
                CreatedUtc: DateTimeOffset.UnixEpoch,
                LastLoginUtc: null);

        // -----------------------------------------------------------------------------------
        // Reviewing
        // -----------------------------------------------------------------------------------

        [Fact]
        public void An_administrator_may_review()
        {
            Assert.True(ReviewAuthority.CanReview(ReviewAuthorityTests.Account(isAdministrator: true)));
        }

        [Fact]
        public void A_reviewer_may_review()
        {
            // The whole purpose of the role: vetting without publish rights.
            Assert.True(ReviewAuthority.CanReview(ReviewAuthorityTests.Account(isReviewer: true)));
        }

        [Fact]
        public void An_ordinary_account_may_NOT_review()
        {
            // Having an account does not make somebody a reviewer. Under the Phase 4 model an
            // account exists only where somebody deliberately granted authority, but a maintainer
            // of one system still must not be handed the whole queue by default.
            Assert.False(ReviewAuthority.CanReview(ReviewAuthorityTests.Account()));
        }

        [Fact]
        public void No_account_at_all_may_not_review()
        {
            Assert.False(ReviewAuthority.CanReview(null));
        }

        // -----------------------------------------------------------------------------------
        // Publishing - the line that must not move
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_REVIEWER_MAY_NOT_PUBLISH()
        {
            // *** THE MOST IMPORTANT ASSERTION IN THIS FILE. *** Phase 6's role table: Reviewer's
            // blast radius if the account is stolen is "none - no published data can change".
            // That holds only while this is false. A stolen Reviewer account should yield a
            // recommendation that a human still has to act on, never a publish.
            Assert.False(ReviewAuthority.CanPublish(ReviewAuthorityTests.Account(isReviewer: true)));
        }

        [Fact]
        public void An_administrator_may_publish()
        {
            Assert.True(ReviewAuthority.CanPublish(ReviewAuthorityTests.Account(isAdministrator: true)));
        }

        [Fact]
        public void An_ordinary_account_may_not_publish()
        {
            Assert.False(ReviewAuthority.CanPublish(ReviewAuthorityTests.Account()));
        }

        [Fact]
        public void No_account_at_all_may_not_publish()
        {
            Assert.False(ReviewAuthority.CanPublish(null));
        }

        [Fact]
        public void An_account_that_is_BOTH_administrator_and_reviewer_may_publish()
        {
            // The reviewer flag must not subtract authority - it is additive. An administrator who
            // was also marked a reviewer at some point is still an administrator.
            Assert.True(ReviewAuthority.CanPublish(
                ReviewAuthorityTests.Account(isAdministrator: true, isReviewer: true)));
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
            AccountRecord locked = ReviewAuthorityTests.Account(isAdministrator: true, isLocked: true);

            Assert.False(ReviewAuthority.CanReview(locked));
            Assert.False(ReviewAuthority.CanPublish(locked));
        }

        [Fact]
        public void A_LOCKED_reviewer_may_not_review()
        {
            Assert.False(ReviewAuthority.CanReview(
                ReviewAuthorityTests.Account(isReviewer: true, isLocked: true)));
        }

        [Fact]
        public void An_UNVERIFIED_administrator_may_neither_review_nor_publish()
        {
            // An unverified address is an unproven one - the account may belong to somebody who
            // never asked for it. Authority must not rest on an address nobody has confirmed.
            AccountRecord unverified = ReviewAuthorityTests.Account(isAdministrator: true, isVerified: false);

            Assert.False(ReviewAuthority.CanReview(unverified));
            Assert.False(ReviewAuthority.CanPublish(unverified));
        }

        [Fact]
        public void An_UNVERIFIED_reviewer_may_not_review()
        {
            Assert.False(ReviewAuthority.CanReview(
                ReviewAuthorityTests.Account(isReviewer: true, isVerified: false)));
        }

        // -----------------------------------------------------------------------------------
        // The relationship between the two answers
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Anyone_who_may_publish_may_also_review()
        {
            // Publishing without being able to open the thing first would be absurd, and a future
            // edit that made these two independent would produce exactly that. Stated as a
            // property so it survives whatever Phase 6 does to CanPublish.
            foreach (AccountRecord account in new[]
            {
                ReviewAuthorityTests.Account(isAdministrator: true),
                ReviewAuthorityTests.Account(isAdministrator: true, isReviewer: true)
            })
            {
                Assert.True(!ReviewAuthority.CanPublish(account) || ReviewAuthority.CanReview(account));
            }
        }
    }
}
