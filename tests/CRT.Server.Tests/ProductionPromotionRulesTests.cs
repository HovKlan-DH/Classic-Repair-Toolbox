using CRT.Server.Handlers.Submissions;
using Handlers.DataHandling;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers ProductionPromotionRules - "is BETA ahead of production", and what a contributor is
    // told about a merged submission (2026-09-25).
    // ###########################################################################################
    public sealed class ProductionPromotionRulesTests
    {
        private static readonly DateTimeOffset Merged = new(2026, 9, 25, 10, 0, 0, TimeSpan.Zero);

        private static BoardRecord Board(string? beta, string? production) =>
            new("A/B/C", "A", "B", "C", beta is null ? null : "2026-September-25", true, beta, null, production, null);

        [Fact]
        public void A_board_whose_BETA_hash_differs_from_productions_is_awaiting_production()
        {
            Assert.True(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.Board("new", "old")));
            Assert.True(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.Board("new", null)));
        }

        [Fact]
        public void A_board_production_already_matches_is_not_awaiting()
        {
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.Board("same", "same")));
        }

        [Fact]
        public void A_board_never_published_through_the_pipeline_is_waiting_for_NOTHING()
        {
            // A shipped board nobody has touched has no BETA content hash; listing it would offer
            // every shipped board as "ahead of production".
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.Board(null, null)));
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(null));
        }

        [Fact]
        public void A_merged_submission_reads_as_PUBLISHED_once_its_board_went_to_production_after_it()
        {
            Assert.Equal("published", ProductionPromotionRules.ContributorFacingState(
                SubmissionState.Merged, ProductionPromotionRulesTests.Merged, ProductionPromotionRulesTests.Merged.AddHours(2)));

            // At the same instant counts - the promotion carried it.
            Assert.Equal("published", ProductionPromotionRules.ContributorFacingState(
                SubmissionState.Merged, ProductionPromotionRulesTests.Merged, ProductionPromotionRulesTests.Merged));
        }

        [Fact]
        public void A_merged_submission_stays_MERGED_while_production_is_older_or_never_published()
        {
            // Merged after the last promotion: it is in BETA only, and saying "published" would
            // tell the contributor their own data has it when it does not.
            Assert.Equal("merged", ProductionPromotionRules.ContributorFacingState(
                SubmissionState.Merged, ProductionPromotionRulesTests.Merged, ProductionPromotionRulesTests.Merged.AddHours(-2)));

            Assert.Equal("merged", ProductionPromotionRules.ContributorFacingState(
                SubmissionState.Merged, ProductionPromotionRulesTests.Merged, null));
        }

        [Theory]
        [InlineData(SubmissionState.Pending)]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.ChangesRequested)]
        public void Every_other_state_is_reported_as_it_is(string state)
        {
            Assert.Equal(state, ProductionPromotionRules.ContributorFacingState(
                state, ProductionPromotionRulesTests.Merged, ProductionPromotionRulesTests.Merged.AddDays(1)));
        }

        // ###########################################################################################
        // *** "returned": TAKEN BACK OUT OF BETA (code review, 2026-09-27). *** A rollback puts a
        // merged submission back to `pending` WITH the maintainer's reason, and reported as plain
        // "pending" CRT showed "Waiting for review" to somebody told "Published to BETA source",
        // with no word for what had happened.
        //
        // *** RECORDED, NOT INFERRED (code review, 2026-09-29). *** The rollback records when it
        // returned the row (migration 0015), at the same instant as the row's decided_utc.
        // ###########################################################################################
        [Fact]
        public void A_pending_submission_a_rollback_returned_reads_as_returned()
        {
            Assert.Equal(
                ProductionPromotionRules.ReturnedState,
                ProductionPromotionRules.ContributorFacingState(
                    SubmissionState.Pending, ProductionPromotionRulesTests.Merged, null, ProductionPromotionRulesTests.Merged));
        }

        // An ordinary submission waiting for its first review has no return record and stays
        // "pending" - a new arrival must never read as returned.
        [Fact]
        public void A_pending_submission_never_returned_is_simply_pending()
        {
            Assert.Equal(
                SubmissionState.Pending,
                ProductionPromotionRules.ContributorFacingState(SubmissionState.Pending, null, null, null));
        }

        // ###########################################################################################
        // *** THE CASE THE INFERENCE GOT WRONG. *** It read "returned" off "pending and carrying a
        // decision comment" - so a pending row given a comment any other way (a manual fix, a future
        // note to the contributor, an amendment) announced a rollback that never happened. With no
        // return record it is simply pending, whatever the comment says.
        // ###########################################################################################
        [Fact]
        public void A_pending_submission_with_a_comment_but_no_rollback_is_not_returned()
        {
            // The rule no longer takes the comment at all: nothing but the return record moves it.
            Assert.Equal(
                SubmissionState.Pending,
                ProductionPromotionRules.ContributorFacingState(
                    SubmissionState.Pending, ProductionPromotionRulesTests.Merged, null, returnedUtc: null));
        }

        // A decision AFTER the rollback moves decided_utc past the return record, which ends it: the
        // row is judged afresh, not "taken back out of BETA" for ever.
        [Fact]
        public void A_decision_after_the_rollback_ends_returned()
        {
            DateTimeOffset returned = ProductionPromotionRulesTests.Merged;

            Assert.Equal(
                SubmissionState.Pending,
                ProductionPromotionRules.ContributorFacingState(
                    SubmissionState.Pending, returned.AddMinutes(5), null, returned));
        }

        // Only PENDING turns into "returned": a submission rejected after being returned keeps its
        // own word.
        [Theory]
        [InlineData(SubmissionState.Rejected)]
        [InlineData(SubmissionState.ChangesRequested)]
        public void A_decided_submission_with_a_return_record_keeps_its_own_state(string state)
        {
            Assert.Equal(
                state,
                ProductionPromotionRules.ContributorFacingState(
                    state, ProductionPromotionRulesTests.Merged, null, ProductionPromotionRulesTests.Merged));
        }

        // -----------------------------------------------------------------------------------
        // WHOSE WORK A PROMOTION CARRIES (owner request, 2026-09-27)
        // -----------------------------------------------------------------------------------

        private static SubmissionRecord Submission(
            long id,
            string? email = "hest@mailscan.dk",
            string? summary = "Corrected U8.",
            DateTimeOffset? decided = null) =>
            new(
                id,
                "Commodore/C64/250407",
                AccountId: null,
                ContactEmail: email,
                UploadTokenHash: null,
                BaseRevision: "r1",
                State: SubmissionState.Merged,
                Summary: summary,
                FormatVersion: 1,
                CreatedUtc: ProductionPromotionRulesTests.Merged.AddDays(-10),
                ExpiresUtc: null,
                DecidedUtc: decided);

        // ###########################################################################################
        // *** NEWEST FIRST. *** The panel's list is unbounded and read from the top, and the most
        // recently accepted submission is the one a maintainer is most likely looking for.
        // ###########################################################################################
        [Fact]
        public void The_carried_submissions_are_newest_first()
        {
            IReadOnlyList<CarriedSubmission> carrying = ProductionPromotionRules.Carrying(
            [
                ProductionPromotionRulesTests.Submission(1, decided: ProductionPromotionRulesTests.Merged.AddDays(-3)),
                ProductionPromotionRulesTests.Submission(2, decided: ProductionPromotionRulesTests.Merged),
                ProductionPromotionRulesTests.Submission(3, decided: ProductionPromotionRulesTests.Merged.AddDays(-1)),
            ]);

            Assert.Equal([2, 3, 1], carrying.Select(submission => submission.Id));
        }

        // The contributor's own words travel as they are - the same text the review queue lists the
        // submission by, so a maintainer recognises what they approved.
        [Fact]
        public void A_carried_submission_carries_its_contributor_and_their_own_description()
        {
            CarriedSubmission carried = Assert.Single(ProductionPromotionRules.Carrying(
                [ProductionPromotionRulesTests.Submission(7, decided: ProductionPromotionRulesTests.Merged)]));

            Assert.Equal(7, carried.Id);
            Assert.Equal("hest@mailscan.dk", carried.ContactEmail);
            Assert.Equal("Corrected U8.", carried.Comment);
            Assert.Equal(ProductionPromotionRulesTests.Merged, carried.DecidedUtc);
        }

        // ###########################################################################################
        // Both fields are nullable on the record - a submission can be sent without an account and
        // saved with no description - and neither is substituted here: the WORDING says "(no
        // contact address)" / "(no description)", so a null must survive the trip to reach it.
        // ###########################################################################################
        [Fact]
        public void A_submission_with_no_address_or_description_still_travels()
        {
            CarriedSubmission carried = Assert.Single(ProductionPromotionRules.Carrying(
                [ProductionPromotionRulesTests.Submission(7, email: null, summary: null)]));

            Assert.Equal(string.Empty, carried.ContactEmail);
            Assert.Null(carried.Comment);
        }

        // A board brought level by hand carries nothing, which is a real answer rather than an error.
        [Fact]
        public void Carrying_nothing_is_an_empty_list_rather_than_null()
        {
            Assert.Empty(ProductionPromotionRules.Carrying([]));
            Assert.Empty(ProductionPromotionRules.Carrying(null));
        }

        // An undecided record cannot order by its decision time; it falls back to when it arrived
        // rather than sorting as the epoch and burying every real entry beneath it.
        [Fact]
        public void A_submission_with_no_decision_time_orders_by_when_it_arrived()
        {
            IReadOnlyList<CarriedSubmission> carrying = ProductionPromotionRules.Carrying(
            [
                ProductionPromotionRulesTests.Submission(1, decided: ProductionPromotionRulesTests.Merged.AddDays(-20)),
                ProductionPromotionRulesTests.Submission(2, decided: null),
            ]);

            // #2 has no decision time, so it orders by CreatedUtc (10 days back) - ahead of #1.
            Assert.Equal([2, 1], carrying.Select(submission => submission.Id));
        }

        // -----------------------------------------------------------------------------------
        // AwaitsAccount - the BETA button's badge (owner request, 2026-09-27)
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** ONCE THIS ACCOUNT HAS APPROVED, THE BOARD WAITS FOR SOMEBODY ELSE. *** A shared-file
        // promotion needs the maintainer AND the administrator; after the first of them, the badge
        // must stop counting it for that person and keep counting it for the other.
        // ###########################################################################################
        [Fact]
        public void A_board_this_account_already_approved_waits_for_the_other_approver()
        {
            GivenApproval[] given = [new(ApproverRole.Maintainer, "Anna", ProductionPromotionRulesTests.Merged, AccountId: 7)];

            Assert.False(ProductionPromotionRules.AwaitsAccount(given, accountId: 7));
            Assert.True(ProductionPromotionRules.AwaitsAccount(given, accountId: 1));
        }

        // Nothing approved yet: anyone who may publish it can act on it.
        [Fact]
        public void With_no_approval_given_the_board_waits_for_everyone_who_may_publish_it()
        {
            Assert.True(ProductionPromotionRules.AwaitsAccount([], accountId: 7));
            Assert.True(ProductionPromotionRules.AwaitsAccount(null, accountId: 7));
        }

        // ###########################################################################################
        // While only administrators publish to stable, an administrator's approval always completes
        // the publish - so a board waits for them even when they approved it before the setting was
        // switched on (code review, 2026-10-05).
        // ###########################################################################################
        [Fact]
        public void While_only_administrators_publish_the_board_waits_for_an_administrator_who_already_approved()
        {
            GivenApproval[] given = [new(ApproverRole.Administrator, "Dennis", ProductionPromotionRulesTests.Merged, AccountId: 1)];

            Assert.False(ProductionPromotionRules.AwaitsAccount(given, accountId: 1));
            Assert.True(ProductionPromotionRules.AwaitsAccount(given, accountId: 1, administratorsOnly: true));
        }
    }
}
