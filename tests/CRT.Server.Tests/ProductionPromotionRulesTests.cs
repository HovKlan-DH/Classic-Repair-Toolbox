using CRT.Server.Handlers.Submissions;
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

        private static SystemRecord System(string? beta, string? production) =>
            new("A/B/C", "A", "B", "C", beta is null ? null : "2026-September-25", true, beta, null, production, null);

        [Fact]
        public void A_system_whose_BETA_hash_differs_from_productions_is_awaiting_production()
        {
            Assert.True(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.System("new", "old")));
            Assert.True(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.System("new", null)));
        }

        [Fact]
        public void A_system_production_already_matches_is_not_awaiting()
        {
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.System("same", "same")));
        }

        [Fact]
        public void A_system_never_published_through_the_pipeline_is_waiting_for_NOTHING()
        {
            // A shipped board nobody has touched has no BETA content hash; listing it would offer
            // every shipped board as "ahead of production".
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(ProductionPromotionRulesTests.System(null, null)));
            Assert.False(ProductionPromotionRules.IsAwaitingProduction(null));
        }

        [Fact]
        public void A_merged_submission_reads_as_PUBLISHED_once_its_system_went_to_production_after_it()
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
    }
}
