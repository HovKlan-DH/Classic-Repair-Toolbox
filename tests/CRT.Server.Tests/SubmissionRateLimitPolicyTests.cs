using CRT.Server.Handlers.Accounts;
using CRT.Server.Handlers.Submissions;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SubmissionRateLimitPolicy - how much one address may submit in a day (security
    // review, 2026-09-25). Contributing needs no account, so before this nothing at all limited
    // how often anybody could create a submission.
    // ###########################################################################################
    public sealed class SubmissionRateLimitPolicyTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 25, 12, 0, 0, TimeSpan.Zero);

        private static List<RecentSubmission> Recent(int count, long bytesEach = 0, TimeSpan? ago = null) =>
            Enumerable.Range(0, count)
                .Select(_ => new RecentSubmission(SubmissionRateLimitPolicyTests.Now - (ago ?? TimeSpan.FromHours(1)), bytesEach))
                .ToList();

        [Fact]
        public void An_address_with_room_left_is_allowed()
        {
            Assert.True(SubmissionRateLimitPolicy.Check(
                SubmissionRateLimitPolicyTests.Recent(SubmissionRateLimitPolicy.MaxSubmissionsPerAddress - 1),
                0,
                SubmissionRateLimitPolicyTests.Now).IsAllowed);
        }

        [Fact]
        public void The_count_limit_refuses_and_says_when_to_come_back()
        {
            RateLimitVerdict verdict = SubmissionRateLimitPolicy.Check(
                SubmissionRateLimitPolicyTests.Recent(SubmissionRateLimitPolicy.MaxSubmissionsPerAddress, ago: TimeSpan.FromHours(4)),
                0,
                SubmissionRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);

            // The oldest entry leaves the window twenty hours from now - that frees the room.
            Assert.Equal(TimeSpan.FromHours(20), verdict.RetryAfter);
        }

        [Fact]
        public void The_byte_budget_refuses_an_upload_that_would_take_it_over()
        {
            long almostAll = SubmissionRateLimitPolicy.MaxUploadBytesPerAddress - 100;

            List<RecentSubmission> recent = SubmissionRateLimitPolicyTests.Recent(1, almostAll);

            Assert.True(SubmissionRateLimitPolicy.Check(recent, 100, SubmissionRateLimitPolicyTests.Now).IsAllowed);
            Assert.False(SubmissionRateLimitPolicy.Check(recent, 101, SubmissionRateLimitPolicyTests.Now).IsAllowed);
        }

        // "Would this one take it over", not "is the bucket already full" - otherwise a single
        // upload larger than the whole budget would slip in because the bucket started empty.
        [Fact]
        public void A_single_upload_larger_than_the_whole_budget_is_refused()
        {
            Assert.False(SubmissionRateLimitPolicy.Check(
                [],
                SubmissionRateLimitPolicy.MaxUploadBytesPerAddress + 1,
                SubmissionRateLimitPolicyTests.Now).IsAllowed);
        }

        // Entries older than the window are ignored rather than trusting the caller to prune them.
        [Fact]
        public void Submissions_older_than_the_window_do_not_count()
        {
            Assert.True(SubmissionRateLimitPolicy.Check(
                SubmissionRateLimitPolicyTests.Recent(
                    SubmissionRateLimitPolicy.MaxSubmissionsPerAddress * 3,
                    SubmissionRateLimitPolicy.MaxUploadBytesPerAddress,
                    ago: SubmissionRateLimitPolicy.Window + TimeSpan.FromMinutes(1)),
                0,
                SubmissionRateLimitPolicyTests.Now).IsAllowed);
        }

        // A negative byte count (a corrupted row, a hostile value) must not buy extra budget.
        [Fact]
        public void A_negative_byte_count_does_not_buy_budget()
        {
            List<RecentSubmission> recent =
            [
                new(SubmissionRateLimitPolicyTests.Now - TimeSpan.FromHours(1), -SubmissionRateLimitPolicy.MaxUploadBytesPerAddress),
                new(SubmissionRateLimitPolicyTests.Now - TimeSpan.FromHours(1), SubmissionRateLimitPolicy.MaxUploadBytesPerAddress)
            ];

            Assert.False(SubmissionRateLimitPolicy.Check(recent, 1, SubmissionRateLimitPolicyTests.Now).IsAllowed);
        }
    }
}
