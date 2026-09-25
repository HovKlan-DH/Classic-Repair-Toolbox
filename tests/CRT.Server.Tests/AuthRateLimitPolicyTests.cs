using CRT.Server.Handlers.Accounts;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers the rule that decides when an authentication attempt is refused.
    //
    // This limiter is not only about password guessing: Argon2id allocates ~128 MiB per concurrent
    // verification at the configured cost, so unlimited attempts are a memory-exhaustion vector
    // whether or not any password is ever guessed. That is why it must be applied BEFORE the
    // hasher runs.
    //
    // The lockout-expiry tests matter as much as the refusal tests. A per-account lockout that
    // does not expire on its own is itself a denial of service - anyone who knows your address
    // could lock you out of your own account by failing to log in as you.
    // ###########################################################################################
    public class AuthRateLimitPolicyTests
    {
        private static readonly DateTimeOffset Now =
            new(2026, 9, 21, 12, 0, 0, TimeSpan.Zero);

        private static IReadOnlyList<DateTimeOffset> Failures(int count, TimeSpan ago)
        {
            return Enumerable.Range(0, count)
                .Select(_ => AuthRateLimitPolicyTests.Now - ago)
                .ToList();
        }

        private static readonly IReadOnlyList<DateTimeOffset> NoFailures = Array.Empty<DateTimeOffset>();

        // -----------------------------------------------------------------------------------
        // The per-account bucket.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void A_first_attempt_is_allowed()
        {
            Assert.True(AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.Now).IsAllowed);
        }

        [Fact]
        public void Attempts_below_the_limit_are_allowed()
        {
            // Four failures, limit is five. Someone who has genuinely forgotten their password
            // gets several tries before anything happens.
            Assert.True(AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.Failures(4, TimeSpan.FromMinutes(1)),
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.Now).IsAllowed);
        }

        [Fact]
        public void Reaching_the_limit_refuses_the_next_attempt()
        {
            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.Failures(5, TimeSpan.FromMinutes(1)),
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);
            Assert.True(verdict.RetryAfter > TimeSpan.Zero);
        }

        [Fact]
        public void Failures_older_than_the_window_do_not_count()
        {
            // The lockout expires on its own. Without this, a handful of failures months apart
            // would eventually lock an account permanently.
            Assert.True(AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.Failures(20, TimeSpan.FromHours(2)),
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.Now).IsAllowed);
        }

        [Fact]
        public void The_retry_delay_is_measured_from_the_OLDEST_attempt_in_the_window()
        {
            // The accidental-permanent-lockout bug. Measuring from the NEWEST attempt holds
            // someone out for a full window after their last try, and every further try extends
            // it - so a person tapping "try again" can never get back in. Measuring from the
            // oldest is correct: that is the attempt whose expiry frees a slot.
            var attempts = new[]
            {
                AuthRateLimitPolicyTests.Now - TimeSpan.FromMinutes(14),  // oldest, expires in 1 min
                AuthRateLimitPolicyTests.Now - TimeSpan.FromMinutes(3),
                AuthRateLimitPolicyTests.Now - TimeSpan.FromMinutes(2),
                AuthRateLimitPolicyTests.Now - TimeSpan.FromMinutes(1),
                AuthRateLimitPolicyTests.Now
            };

            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(
                attempts, AuthRateLimitPolicyTests.NoFailures, AuthRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);

            // About one minute, not about fifteen.
            Assert.True(
                verdict.RetryAfter <= TimeSpan.FromMinutes(1.1),
                $"Expected roughly a minute, got {verdict.RetryAfter}.");
        }

        [Fact]
        public void An_attempt_recorded_in_the_future_does_not_produce_a_negative_delay()
        {
            // A clock adjustment, or a row written fractionally ahead. A negative Retry-After
            // header is nonsense and some clients treat it as "retry immediately", others as an
            // error.
            var attempts = Enumerable.Range(0, 5)
                .Select(_ => AuthRateLimitPolicyTests.Now + TimeSpan.FromMinutes(5))
                .ToList();

            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(
                attempts, AuthRateLimitPolicyTests.NoFailures, AuthRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);
            Assert.True(verdict.RetryAfter >= TimeSpan.Zero);
            Assert.True(verdict.RetryAfter <= AuthRateLimitPolicy.AccountWindow);
        }

        // -----------------------------------------------------------------------------------
        // The per-address bucket. Independent of the per-account one, and needed for a case the
        // per-account limit is blind to.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void One_address_spraying_many_accounts_is_refused_by_the_address_bucket()
        {
            // Password spraying: one common password against many different accounts. Each
            // individual account sees ONE failure, so no per-account limit could ever see it.
            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.Failures(1, TimeSpan.FromSeconds(30)),
                AuthRateLimitPolicyTests.Failures(30, TimeSpan.FromMinutes(5)),
                AuthRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);
        }

        [Fact]
        public void The_address_limit_is_higher_than_the_account_limit()
        {
            // One address legitimately covers a household, a workplace, or a whole country behind
            // CGNAT. It is a backstop against spraying, not a per-person limit - so it must not
            // trip at the same threshold as a single account's.
            Assert.True(AuthRateLimitPolicy.MaxFailuresPerAddress > AuthRateLimitPolicy.MaxFailuresPerAccount);
        }

        [Fact]
        public void A_busy_but_legitimate_address_is_not_refused()
        {
            // Ten failures from one office over fifteen minutes is plausible and must not lock
            // everybody there out.
            Assert.True(AuthRateLimitPolicy.CheckLogin(
                AuthRateLimitPolicyTests.NoFailures,
                AuthRateLimitPolicyTests.Failures(10, TimeSpan.FromMinutes(5)),
                AuthRateLimitPolicyTests.Now).IsAllowed);
        }

        // -----------------------------------------------------------------------------------
        // Mail-sending requests.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Mail_requests_are_limited_more_tightly_than_logins()
        {
            // Registration and reset both SEND MAIL, so an unlimited rate is a way to use this
            // service to flood somebody's inbox. Here a SUCCESS is the thing being abused, which
            // is why these are counted regardless of outcome.
            Assert.True(AuthRateLimitPolicy.MaxMailRequestsPerAddress < AuthRateLimitPolicy.MaxFailuresPerAddress);
        }

        [Fact]
        public void Reaching_the_mail_request_limit_refuses_the_next_one()
        {
            RateLimitVerdict verdict = AuthRateLimitPolicy.CheckMailRequest(
                AuthRateLimitPolicyTests.Failures(10, TimeSpan.FromMinutes(10)),
                AuthRateLimitPolicyTests.Now);

            Assert.False(verdict.IsAllowed);
        }

        [Fact]
        public void Mail_requests_outside_the_window_do_not_count()
        {
            Assert.True(AuthRateLimitPolicy.CheckMailRequest(
                AuthRateLimitPolicyTests.Failures(50, TimeSpan.FromHours(3)),
                AuthRateLimitPolicyTests.Now).IsAllowed);
        }

        [Fact]
        public void The_mail_window_is_longer_than_the_login_window()
        {
            // Mail abuse is measured over a longer span: ten mails in an hour is the concern,
            // not ten in fifteen minutes.
            Assert.True(AuthRateLimitPolicy.MailWindow > AuthRateLimitPolicy.AccountWindow);
        }
    }
}
