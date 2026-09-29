using CRT.Server.Handlers.Accounts;
using Xunit;

namespace CRT.Server.Tests
{
    // ###########################################################################################
    // Covers SessionExtensionRules - WHEN a session's expiry moves and WHAT it moves to
    // (sliding expiry, owner request 2026-09-22).
    //
    // The behaviour being pinned is a promise to a person: sign in once and keep working. The
    // failure modes on either side are both real - extending too eagerly writes a database row on
    // every request of a read-only screen, and never extending asks for a password that was
    // promised not to be needed - so both edges are asserted rather than just the happy path.
    // ###########################################################################################
    public sealed class SessionExtensionRulesTests
    {
        private static readonly DateTimeOffset Now = new(2026, 9, 22, 12, 0, 0, TimeSpan.Zero);

        private static readonly TimeSpan Lifetime = TimeSpan.FromDays(30);

        [Fact]
        public void A_FRESH_session_is_not_extended()
        {
            // Just issued, so nearly a full lifetime remains. Extending here would move the expiry
            // by seconds and write a row for nothing - and the Maintainer tab calls the queue on
            // launch, on every refresh click and after every decision.
            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(30),
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void A_session_PAST_HALFWAY_is_extended()
        {
            // 14 days left of 30 - under half, so the window slides.
            Assert.True(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(14),
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void EXACTLY_half_remaining_does_NOT_extend()
        {
            // The boundary is strictly "less than half". Either answer would be defensible, but it
            // has to be pinned: a "<=" here would fire an extra write for every session that
            // happens to land exactly on the boundary.
            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(15),
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void An_ALREADY_EXPIRED_session_is_NEVER_extended()
        {
            // *** THE SECURITY-CRITICAL CASE. *** Without this, a stale token recovered from an old
            // backup or a stolen disk could be extended back to life indefinitely, and a session
            // would never actually die. Expiry is final; only a fresh login issues a new one.
            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddSeconds(-1),
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));

            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(-400),
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void A_session_expiring_at_EXACTLY_now_is_not_extended()
        {
            // The boundary of the case above. "<=" rather than "<", because a session whose expiry
            // is this instant is spent.
            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void A_NONSENSE_lifetime_extends_nothing_rather_than_moving_expiry_BACKWARDS()
        {
            // RefreshTokenDays is validated at startup, so this is defence against a future caller
            // computing a lifetime rather than reading one. A zero or negative lifetime would
            // otherwise set an expiry at or before `now` - instantly killing the session it was
            // asked to prolong.
            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(1),
                SessionExtensionRulesTests.Now,
                TimeSpan.Zero));

            Assert.False(SessionExtensionRules.ShouldExtend(
                SessionExtensionRulesTests.Now.AddDays(1),
                SessionExtensionRulesTests.Now,
                TimeSpan.FromDays(-5)));
        }

        [Fact]
        public void An_expiry_near_DateTimeOffset_MinValue_does_not_THROW()
        {
            // *** TOTALITY, and it is not hypothetical. *** ReviewSession.IsUsableAt documents two
            // consecutive overflow bugs of exactly this shape, both ArgumentOutOfRangeException
            // from subtracting a TimeSpan off a DateTimeOffset near its bounds. This method sits on
            // the path of every authenticated request, so a throw here is a 500 for a caller who
            // did nothing wrong.
            Assert.False(SessionExtensionRules.ShouldExtend(
                DateTimeOffset.MinValue,
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void An_expiry_near_DateTimeOffset_MaxValue_does_not_THROW()
        {
            // The other end, which is where the FIRST attempted fix for that bug moved the
            // overflow to rather than removing it.
            Assert.False(SessionExtensionRules.ShouldExtend(
                DateTimeOffset.MaxValue,
                SessionExtensionRulesTests.Now,
                SessionExtensionRulesTests.Lifetime));
        }

        // -----------------------------------------------------------------------------------
        // What the new expiry actually is.
        // -----------------------------------------------------------------------------------

        [Fact]
        public void The_new_expiry_is_a_full_lifetime_from_NOW()
        {
            Assert.Equal(
                SessionExtensionRulesTests.Now.AddDays(30),
                SessionExtensionRules.NewExpiry(
                    SessionExtensionRulesTests.Now, SessionExtensionRulesTests.Lifetime));
        }

        [Fact]
        public void The_new_expiry_is_measured_from_NOW_rather_than_from_the_OLD_expiry()
        {
            // *** THE ANTI-COMPOUNDING RULE. *** Measuring from the old expiry would add a lifetime
            // each time, so a session used often would drift years into the future and outlive any
            // policy anybody intended. From `now`, the guarantee is exact and sayable: a session
            // dies one lifetime after it was last used, whatever happened before that.
            //
            // Asserted as an INVARIANT over repeated extension, since a single call cannot tell
            // the two rules apart.
            DateTimeOffset expiry = SessionExtensionRulesTests.Now.AddDays(30);
            DateTimeOffset clock = SessionExtensionRulesTests.Now;

            for (int pass = 0; pass < 5; pass++)
            {
                clock = clock.AddDays(20);

                if (SessionExtensionRules.ShouldExtend(expiry, clock, SessionExtensionRulesTests.Lifetime))
                    expiry = SessionExtensionRules.NewExpiry(clock, SessionExtensionRulesTests.Lifetime);
            }

            // Exactly one lifetime past the last extension - never five.
            Assert.Equal(clock.AddDays(30), expiry);
        }

        [Fact]
        public void A_lifetime_that_would_OVERFLOW_saturates_rather_than_throwing()
        {
            // Nonsensical configuration, but this is on the authenticated request path where an
            // exception is a 500 for a user who did nothing wrong.
            Assert.Equal(
                DateTimeOffset.MaxValue,
                SessionExtensionRules.NewExpiry(
                    SessionExtensionRulesTests.Now, TimeSpan.MaxValue));
        }

        [Fact]
        public void NewExpiry_REFUSES_a_non_positive_lifetime()
        {
            // Throws rather than saturating, unlike the overflow case above, because there is no
            // sensible answer: every candidate value is at or before `now` and would kill the
            // session. ShouldExtend already refuses these, so reaching here is a caller bug worth
            // surfacing loudly rather than absorbing.
            Assert.Throws<ArgumentOutOfRangeException>(() =>
                SessionExtensionRules.NewExpiry(SessionExtensionRulesTests.Now, TimeSpan.Zero));
        }
    }
}
