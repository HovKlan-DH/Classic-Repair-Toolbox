namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // How many authentication attempts an address or an IP may make, and how long a refusal
    // lasts. Pure: it is handed the attempts already recorded and the current time, and returns a
    // verdict. It holds no state, reads no clock and touches no store, so every rule is a test.
    //
    // WHY THIS EXISTS AT ALL, beyond the obvious password guessing: Argon2id allocates ~128 MiB
    // per concurrent verification at the configured cost. Unlimited login attempts are therefore
    // a MEMORY EXHAUSTION vector on a small box regardless of whether any password is ever
    // guessed - so the limiter must be applied BEFORE the hasher runs, never after. That ordering
    // is the whole point and is easy to get backwards.
    //
    // THREE INDEPENDENT BUCKETS, and all are needed:
    //
    //   Per ACCOUNT AND ADDRESS (the normalised email, from one IP) - stops someone working through
    //     a password list against one known address. The tight one: five in fifteen minutes.
    //   Per ACCOUNT, from anywhere - the same list from a rotating pool of addresses. A clearly
    //     HIGHER bar (code review, 2026-10-13): it used to be the tight one, so five wrong passwords
    //     from ANYBODY locked the account for everybody for fifteen minutes - and the
    //     administrator's address is printed at the foot of every mail the service sends. Now a
    //     stranger's guesses lock out only the stranger's own address, and it takes many addresses
    //     guessing at once to slow the real owner down.
    //   Per IP - stops someone spraying one common password across many accounts, which no
    //     per-account limit can see, since each individual account gets only one attempt.
    //
    // FAILURES COUNT, SUCCESSES CLEAR. A person who gets it right on the fourth try is not an
    // attacker and must not be locked out afterwards.
    //
    // *** THE LOCKOUT IS A DELAY, NOT A DISABLE. *** A per-account lockout that persists is
    // itself a denial of service: anyone who knows your address can lock you out of your own
    // account indefinitely by failing to log in as you. So the window is short, it expires on its
    // own, and no administrator action is needed to clear it.
    // ###########################################################################################
    public static class AuthRateLimitPolicy
    {
        // Five failures for one account from one address before a refusal. Generous enough that a
        // genuinely forgotten password with three or four guesses does not trip it, small enough
        // that an online guessing attack from one address gets nowhere - at five per fifteen
        // minutes, a thousand-password list takes over two days.
        public const int MaxFailuresPerAccountAndAddress = 5;

        // ###########################################################################################
        // The account's own bar, whatever the addresses (code review, 2026-10-13) - five times the
        // one above, so it takes at least five addresses guessing at once to reach it, and one
        // stranger never does. At twenty-five per fifteen minutes a thousand-password list still
        // takes ten hours, however many addresses carry it. Below the per-address bar, which covers
        // a whole household or CGNAT range.
        // ###########################################################################################
        public const int MaxFailuresPerAccount = 25;

        // The window over which those failures are counted, and how long a refusal lasts.
        public static readonly TimeSpan AccountWindow = TimeSpan.FromMinutes(15);

        // Higher, because one address legitimately covers a household, a workplace or a whole
        // country behind CGNAT. It is a backstop against spraying, not a per-person limit.
        public const int MaxFailuresPerAddress = 30;

        public static readonly TimeSpan AddressWindow = TimeSpan.FromMinutes(15);

        // Registration and password-reset requests both send mail, so an unlimited rate is a way
        // to use this service to flood somebody's inbox. Lower, and counted per IP regardless of
        // success, because here a SUCCESS is the thing being abused.
        public const int MaxMailRequestsPerAddress = 10;

        public static readonly TimeSpan MailWindow = TimeSpan.FromHours(1);

        // ###########################################################################################
        // Decides whether a login attempt may proceed.
        //
        // recentAccountAndAddressFailures / recentAccountFailures / recentAddressFailures: the times
        // of failures already recorded inside the relevant window - for this account from this
        // address, for this account from anywhere, and from this address for any account - oldest
        // or newest order irrelevant.
        //
        // The verdict carries a RetryAfter so the endpoint can answer 429 with a meaningful
        // header rather than leaving the caller to guess.
        // ###########################################################################################
        public static RateLimitVerdict CheckLogin(
            IReadOnlyList<DateTimeOffset> recentAccountAndAddressFailures,
            IReadOnlyList<DateTimeOffset> recentAccountFailures,
            IReadOnlyList<DateTimeOffset> recentAddressFailures,
            DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(recentAccountAndAddressFailures);
            ArgumentNullException.ThrowIfNull(recentAccountFailures);
            ArgumentNullException.ThrowIfNull(recentAddressFailures);

            RateLimitVerdict pairVerdict = AuthRateLimitPolicy.CheckBucket(
                recentAccountAndAddressFailures,
                AuthRateLimitPolicy.MaxFailuresPerAccountAndAddress,
                AuthRateLimitPolicy.AccountWindow,
                now);

            if (!pairVerdict.IsAllowed)
                return pairVerdict;

            RateLimitVerdict accountVerdict = AuthRateLimitPolicy.CheckBucket(
                recentAccountFailures,
                AuthRateLimitPolicy.MaxFailuresPerAccount,
                AuthRateLimitPolicy.AccountWindow,
                now);

            if (!accountVerdict.IsAllowed)
                return accountVerdict;

            return AuthRateLimitPolicy.CheckBucket(
                recentAddressFailures,
                AuthRateLimitPolicy.MaxFailuresPerAddress,
                AuthRateLimitPolicy.AddressWindow,
                now);
        }

        // ###########################################################################################
        // Decides whether a mail-sending request (register, resend verification, forgot password)
        // may proceed. Counted per IP and counting SUCCESSES too - see MaxMailRequestsPerAddress.
        // ###########################################################################################
        public static RateLimitVerdict CheckMailRequest(
            IReadOnlyList<DateTimeOffset> recentRequests,
            DateTimeOffset now)
        {
            ArgumentNullException.ThrowIfNull(recentRequests);

            return AuthRateLimitPolicy.CheckBucket(
                recentRequests,
                AuthRateLimitPolicy.MaxMailRequestsPerAddress,
                AuthRateLimitPolicy.MailWindow,
                now);
        }

        // ###########################################################################################
        // The shared rule. Entries outside the window are ignored rather than requiring the caller
        // to have pruned them - a caller that over-supplies must not be able to cause a false
        // refusal.
        //
        // RetryAfter is measured from the OLDEST attempt still inside the window, because that is
        // the one whose expiry frees a slot. Measuring from the newest would hold someone out for
        // a full window after their last try, which turns a burst of five into fifteen minutes of
        // lockout that extends every time they try again - the accidental permanent lockout.
        // ###########################################################################################
        private static RateLimitVerdict CheckBucket(
            IReadOnlyList<DateTimeOffset> attempts,
            int maximum,
            TimeSpan window,
            DateTimeOffset now)
        {
            DateTimeOffset cutoff = now - window;

            var inWindow = attempts.Where(attempt => attempt > cutoff).ToList();

            if (inWindow.Count < maximum)
                return RateLimitVerdict.Allowed();

            DateTimeOffset oldest = inWindow.Min();
            TimeSpan retryAfter = (oldest + window) - now;

            // Clamp: a clock adjustment or an attempt recorded fractionally in the future would
            // otherwise produce a negative or absurd delay in a Retry-After header.
            if (retryAfter < TimeSpan.Zero)
                retryAfter = TimeSpan.Zero;

            if (retryAfter > window)
                retryAfter = window;

            return RateLimitVerdict.Refused(retryAfter);
        }
    }

    // ###########################################################################################
    // RetryAfter is meaningful only when IsAllowed is false.
    // ###########################################################################################
    public readonly record struct RateLimitVerdict(bool IsAllowed, TimeSpan RetryAfter)
    {
        public static RateLimitVerdict Allowed() => new(true, TimeSpan.Zero);

        public static RateLimitVerdict Refused(TimeSpan retryAfter) => new(false, retryAfter);
    }
}
