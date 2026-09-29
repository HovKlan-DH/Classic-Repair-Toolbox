using System;

namespace CRT.Server.Handlers.Accounts
{
    // ###########################################################################################
    // WHEN A SESSION'S EXPIRY SHOULD BE PUSHED FORWARD - sliding expiry, so a maintainer signs in
    // once rather than on a schedule (owner request, 2026-09-22).
    //
    // *** THIS IS NOT REFRESH, AND THE DIFFERENCE IS THE WHOLE POINT. *** AccountFlows.RefreshAsync
    // issues a NEW token and marks the old one `replaced_by_id`, which arms reuse detection: the
    // old value presented again revokes EVERY session for that account. That is correct for a
    // browser, where a rotated token narrows the window a stolen one is useful for.
    //
    // It is wrong for a desktop app holding its token in a file. Rotation puts a write-to-disk
    // between "the server has already poisoned the old token" and "we have safely stored the new
    // one", and a crash, a power cut or a failed write inside that window locks the account out of
    // every session and writes a `session.reuse_detected` audit entry that reads as an attack.
    // The user then has to reset a password because their laptop lost power.
    //
    // Extension has no such window because THE TOKEN NEVER CHANGES. It is one UPDATE to
    // `expires_utc` on the row already being used. Nothing is issued, nothing is poisoned, and a
    // failure to extend costs nothing at all - the session simply keeps its old expiry and the
    // next request tries again.
    //
    // *** A SLIDING WINDOW IS SHORTER THAN A LONG FIXED ONE, NOT LONGER. *** The alternative
    // considered was simply raising RefreshTokenDays to 365. That keeps a stolen token alive for
    // a year even if the thief never touches it. A sliding 30-day window dies 30 days after LAST
    // USE, so an abandoned or stolen machine goes cold on its own while the person actually using
    // the app every week never signs in again.
    // ###########################################################################################
    public static class SessionExtensionRules
    {
        // ###########################################################################################
        // Only extend once the session is meaningfully used up.
        //
        // *** WITHOUT THIS THRESHOLD, EVERY REQUEST WRITES A ROW. *** The Maintainer tab calls the
        // queue on launch, on every refresh click and after every decision, so extending on each
        // one turns a read-only screen into a steady stream of UPDATEs against the sessions table
        // for no benefit - the expiry would move by seconds.
        //
        // Half the lifetime is the natural point: the write happens at most once per half-window
        // per session, and the expiry can never be more than half a window stale.
        // ###########################################################################################
        public static bool ShouldExtend(
            DateTimeOffset expiresUtc,
            DateTimeOffset now,
            TimeSpan lifetime)
        {
            // A lifetime that is zero or negative is a misconfiguration; extending against it
            // would move the expiry BACKWARDS. Refusing leaves the session exactly as it is.
            if (lifetime <= TimeSpan.Zero)
                return false;

            // Already expired. Extending here would resurrect a dead session, which is precisely
            // what an attacker holding a stale token would want. Expiry is final; only a fresh
            // login issues a new one.
            if (expiresUtc <= now)
                return false;

            // *** THE SUBTRACTION IS ON THE DIFFERENCE, NOT ON A DateTimeOffset. *** The same
            // overflow trap ReviewSession.IsUsableAt documents: `expiresUtc - halfLifetime` throws
            // ArgumentOutOfRangeException when expiresUtc is near DateTimeOffset.MinValue, and a
            // method on the path of every authenticated request must be total.
            TimeSpan remaining = expiresUtc - now;

            return remaining < SessionExtensionRules.Halve(lifetime);
        }

        // ###########################################################################################
        // The new expiry: a full lifetime from NOW, never from the old expiry.
        //
        // Measuring from the old expiry would compound - a session used often would drift further
        // and further into the future until it outlived any policy anybody intended. From `now`,
        // the guarantee is exact and easy to state: a session dies one lifetime after it was last
        // used, whatever happened before that.
        // ###########################################################################################
        public static DateTimeOffset NewExpiry(DateTimeOffset now, TimeSpan lifetime)
        {
            if (lifetime <= TimeSpan.Zero)
                throw new ArgumentOutOfRangeException(nameof(lifetime), "A session lifetime must be positive.");

            // Saturates rather than throwing. A lifetime large enough to overflow is nonsensical
            // configuration, but this sits on the authenticated request path where an exception is
            // a 500 for a user who did nothing wrong.
            if (lifetime > DateTimeOffset.MaxValue - now)
                return DateTimeOffset.MaxValue;

            return now + lifetime;
        }

        // Halving a TimeSpan via its ticks, which cannot overflow downward.
        private static TimeSpan Halve(TimeSpan value) => TimeSpan.FromTicks(value.Ticks / 2);
    }
}
