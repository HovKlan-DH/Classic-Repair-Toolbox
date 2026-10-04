using System;
using System.Collections.Generic;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // WHEN THE QUEUE CHECKS ITSELF (owner request, 2026-09-26: the "Refresh" button made obsolete -
    // "ideally it just updates if there are new stuff, but of course it should not conflict if the
    // maintainer is in the middle of something").
    //
    // Every PollInterval while the Maintainer tab is on screen with CRT's window IN FRONT, and when
    // the maintainer comes back to it
    // (unless it asked within ActivationGap). Off screen, since 2026-09-30, a lighter check keeps the
    // tab's own badge current instead - see MinuteCheck, which also says what that costs.
    //
    // What a check may change on screen is TabMaintainer.QueueRefresh.cs's: the list, never the
    // submission that is open.
    // ###########################################################################################
    // What one minute check asks the server for - see QueueRefreshRules.MinuteCheck.
    public enum QueueCheck
    {
        Nothing,
        Everything,
        BadgesOnly
    }

    public static class QueueRefreshRules
    {
        public static readonly TimeSpan PollInterval = TimeSpan.FromMinutes(1);

        public static readonly TimeSpan ActivationGap = TimeSpan.FromSeconds(15);

        // Whether a check is due: never asked, or asked at least `gap` ago.
        public static bool IsDue(DateTimeOffset? lastAskedUtc, DateTimeOffset now, TimeSpan gap) =>
            lastAskedUtc is null || now - lastAskedUtc.Value >= gap;

        // ###########################################################################################
        // *** WHAT THE MINUTE CHECK ASKS FOR (owner request, 2026-09-30). *** The Maintainer tab's
        // badge in CRT's row of tabs counts what waits for this account, and has to stay right while
        // the maintainer works on another tab - "It should check from server once every minute".
        //
        //   - the tab ON SCREEN with CRT's window in front: EVERYTHING, as before - the queue, the
        //     BETA list and whatever the screen on show reads (ReadsSystemsOverview);
        //   - otherwise, while the badge can be SEEN (the tab turned on, CRT's window not
        //     minimised): only the two lists the badge counts are READ - the queue and the BETA
        //     list. Not the drop-down listing, and never the Systems overview, which walks both data
        //     trees. (Reading the queue is the same background read as on screen, so what it does
        //     with a changed row is unchanged: it re-reads that submission's detail - touching no
        //     table - and an expired session lands on the sign-in screen, which is what empties the
        //     badge. Code review, 2026-10-01: the earlier wording read as if nothing else happened.)
        //   - otherwise NOTHING: a badge nobody can see is not worth a request a minute.
        //
        // *** THE PRICE, STATED: A MAINTAINER NOW STAYS SIGNED IN WHILE CRT IS USED AT ALL. *** Every
        // request slides the session's 30-day expiry forward (SessionExtensionRules), and a check
        // that only ran while the tab was on screen let a session lapse after 30 days of not
        // reviewing. With the badge, 30 days of not RUNNING CRT - the owner asked for a badge that
        // is current, and a badge that is current has to ask.
        // ###########################################################################################
        public static QueueCheck MinuteCheck(bool onScreenAndInFront, bool badgeCanBeSeen) =>
            onScreenAndInFront
                ? QueueCheck.Everything
                : badgeCanBeSeen ? QueueCheck.BadgesOnly : QueueCheck.Nothing;

        // ###########################################################################################
        // *** WHAT THE MINUTE CHECK READS BESIDE THE QUEUE (code review, 2026-09-27). *** It read
        // every list whatever screen was shown - and the Systems overview is the costly one: the
        // server walks both data trees and counts a month of board views for it, every minute, for
        // every open maintainer window. The BETA list and the drop-down listing are still read every
        // time, because the Approve gate and the mode badges read them on any screen; the overview
        // only while the Systems screen is shown, or before it has been read at all (its badge
        // counts the systems). A check that is not the minute check - signing in, after a decision,
        // pressing a screen's button - reads everything, as before.
        // ###########################################################################################
        public static bool ReadsSystemsOverview(bool minuteCheck, MaintainerMode shown, bool overviewKnown) =>
            !minuteCheck || !overviewKnown || shown == MaintainerMode.Systems;

        // ###########################################################################################
        // Whether a system's row changed enough to read its panel again. The view count moves with
        // every board anybody opens in CRT, so comparing it too re-read the whole panel - accounts,
        // audit trail and a year of view facts - nearly every minute while nothing else had changed.
        // The panel's own numbers are read again the next time something else changes, or the
        // system is chosen again.
        // ###########################################################################################
        public static bool SystemChanged(SystemOverviewEntry before, SystemOverviewEntry after)
        {
            ArgumentNullException.ThrowIfNull(before);
            ArgumentNullException.ThrowIfNull(after);

            return !Equals(before with { ViewsLast30Days = null }, after with { ViewsLast30Days = null });
        }

        // ###########################################################################################
        // *** WHETHER BETA'S BOARD MOVED BETWEEN TWO READINGS OF A SYSTEM (owner report, 2026-10-04). ***
        // The Systems screen read a system's table and files once and kept them until another system
        // was chosen - so a fix approved, and even promoted, went on showing the warning it fixed
        // ("it continues to show the error ... After an app restart ... it did not show as an error
        // any more"). `readAt` is the system as it was when the table or files were read, `now` as it
        // is; true means they are out of date.
        //
        // The server's record of BETA's content when either reading carries it - every publish to
        // BETA and every push-back moves it. From a server older than that, the facts that move with
        // both: the BETA revision, and whether the system waits for production.
        // ###########################################################################################
        public static bool BetaBoardChanged(SystemOverviewEntry readAt, SystemOverviewEntry now)
        {
            ArgumentNullException.ThrowIfNull(readAt);
            ArgumentNullException.ThrowIfNull(now);

            if (readAt.BetaContentHash is not null || now.BetaContentHash is not null)
                return !string.Equals(readAt.BetaContentHash, now.BetaContentHash, StringComparison.Ordinal);

            return readAt.InBeta != now.InBeta ||
                   !string.Equals(readAt.BetaRevision, now.BetaRevision, StringComparison.Ordinal) ||
                   readAt.IsAwaitingProduction != now.IsAwaitingProduction;
        }

        // ###########################################################################################
        // Whether the STABLE source's board moved between two readings of a system (2026-10-04, the
        // Board data and Files views' Stable half) - only a promotion moves it, and a promotion sets
        // the system's stable revision and when it was published there. True means the stable
        // table or files read at `readAt` are out of date.
        // ###########################################################################################
        public static bool StableBoardChanged(SystemOverviewEntry readAt, SystemOverviewEntry now)
        {
            ArgumentNullException.ThrowIfNull(readAt);
            ArgumentNullException.ThrowIfNull(now);

            return readAt.InProduction != now.InProduction ||
                   !string.Equals(readAt.ProductionRevision, now.ProductionRevision, StringComparison.Ordinal) ||
                   readAt.ProductionPublishedUtc != now.ProductionPublishedUtc;
        }

        // ###########################################################################################
        // Whether a system's TABLE is out of date: its board moved (BetaBoardChanged), or whether this
        // account may change it did - the server's MayEdit, which is off while the system waits under
        // BETA > Stable and for a system closed to contributions, and follows who maintains it. A
        // promotion moves only the first of those, and left the table read-only with "it waits
        // under BETA > Stable" after it no longer did. The Files view follows the board alone.
        //
        // *** WHO maintains it, not how many (code review, 2026-10-04). *** One maintainer swapped
        // for another keeps the count, and the account swapped in or out had a table that said the
        // opposite of what the server would do. With the system's maintainers at both readings (its
        // detail's account ids), any difference counts; without them, the count stands in.
        // ###########################################################################################
        public static bool SystemTableChanged(
            SystemOverviewEntry readAt,
            SystemOverviewEntry now,
            IReadOnlyCollection<long>? maintainersReadAt = null,
            IReadOnlyCollection<long>? maintainersNow = null)
        {
            ArgumentNullException.ThrowIfNull(readAt);
            ArgumentNullException.ThrowIfNull(now);

            bool maintainersMoved = maintainersReadAt is not null && maintainersNow is not null
                ? !maintainersReadAt.ToHashSet().SetEquals(maintainersNow)
                : readAt.MaintainerCount != now.MaintainerCount;

            return QueueRefreshRules.BetaBoardChanged(readAt, now) ||
                   readAt.IsAwaitingProduction != now.IsAwaitingProduction ||
                   readAt.IsAccepting != now.IsAccepting ||
                   maintainersMoved;
        }
    }
}
