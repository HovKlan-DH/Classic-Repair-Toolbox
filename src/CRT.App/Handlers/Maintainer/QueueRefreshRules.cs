using System;
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
    // (unless it asked within ActivationGap). Not while it sits in the background: nobody is looking,
    // and each check extends the session - CRT left open behind other windows, or on another tab, would
    // otherwise keep a maintainer signed in for as long as it runs.
    //
    // What a check may change on screen is TabMaintainer.QueueRefresh.cs's: the list, never the
    // submission that is open.
    // ###########################################################################################
    public static class QueueRefreshRules
    {
        public static readonly TimeSpan PollInterval = TimeSpan.FromMinutes(1);

        public static readonly TimeSpan ActivationGap = TimeSpan.FromSeconds(15);

        // Whether a check is due: never asked, or asked at least `gap` ago.
        public static bool IsDue(DateTimeOffset? lastAskedUtc, DateTimeOffset now, TimeSpan gap) =>
            lastAskedUtc is null || now - lastAskedUtc.Value >= gap;

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
    }
}
