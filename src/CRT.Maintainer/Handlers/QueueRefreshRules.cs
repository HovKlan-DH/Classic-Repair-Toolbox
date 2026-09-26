using System;

namespace CRT.Maintainer.Handlers
{
    // ###########################################################################################
    // WHEN THE QUEUE CHECKS ITSELF (owner request, 2026-09-26: the "Refresh" button made obsolete -
    // "ideally it just updates if there are new stuff, but of course it should not conflict if the
    // maintainer is in the middle of something").
    //
    // Every PollInterval while the window is IN FRONT, and when the maintainer comes back to it
    // (unless it asked within ActivationGap). Not while it sits in the background: nobody is looking,
    // and each check extends the session - an application left open behind other windows would
    // otherwise keep a maintainer signed in for as long as it runs.
    //
    // What a check may change on screen is MaintainerMain.QueueRefresh.cs's: the list, never the
    // submission that is open.
    // ###########################################################################################
    public static class QueueRefreshRules
    {
        public static readonly TimeSpan PollInterval = TimeSpan.FromMinutes(1);

        public static readonly TimeSpan ActivationGap = TimeSpan.FromSeconds(15);

        // Whether a check is due: never asked, or asked at least `gap` ago.
        public static bool IsDue(DateTimeOffset? lastAskedUtc, DateTimeOffset now, TimeSpan gap) =>
            lastAskedUtc is null || now - lastAskedUtc.Value >= gap;
    }
}
