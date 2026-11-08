using System;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // HOW LONG AUTOMATIC CONNECT WAITS BEFORE ITS NEXT TRY (owner report, 2026-11-07: an Identify command the scope did not
    // answer was sent again on a new connection every few seconds for minutes - "it seems to repeat forever(?) this
    // command" - after which the scope answered nothing, a correct "*IDN?" included). A second after a session ended or a
    // try went through, then twice as long after each try that failed in a row, never more than half a minute - a scope
    // switched on again is still found within it. The first failure of a run is said once, with why (FailedLine).
    // ###########################################################################################
    public static class ScopeAutoConnectBackoff
    {
        public static readonly TimeSpan Longest = TimeSpan.FromSeconds(30);

        // The wait after `failures` tries failed in a row - 0 none: 1 s, 2 s, 4 s ... 30 s at most.
        public static TimeSpan DelayAfter(int failures)
        {
            if (failures <= 0)
            {
                return TimeSpan.FromSeconds(1);
            }

            double seconds = Math.Pow(2, Math.Min(failures, 16));

            return seconds >= Longest.TotalSeconds ? Longest : TimeSpan.FromSeconds(seconds);
        }

        // ###########################################################################################
        // What one turn of the automatic connect loop came to: a try that failed or went through - or no try at all,
        // because a session is established, automatic connect is held or off, or there is no scope to try.
        // ###########################################################################################
        public enum Turn
        {
            Failed,
            Connected,
            NoTry,
        }

        // ###########################################################################################
        // The tries failed in a row once a turn is over: one more after a try that failed, none after anything else
        // (code review, 2026-11-08). Counted down only by a try that went through, a run outlived a session the user
        // connected by hand, a hold and a scope switched off - so the first try after a session ended waited half a
        // minute instead of a second, and said nothing of why it failed.
        // ###########################################################################################
        public static int FailuresAfter(int failures, Turn turn) => turn == Turn.Failed ? Math.Max(failures, 0) + 1 : 0;

        // The one line a run of failed tries writes in the tab's output.
        public static string FailedLine(string reason) =>
            $"Automatic connect failed - {reason}. CRT keeps trying, waiting longer between tries while it fails (at most {Longest.TotalSeconds:0} seconds).";
    }
}
