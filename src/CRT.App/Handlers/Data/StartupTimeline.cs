using System;
using System.Diagnostics;
using System.Globalization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Records how long each phase of application startup took.
    //
    // The point of this class is the FIRST milestone. The log file only begins at
    // "Logger.Initialize()", which already sits inside "App.OnFrameworkInitializationCompleted" -
    // so everything before it (Velopack, the Avalonia AppBuilder, platform detection, Skia, and
    // parsing App.axaml) was previously invisible in the logs, even though it is the stretch where
    // the user is staring at nothing. Measuring from the OS-reported process start time is the only
    // way to see it.
    //
    // The type takes both the process start time and each "now" as arguments rather than reading
    // the clock itself, so the whole thing is testable without a running process.
    // ###########################################################################################
    public sealed class StartupTimeline
    {
        private readonly DateTime _processStart;
        private TimeSpan _previousMilestone = TimeSpan.Zero;

        public StartupTimeline(DateTime processStart)
        {
            this._processStart = processStart;
        }

        // ###########################################################################################
        // Returns the OS-reported start time of this process, or null when the platform refuses to
        // report it. Callers should say so in the log rather than quietly measuring from a fake
        // origin - an under-reported startup time is worse than an absent one.
        // ###########################################################################################
        public static DateTime? TryResolveProcessStartTime()
        {
            try
            {
                using var process = Process.GetCurrentProcess();
                return process.StartTime;
            }
            catch
            {
                return null;
            }
        }

        // ###########################################################################################
        // Records a named milestone and returns the log line describing it, reporting both the total
        // elapsed time since process start and the time spent in the phase since the last milestone.
        // ###########################################################################################
        public string Record(string milestone, DateTime now)
        {
            var sinceStart = Clamp(now - this._processStart);
            var sincePrevious = Clamp(sinceStart - this._previousMilestone);

            this._previousMilestone = sinceStart;

            var name = string.IsNullOrWhiteSpace(milestone) ? "(unnamed)" : milestone.Trim();

            return $"Startup milestone [{name}] [total {FormatDuration(sinceStart)}] [phase {FormatDuration(sincePrevious)}]";
        }

        // ###########################################################################################
        // Formats a duration for the log: whole milliseconds below one second, otherwise seconds to
        // two decimals. Always invariant - a Danish or German locale would otherwise write "3,14 s"
        // and make logs from different users inconsistent to read and to grep.
        // ###########################################################################################
        public static string FormatDuration(TimeSpan duration)
        {
            var clamped = Clamp(duration);

            // Round to whole milliseconds BEFORE choosing the unit, so 999.6 ms reports as "1.00 s"
            // rather than the nonsensical "1000 ms".
            var milliseconds = Math.Round(clamped.TotalMilliseconds);

            return milliseconds < 1000
                ? string.Format(CultureInfo.InvariantCulture, "{0:0} ms", milliseconds)
                : string.Format(CultureInfo.InvariantCulture, "{0:0.00} s", clamped.TotalSeconds);
        }

        // ###########################################################################################
        // Clamps a duration to zero. The system clock can move backwards mid-startup (NTP correction,
        // a user changing the time, a VM resuming), and a negative startup time in the log is noise
        // that looks like a bug in the app rather than a jump in the clock.
        // ###########################################################################################
        private static TimeSpan Clamp(TimeSpan duration)
        {
            return duration < TimeSpan.Zero ? TimeSpan.Zero : duration;
        }
    }
}
