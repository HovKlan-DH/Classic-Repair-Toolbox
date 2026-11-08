using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** THE LAST TIME CRT CLOSED BY ITSELF (owner request, 2026-11-07: errors and warnings "in the UI ... not to be
    // dependent on manually viewing that logfile"). *** A fatal error ends CRT with no word on screen - CrashLogger writes
    // its report to the crash file and the process goes. So the NEXT launch reads that file and, for a fatal report newer
    // than the last one said (UserSettings.LastCrashReportSaid), Main's card says CRT closed unexpectedly and when, the
    // report itself listed in "Errors and warnings" to copy for the "Feedback" tab (Main.Problems.cs).
    //
    // Pure, over the file's text as CrashLogger.BuildReport writes it ("=== CRASH REPORT ===", "Time      : ...",
    // "Severity  : FATAL"), so every rule is tested without a crash. A non-fatal report is not a closing: CRT went on.
    // ###########################################################################################
    internal static class CrashReports
    {
        internal sealed record FatalReport(DateTime When, string Text);

        private const string ReportStart = "=== CRASH REPORT ===";

        private static readonly Regex TimeLine = new(@"^Time\s*:\s*(?<time>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s*$", RegexOptions.Multiline | RegexOptions.CultureInvariant);

        private static readonly Regex FatalLine = new(@"^Severity\s*:\s*FATAL\s*$", RegexOptions.Multiline | RegexOptions.CultureInvariant);

        // The newest FATAL report in `text`, or null - none, or none whose time can be read.
        public static FatalReport? LatestFatal(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            FatalReport? latest = null;
            int start = text.IndexOf(CrashReports.ReportStart, StringComparison.Ordinal);

            while (start >= 0)
            {
                int next = text.IndexOf(CrashReports.ReportStart, start + CrashReports.ReportStart.Length, StringComparison.Ordinal);
                string report = (next < 0 ? text[start..] : text[start..next]).TrimEnd();

                Match time = CrashReports.TimeLine.Match(report);

                if (time.Success &&
                    CrashReports.FatalLine.IsMatch(report) &&
                    DateTime.TryParseExact(time.Groups["time"].Value, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime when) &&
                    (latest is null || when >= latest.When))
                {
                    latest = new FatalReport(when, report);
                }

                start = next;
            }

            return latest;
        }

        // ###########################################################################################
        // Whether `report` is one not said yet - newer than the one `saidUpTo` names (stored as CrashLogger's time text).
        //
        // *** NOTHING STORED SAYS NOTHING (code review, 2026-11-07). *** The crash file is never emptied, so with nothing
        // stored - the first launch of a CRT that keeps it - its newest fatal report may be from any run ever (an earlier
        // pre-release's, September's), and was said as "the last time it ran" though runs had ended well since. Main stores
        // the launch's own time then (KnowsWhatWasSaid), so a crash of that run is said at the next launch.
        // ###########################################################################################
        public static bool IsNotSaidYet(FatalReport report, string? saidUpTo)
        {
            ArgumentNullException.ThrowIfNull(report);

            return CrashReports.TryReadTime(saidUpTo, out DateTime said) && report.When > said;
        }

        // Whether `saidUpTo` says how far the reports were said - false for nothing stored, or a time that cannot be read.
        public static bool KnowsWhatWasSaid(string? saidUpTo) => CrashReports.TryReadTime(saidUpTo, out _);

        public static string TimeTextOf(FatalReport report) => CrashReports.TimeTextOf(report.When);

        public static string TimeTextOf(DateTime when) => when.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

        private static bool TryReadTime(string? text, out DateTime when) =>
            DateTime.TryParseExact(text, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out when);

        // "CRT closed unexpectedly the last time it ran, on 2026-11-07 at 19:09."
        public static string NoticeOf(FatalReport report)
        {
            ArgumentNullException.ThrowIfNull(report);

            string message = ProblemsLog.ReadableOf(report.Text);
            string why = message.StartsWith("CRT met an error it did not expect: ", StringComparison.Ordinal)
                ? " It met: " + message["CRT met an error it did not expect: ".Length..].TrimEnd('.') + "."
                : string.Empty;

            return $"CRT closed unexpectedly the last time it ran, on {report.When.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)} at " +
                   $"{report.When.ToString("HH:mm", CultureInfo.InvariantCulture)}.{why} \"Show details\" has the report - you can send it to the " +
                   "developer through the \"Feedback\" tab.";
        }
    }
}
