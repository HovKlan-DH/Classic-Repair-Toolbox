using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// CrashReports - the last time CRT closed by itself, read from the crash file at the next launch (owner request,
// 2026-11-07: errors in the UI, "not to be dependent on manually viewing that logfile"). A fatal error ends CRT with
// nothing on screen; the next launch says it once.
// ###########################################################################################
public sealed class CrashReportsTests
{
    private static string Report(string time, string severity, string message) =>
        $"=== CRASH REPORT ===\r\nTime      : {time}\r\nSeverity  : {severity}\r\nSource    : Dispatcher (UI thread)\r\n" +
        $"Exception : System.ArgumentException\r\nMessage   : {message}\r\nStack trace:\r\n   at Somewhere()\r\n\r\n";

    // The newest FATAL report - a non-fatal one is no closing, CRT went on; the reports are as CrashLogger writes them.
    [Fact]
    public void The_newest_fatal_report_is_the_last_closing()
    {
        string file =
            Report("2026-11-03 09:43:35", "NON-FATAL", "Index was out of range.") +
            Report("2026-11-05 10:00:00", "FATAL", "Older one.") +
            Report("2026-11-07 19:09:51", "FATAL", "Attempt to call InvalidateArrange on wrong LayoutManager.") +
            Report("2026-11-07 19:30:00", "NON-FATAL", "Later, but CRT went on.");

        CrashReports.FatalReport? last = CrashReports.LatestFatal(file);

        Assert.NotNull(last);
        Assert.Equal(new DateTime(2026, 11, 7, 19, 9, 51), last.When);
        Assert.Contains("wrong LayoutManager", last.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("Later, but", last.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void No_file_or_no_fatal_report_is_no_closing()
    {
        Assert.Null(CrashReports.LatestFatal(null));
        Assert.Null(CrashReports.LatestFatal(string.Empty));
        Assert.Null(CrashReports.LatestFatal(Report("2026-11-03 09:43:35", "NON-FATAL", "x")));
        Assert.Null(CrashReports.LatestFatal("=== CRASH REPORT ===\r\nTime      : not a time\r\nSeverity  : FATAL\r\n"));
    }

    // ###########################################################################################
    // Said once: a report newer than the one stored is not said yet; the stored one, or older, is. With NOTHING stored -
    // the first launch of a CRT that keeps it - no report is said (code review, 2026-11-07): the crash file is never
    // emptied, so its newest fatal report may be from any run ever, and was said as "the last time it ran".
    // ###########################################################################################
    [Fact]
    public void A_report_is_said_once_and_none_with_nothing_stored()
    {
        var report = new CrashReports.FatalReport(new DateTime(2026, 11, 7, 19, 9, 51), "x");

        Assert.False(CrashReports.IsNotSaidYet(report, null));
        Assert.True(CrashReports.IsNotSaidYet(report, "2026-11-05 10:00:00"));
        Assert.False(CrashReports.IsNotSaidYet(report, "not a time"));
        Assert.False(CrashReports.IsNotSaidYet(report, "2026-11-07 19:09:51"));
        Assert.False(CrashReports.IsNotSaidYet(report, "2026-11-08 08:00:00"));
        Assert.Equal("2026-11-07 19:09:51", CrashReports.TimeTextOf(report));

        Assert.False(CrashReports.KnowsWhatWasSaid(null));
        Assert.False(CrashReports.KnowsWhatWasSaid("not a time"));
        Assert.True(CrashReports.KnowsWhatWasSaid("2026-11-05 10:00:00"));
        Assert.Equal("2026-11-07 20:00:00", CrashReports.TimeTextOf(new DateTime(2026, 11, 7, 20, 0, 0, 500)));
    }

    // What the card says: when, what it met, and where the report is.
    [Fact]
    public void The_card_says_when_what_and_where_the_report_is()
    {
        CrashReports.FatalReport report = CrashReports.LatestFatal(Report("2026-11-07 19:09:51", "FATAL", "Attempt to call InvalidateArrange on wrong LayoutManager."))!;

        Assert.Equal(
            "CRT closed unexpectedly the last time it ran, on 2026-11-07 at 19:09. It met: Attempt to call InvalidateArrange on wrong LayoutManager. " +
            "\"Show details\" has the report - you can send it to the developer through the \"Feedback\" tab.",
            CrashReports.NoticeOf(report));
    }
}
