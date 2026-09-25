using System.Globalization;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for StartupTimeline - the startup instrumentation added so that "the app feels slow to
// launch" can be turned into a number per phase and per OS.
//
// The class takes both the process start time and each "now" as arguments precisely so these tests
// can drive it with a fixed clock. Nothing here starts a process or reads the real time, apart from
// the single sanity check on TryResolveProcessStartTime at the bottom.
public class StartupTimelineTests
{
    private static readonly DateTime ProcessStart = new(2026, 8, 24, 19, 40, 37, DateTimeKind.Local);

    // -------------------------------------------------------------- FormatDuration

    [Theory]
    [InlineData(0, "0 ms")]
    [InlineData(1, "1 ms")]
    [InlineData(842, "842 ms")]
    [InlineData(999, "999 ms")]
    [InlineData(1000, "1.00 s")]
    [InlineData(3140, "3.14 s")]
    [InlineData(65000, "65.00 s")]
    public void Durations_below_a_second_are_milliseconds_and_the_rest_are_seconds(int milliseconds, string expected)
    {
        Assert.Equal(expected, StartupTimeline.FormatDuration(TimeSpan.FromMilliseconds(milliseconds)));
    }

    // The unit is chosen from the ROUNDED millisecond count, not the raw one. Choosing first and
    // rounding second would print "1000 ms" for anything in [999.5, 1000) - a unit that never
    // appears anywhere else in the log.
    [Fact]
    public void A_duration_that_rounds_up_to_a_full_second_is_reported_in_seconds()
    {
        Assert.Equal("1.00 s", StartupTimeline.FormatDuration(TimeSpan.FromMilliseconds(999.6)));
    }

    // Startup timings get read off logs sent in by users, whose machines are not all English.
    // Without InvariantCulture a Danish or German locale writes "3,14 s", which breaks both reading
    // the log at a glance and grepping across logs from different users.
    [Theory]
    [InlineData("da-DK")]
    [InlineData("de-DE")]
    [InlineData("en-US")]
    public void Durations_use_a_decimal_point_regardless_of_the_machine_locale(string culture)
    {
        var original = CultureInfo.CurrentCulture;

        try
        {
            CultureInfo.CurrentCulture = new CultureInfo(culture);
            Assert.Equal("3.14 s", StartupTimeline.FormatDuration(TimeSpan.FromMilliseconds(3140)));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    // -------------------------------------------------------------- Record

    [Fact]
    public void The_first_milestone_reports_the_same_total_and_phase()
    {
        var timeline = new StartupTimeline(ProcessStart);

        var line = timeline.Record("Runtime and UI framework ready", ProcessStart.AddMilliseconds(1200));

        Assert.Equal("Startup milestone [Runtime and UI framework ready] [total 1.20 s] [phase 1.20 s]", line);
    }

    // The whole point of the class: "total" stays measured from process start while "phase" shows
    // what the step itself cost, so a slow phase can be identified without subtracting timestamps
    // by hand.
    [Fact]
    public void A_later_milestone_reports_total_from_process_start_but_phase_from_the_previous_milestone()
    {
        var timeline = new StartupTimeline(ProcessStart);

        timeline.Record("Runtime and UI framework ready", ProcessStart.AddMilliseconds(1200));
        var second = timeline.Record("Splash visible", ProcessStart.AddMilliseconds(1500));
        var third = timeline.Record("Data initialised", ProcessStart.AddMilliseconds(4500));

        Assert.Equal("Startup milestone [Splash visible] [total 1.50 s] [phase 300 ms]", second);
        Assert.Equal("Startup milestone [Data initialised] [total 4.50 s] [phase 3.00 s]", third);
    }

    // The system clock can move backwards during startup - an NTP correction, a resumed VM, or a
    // user changing the time. A negative duration in the log reads as a bug in the app rather than
    // a jump in the clock, so both figures clamp at zero instead.
    [Fact]
    public void A_clock_that_jumps_backwards_reports_zero_rather_than_a_negative_duration()
    {
        var timeline = new StartupTimeline(ProcessStart);

        timeline.Record("Runtime and UI framework ready", ProcessStart.AddMilliseconds(1200));
        var afterJump = timeline.Record("Splash visible", ProcessStart.AddSeconds(-5));

        Assert.Equal("Startup milestone [Splash visible] [total 0 ms] [phase 0 ms]", afterJump);
    }

    // A milestone recorded earlier than the previous one must not make the NEXT phase absorb the
    // gap it skipped; the timeline restarts from wherever the clock now says it is.
    [Fact]
    public void A_milestone_after_a_backwards_jump_measures_its_phase_from_the_jumped_to_time()
    {
        var timeline = new StartupTimeline(ProcessStart);

        timeline.Record("Runtime and UI framework ready", ProcessStart.AddSeconds(10));
        timeline.Record("Splash visible", ProcessStart.AddSeconds(1));
        var third = timeline.Record("Data initialised", ProcessStart.AddSeconds(3));

        Assert.Equal("Startup milestone [Data initialised] [total 3.00 s] [phase 2.00 s]", third);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void A_milestone_with_no_name_is_still_logged_under_a_placeholder(string milestone)
    {
        var timeline = new StartupTimeline(ProcessStart);

        var line = timeline.Record(milestone, ProcessStart.AddMilliseconds(500));

        Assert.Equal("Startup milestone [(unnamed)] [total 500 ms] [phase 500 ms]", line);
    }

    [Fact]
    public void A_milestone_name_is_trimmed_so_the_log_columns_stay_aligned()
    {
        var timeline = new StartupTimeline(ProcessStart);

        var line = timeline.Record("  Splash visible  ", ProcessStart.AddMilliseconds(500));

        Assert.Equal("Startup milestone [Splash visible] [total 500 ms] [phase 500 ms]", line);
    }

    // -------------------------------------------------------------- TryResolveProcessStartTime

    // Reads the CURRENT process - it never starts one, so this stays inside the "no processes in
    // tests" rule. It is allowed to return null (the class is built to handle a platform that
    // refuses the value); what it must never do is throw or claim the process started in the future.
    [Fact]
    public void Resolving_the_process_start_time_either_succeeds_with_a_past_time_or_returns_null()
    {
        var startTime = StartupTimeline.TryResolveProcessStartTime();

        if (startTime.HasValue)
        {
            Assert.True(startTime.Value <= DateTime.Now.AddSeconds(1), $"Process start time [{startTime.Value}] is in the future");
        }
    }
}
