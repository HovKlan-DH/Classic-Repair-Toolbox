using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// Tests for CrashLogger.BuildReport - the formatter behind the crash file a user is asked to send
// in after the application dies unexpectedly.
//
// Only the FORMATTING is tested here, and that is deliberate. The handlers themselves live in
// "App.SetupGlobalExceptionLogging" and can only be proven by actually crashing the application,
// which rule 6 (no tests that need a display or take down the process) puts out of scope; the
// handlers are written as a thin rim over this method precisely so the part that decides what a
// report SAYS is testable without provoking a real crash.
//
// Nothing here writes to the real crash file: "BuildReport" is pure string work, and "Log" - which
// resolves the user's AppData path - is never called.
public class CrashLoggerTests
{
    // Builds a real, thrown exception so it carries an actual stack trace. Constructing one with
    // "new" gives a null StackTrace, which is a different case (covered separately below).
    private static Exception Thrown(Func<Exception> create)
    {
        try
        {
            throw create();
        }
        catch (Exception ex)
        {
            return ex;
        }
    }

    // -------------------------------------------------------------- The header

    [Fact]
    public void A_report_names_the_source_so_the_reader_knows_which_handler_caught_it()
    {
        // The source is the single most useful field in a report: it distinguishes a crash that
        // killed the process from one survived on the UI thread or one from an unawaited task.
        var report = CrashLogger.BuildReport("Dispatcher (UI thread)", new InvalidOperationException("boom"), isFatal: true);

        Assert.Contains("Dispatcher (UI thread)", report);
    }

    [Fact]
    public void A_fatal_and_a_non_fatal_report_are_distinguishable()
    {
        // Whether the application actually died changes how a bug report should be read - a
        // non-fatal unobserved-task report may well be unrelated to what the user complained about.
        var fatal = CrashLogger.BuildReport("AppDomain", new Exception("x"), isFatal: true);
        var nonFatal = CrashLogger.BuildReport("Unobserved Task", new Exception("x"), isFatal: false);

        Assert.Contains("FATAL", fatal);
        Assert.Contains("NON-FATAL", nonFatal);

        // "NON-FATAL" contains "FATAL" as a substring, so the fatal report must be checked for the
        // absence of the longer word rather than the presence of the shorter one.
        Assert.DoesNotContain("NON-FATAL", fatal);
    }

    [Fact]
    public void A_report_carries_the_environment_lines_a_bug_report_rarely_states()
    {
        // Several crash classes here are platform specific (Skia on the render thread, the MiniPro
        // binary that only ships on Windows), and users seldom say which OS they are on.
        var report = CrashLogger.BuildReport("AppDomain", new Exception("x"), isFatal: true);

        Assert.Contains("OS        :", report);
        Assert.Contains("Runtime   :", report);
        Assert.Contains("Arch      :", report);
        Assert.Contains("Time      :", report);
    }

    // -------------------------------------------------------------- The exception itself

    [Fact]
    public void The_exception_type_message_and_stack_trace_are_all_written_out()
    {
        var exception = Thrown(() => new InvalidOperationException("the specific failure"));

        var report = CrashLogger.BuildReport("Dispatcher", exception, isFatal: true);

        Assert.Contains("System.InvalidOperationException", report);
        Assert.Contains("the specific failure", report);
        Assert.Contains("Stack trace:", report);

        // The frame this test threw from must actually appear - a report whose stack trace section
        // is present but empty is exactly as useless as no report at all.
        Assert.Contains(nameof(Thrown), report);
    }

    [Fact]
    public void An_exception_that_was_never_thrown_says_so_rather_than_leaving_a_blank()
    {
        // A constructed-but-unthrown exception has a null StackTrace. Saying "<none>" tells the
        // reader the trace is genuinely absent, rather than leaving them to wonder whether the
        // writer dropped it.
        var report = CrashLogger.BuildReport("Startup", new Exception("never thrown"), isFatal: true);

        Assert.Contains("Stack trace: <none>", report);
    }

    [Fact]
    public void Every_inner_exception_is_written_out_because_the_innermost_is_usually_the_real_fault()
    {
        var inner = Thrown(() => new FileNotFoundException("the actual cause"));
        var middle = new InvalidOperationException("the middle layer", inner);
        var outer = new ApplicationException("the outermost wrapper", middle);

        var report = CrashLogger.BuildReport("Dispatcher", outer, isFatal: true);

        Assert.Contains("the outermost wrapper", report);
        Assert.Contains("the middle layer", report);
        Assert.Contains("the actual cause", report);
        Assert.Contains("Inner exception #1", report);
        Assert.Contains("Inner exception #2", report);
    }

    [Fact]
    public void An_AggregateException_reports_all_of_its_inner_exceptions_not_just_the_first()
    {
        // This is the shape TaskScheduler.UnobservedTaskException hands over, and a plain walk down
        // ".InnerException" would silently report only one of several failures.
        var aggregate = new AggregateException(
            new InvalidOperationException("first failure"),
            new TimeoutException("second failure"));

        var report = CrashLogger.BuildReport("Unobserved Task", aggregate, isFatal: false);

        Assert.Contains("first failure", report);
        Assert.Contains("second failure", report);
    }

    // ###########################################################################################
    // A GENUINE cycle - an exception whose own inner exception is itself - must terminate.
    //
    // This runs while the application is already failing, so a malformed chain must not hang.
    //
    // The cycle is built by reflection because neither constructor allows it: an exception cannot
    // be passed itself as its own inner exception at construction time, and InnerException has no
    // setter. Writing the private backing field is the only way to produce the shape the guard in
    // Flatten exists for, and a test that cannot produce that shape does not test the guard.
    //
    // THE PREVIOUS VERSION OF THIS TEST DID NOT. It built an AggregateException over an EMPTY list
    // and wrapped it twice, which is a two-level DAG that terminates in three nodes on its own -
    // so it passed with the "seen" guard deleted, and the real protection shipped unverified while
    // looking covered. The assertion below is the difference: without the guard this hangs.
    // ###########################################################################################
    [Fact]
    public async Task A_self_referencing_exception_chain_terminates_instead_of_spinning()
    {
        var selfReferencing = new InvalidOperationException("points at itself");

        // "_innerException" is the private field behind Exception.InnerException, which is
        // get-only. If a future runtime renames it this test fails loudly rather than silently
        // going back to testing nothing, which is exactly what it is here to avoid.
        var innerField = typeof(Exception).GetField(
            "_innerException",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

        Assert.NotNull(innerField);
        innerField!.SetValue(selfReferencing, selfReferencing);
        Assert.Same(selfReferencing, selfReferencing.InnerException);

        // The guard's whole job is that this RETURNS. Raced against a timeout rather than called
        // directly, so a regression fails the run instead of hanging the suite until CI kills it.
        var build = Task.Run(() => CrashLogger.BuildReport("Dispatcher", selfReferencing, isFatal: true));
        var finished = await Task.WhenAny(build, Task.Delay(TimeSpan.FromSeconds(10)));

        Assert.True(ReferenceEquals(finished, build), "BuildReport did not terminate on a self-referencing chain.");
        Assert.Contains("points at itself", await build);
    }

    // The AggregateException equivalent: one that holds ITSELF among its inner exceptions, which a
    // naive walk re-enqueues forever. Same reflection reasoning as above - the constructor copies
    // the list it is given, so the self-reference has to be written in afterwards.
    [Fact]
    public async Task A_self_referencing_aggregate_exception_terminates_instead_of_spinning()
    {
        var aggregate = new AggregateException(new InvalidOperationException("real failure"));

        // Found by TYPE rather than by name: the backing field for InnerExceptions is a runtime
        // internal ("m_innerExceptions" on some versions, "_innerExceptions" on others), and
        // hardcoding either makes this test silently version-specific. There is exactly one
        // ReadOnlyCollection<Exception> field on the type, so the search is unambiguous.
        var innerField = typeof(AggregateException)
            .GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
            .Single(f => f.FieldType == typeof(System.Collections.ObjectModel.ReadOnlyCollection<Exception>));

        innerField.SetValue(aggregate, new List<Exception> { aggregate }.AsReadOnly());
        Assert.Same(aggregate, aggregate.InnerExceptions[0]);

        var build = Task.Run(() => CrashLogger.BuildReport("Unobserved Task", aggregate, isFatal: false));
        var finished = await Task.WhenAny(build, Task.Delay(TimeSpan.FromSeconds(10)));

        Assert.True(ReferenceEquals(finished, build), "BuildReport did not terminate on a self-referencing aggregate.");
        Assert.Contains("AggregateException", await build);
    }

    [Fact]
    public void A_null_exception_still_produces_a_report_rather_than_throwing()
    {
        // AppDomain.UnhandledException can carry a non-Exception object, in which case the cast
        // yields null. The reporter must not itself throw on the way to recording a crash.
        var report = CrashLogger.BuildReport("AppDomain", null, isFatal: true);

        Assert.Contains("<none supplied>", report);
        Assert.Contains("AppDomain", report);
    }

    [Fact]
    public void A_blank_source_is_labelled_rather_than_left_empty()
    {
        var report = CrashLogger.BuildReport("   ", new Exception("x"), isFatal: true);

        Assert.Contains("Source    : unknown", report);
    }

    [Fact]
    public void Each_report_opens_with_a_marker_so_several_crashes_can_be_told_apart_in_one_file()
    {
        // The crash file is append-only across restarts, so it routinely holds many reports. The
        // header is what lets a reader find where one ends and the next begins.
        var report = CrashLogger.BuildReport("Dispatcher", new Exception("x"), isFatal: true);

        Assert.StartsWith("=== CRASH REPORT ===", report);
    }

    // ###########################################################################################
    // THE SIZE TRIM.
    //
    // MaximumReportsPerSession bounds one RUN; this file deliberately survives every restart, so
    // without a size ceiling an installation with a repeating fault grows it forever - and the
    // Feedback tab attaches the whole thing to every submission.
    //
    // Driven against a temp file rather than through Log, which resolves the user's real AppData
    // path and which nothing here may call.
    // ###########################################################################################
    [Fact]
    public void An_oversized_crash_file_is_trimmed_and_keeps_the_NEWEST_reports()
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_("crash.log");

        string separator = Environment.NewLine + Environment.NewLine;

        // Padding so each report is fat enough that a handful of them clear the cap, without
        // building a file so large the test is slow.
        string padding = new('x', 64 * 1024);

        var builder = new System.Text.StringBuilder();
        int reportCount = (int)(CrashLogger.MaximumCrashFileBytes / padding.Length) + 8;

        for (int i = 0; i < reportCount; i++)
        {
            builder.Append("=== CRASH REPORT ===").Append(Environment.NewLine)
                .Append("Marker    : report-").Append(i).Append(Environment.NewLine)
                .Append(padding).Append(separator);
        }

        File.WriteAllText(path, builder.ToString(), System.Text.Encoding.UTF8);
        Assert.True(new FileInfo(path).Length > CrashLogger.MaximumCrashFileBytes);

        CrashLogger.TrimCrashFileIfOversized(path);

        var trimmed = File.ReadAllText(path, System.Text.Encoding.UTF8);

        Assert.True(new FileInfo(path).Length < CrashLogger.MaximumCrashFileBytes);

        // The LAST report survives - across months of launches the recent crashes are the ones a
        // user is writing in about - and the FIRST is the one discarded. That is deliberately the
        // opposite of the per-session cap, which keeps the first reports of one storm.
        Assert.Contains($"report-{reportCount - 1}", trimmed);
        Assert.DoesNotContain("report-0" + Environment.NewLine, trimmed);

        // And it says so, rather than silently appearing to have always been short.
        Assert.Contains("trimmed", trimmed);
    }

    // Cut on a REPORT BOUNDARY, never mid-report: a file sliced at an arbitrary byte offset opens
    // on half a stack trace with no header saying what it belonged to.
    [Fact]
    public void A_trimmed_crash_file_starts_on_a_report_boundary()
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_("crash.log");

        string separator = Environment.NewLine + Environment.NewLine;
        string padding = new('y', 64 * 1024);

        var builder = new System.Text.StringBuilder();
        int reportCount = (int)(CrashLogger.MaximumCrashFileBytes / padding.Length) + 8;

        for (int i = 0; i < reportCount; i++)
        {
            builder.Append("=== CRASH REPORT ===").Append(Environment.NewLine)
                .Append(padding).Append(separator);
        }

        File.WriteAllText(path, builder.ToString(), System.Text.Encoding.UTF8);

        CrashLogger.TrimCrashFileIfOversized(path);

        var trimmed = File.ReadAllText(path, System.Text.Encoding.UTF8);

        // After the notice line, the very next thing is a whole report's header.
        var afterNotice = trimmed[(trimmed.IndexOf(separator, StringComparison.Ordinal) + separator.Length)..];
        Assert.StartsWith("=== CRASH REPORT ===", afterNotice);
    }

    // A file already under the cap is left completely alone - this runs after every single crash
    // report, so the common case must not rewrite the file.
    [Fact]
    public void A_crash_file_under_the_cap_is_left_untouched()
    {
        using var workspace = new TempWorkspace();
        string path = workspace.Path_("crash.log");

        string original = "=== CRASH REPORT ===" + Environment.NewLine + "small"
            + Environment.NewLine + Environment.NewLine;
        File.WriteAllText(path, original, System.Text.Encoding.UTF8);

        CrashLogger.TrimCrashFileIfOversized(path);

        Assert.Equal(original, File.ReadAllText(path, System.Text.Encoding.UTF8));
    }

    // A path that does not exist must not throw - this runs from crash handlers, where an
    // exception is the one thing that cannot be allowed.
    [Fact]
    public void Trimming_a_missing_crash_file_is_harmless()
    {
        using var workspace = new TempWorkspace();

        CrashLogger.TrimCrashFileIfOversized(workspace.Path_("does-not-exist.log"));
    }
}
