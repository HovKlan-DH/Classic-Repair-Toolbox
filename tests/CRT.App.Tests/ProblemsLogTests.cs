using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ProblemsLog - every warning and error of the run, for the screen (owner request, 2026-11-07: "I think we must show every
// ERROR and WARNING in here in the UI, and then the logfile is not so relevant for a human to read anymore"). Logger fills
// the run's own only once it writes (Logger.Initialize, never called by a test), so these hand an instance their lines.
// ###########################################################################################
public sealed class ProblemsLogTests
{
    private static LoggedProblem Error(string message, int second = 0) =>
        new(new DateTime(2026, 11, 7, 19, 10, second), ProblemLevel.Error, message);

    // ###########################################################################################
    // *** THE LINE A PERSON READS: what and why, never a stack trace. *** A handler's last guard logs the whole exception;
    // shown, it is its first line with the exception's type name left out and its bracket closed.
    // ###########################################################################################
    [Theory]
    [InlineData(
        "The Contribute tab's [UI or Table] failed - [System.InvalidOperationException: Sequence contains no elements\r\n   at System.Linq.Enumerable.First()",
        "The Contribute tab's [UI or Table] failed - [Sequence contains no elements]")]
    [InlineData(
        "Failed to export workbook [#3]: [The process cannot access the file because it is being used by another process.]",
        "Failed to export workbook [#3]: [The process cannot access the file because it is being used by another process.]")]
    [InlineData(
        "Failed to load board file [Commodore/C64/250407/board.json]",
        "Failed to load board file [Commodore/C64/250407/board.json]")]
    [InlineData(
        "OnOpenWorkbooksFolderClick failed - [System.IO.IOException: The device is not ready.\n   at ...]",
        "OnOpenWorkbooksFolderClick failed - [The device is not ready.]")]
    [InlineData("Last error: 5 [scope]", "Last error: 5 [scope]")]
    public void What_is_shown_is_the_first_line_without_exception_type_names(string logged, string shown)
    {
        Assert.Equal(shown, ProblemsLog.ReadableOf(logged));
    }

    // A crash report says what it met.
    [Fact]
    public void A_crash_report_reads_as_what_it_met()
    {
        string report = "=== CRASH REPORT ===\nTime      : 2026-11-07 19:09:51\nSeverity  : FATAL\nException : System.ArgumentException\nMessage   : Attempt to call InvalidateArrange on wrong LayoutManager.\nStack trace:\n   at Avalonia...";

        Assert.Equal("CRT met an error it did not expect: Attempt to call InvalidateArrange on wrong LayoutManager.", ProblemsLog.ReadableOf(report));
        Assert.Equal("CRT met an error it did not expect", ProblemsLog.ReadableOf("=== CRASH REPORT ===\nTime      : x"));
    }

    // Newest first, the oldest let go of past the limit - a loop logging one failure must not fill the memory.
    [Fact]
    public void It_lists_the_newest_first_and_keeps_only_the_latest()
    {
        var log = new ProblemsLog();
        var added = new List<LoggedProblem>();
        log.Added += added.Add;

        for (int i = 0; i < ProblemsLog.Limit + 5; i++)
        {
            log.Add(Error($"Number {i}", i % 60));
        }

        Assert.Equal(ProblemsLog.Limit, log.Count);
        Assert.Equal(ProblemsLog.Limit + 5, added.Count);
        Assert.Equal($"Number {ProblemsLog.Limit + 4}", log.Newest()[0].Message);
        Assert.Equal("Number 5", log.Newest()[^1].Message);
    }

    // ###########################################################################################
    // *** WATCHING TAKES WHAT IS HELD AND WHAT COMES AFTER AS ONE STEP (code review, 2026-11-07). *** Main listened, then
    // read what the log held: an error logged on another thread in between was in both, and the card counted it twice.
    // Errors are logged from another thread while the log is watched - each must be handed over exactly once.
    // ###########################################################################################
    [Fact]
    public async Task What_watching_finds_and_what_it_is_handed_after_never_overlap()
    {
        const int Count = 300;

        for (int round = 0; round < 300; round++)
        {
            var log = new ProblemsLog();
            using var started = new ManualResetEventSlim();

            Task adding = Task.Run(() =>
            {
                started.Set();

                for (int i = 0; i < Count; i++)
                {
                    log.Add(Error($"Number {i}", i % 60));
                }
            });

            started.Wait();
            var handed = new System.Collections.Concurrent.ConcurrentQueue<LoggedProblem>();
            IReadOnlyList<LoggedProblem> held = log.Watch(handed.Enqueue);
            await adding;

            Assert.Equal(Count, held.Count + handed.Count);
            Assert.Equal(
                Enumerable.Range(0, Count).Select(i => $"Number {i}"),
                held.Reverse().Concat(handed).Select(problem => problem.Message));
        }
    }

    // "Copy all": every line in full, oldest first, under which CRT it was.
    [Fact]
    public void Copy_all_holds_each_line_in_full_oldest_first()
    {
        LoggedProblem later = Error("Second\n   at stack", 30);
        var earlier = new LoggedProblem(new DateTime(2026, 11, 7, 19, 10, 5), ProblemLevel.Warning, "First");

        string text = ProblemsLog.AsText([later, earlier], "4.0.0-alpha.3");

        Assert.Equal(
            "Errors and warnings of Classic Repair Toolbox 4.0.0-alpha.3\n\n2026-11-07 19:10:05 Warning: First\n\n2026-11-07 19:10:30 Error: Second\n   at stack\n",
            text);
    }
}
