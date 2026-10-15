using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// BufferedLogFile - the log written in batches rather than one file open per line (code review,
// 2026-10-13: 655 opens in a 15-second idle session, on the UI thread too). Logger itself is never
// initialised from a test (it would write the user's real log), so its file half is tested here,
// on a temp file.
public sealed class BufferedLogFileTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static string[] LinesOf(string path) =>
        File.Exists(path) ? File.ReadAllLines(path) : [];

    [Fact]
    public void Ordinary_lines_wait_for_the_next_flush()
    {
        string path = this.thisWorkspace.Path_("crt.log");
        var file = new BufferedLogFile(path);

        file.Append("one", LogCategory.INFO);
        file.Append("two", LogCategory.DBUG);

        Assert.Empty(LinesOf(path));

        file.Flush();

        Assert.Equal(["one", "two"], LinesOf(path));
        Assert.Equal(0, file.PendingCharacters);
    }

    // The lines read after something went wrong are on disk at once - with everything before
    // them, in order - in case the process is about to die.
    [Theory]
    [InlineData(LogCategory.WARN)]
    [InlineData(LogCategory.CRIT)]
    public void A_warning_or_critical_line_is_written_at_once_with_everything_before_it(LogCategory category)
    {
        string path = this.thisWorkspace.Path_("crt.log");
        var file = new BufferedLogFile(path);

        file.Append("before", LogCategory.INFO);
        file.Append("it went wrong", category);

        Assert.Equal(["before", "it went wrong"], LinesOf(path));
    }

    [Fact]
    public void Flushes_append_to_what_is_already_there()
    {
        string path = this.thisWorkspace.WriteFile("crt.log", "earlier" + Environment.NewLine);
        var file = new BufferedLogFile(path);

        file.Append("later", LogCategory.INFO);
        file.Flush();
        file.Flush();

        Assert.Equal(["earlier", "later"], LinesOf(path));
    }

    // A write that fails keeps its lines for the next try - and never throws at the caller.
    [Fact]
    public void Lines_a_failed_write_could_not_place_are_written_by_the_next_flush()
    {
        string folder = this.thisWorkspace.Path_("not yet");
        string path = Path.Combine(folder, "crt.log");
        var file = new BufferedLogFile(path);

        file.Append("kept", LogCategory.INFO);
        file.Flush();

        Assert.True(file.PendingCharacters > 0);

        Directory.CreateDirectory(folder);
        file.Flush();

        Assert.Equal(["kept"], LinesOf(path));
    }

    // A log nobody can write must not grow without end: past the cap, the oldest lines go.
    [Fact]
    public void Unwritable_lines_are_capped_keeping_the_newest()
    {
        var file = new BufferedLogFile(Path.Combine(this.thisWorkspace.Path_("missing"), "crt.log"));
        string line = new('x', 1_000);

        for (int i = 0; i < 1_500; i++)
        {
            file.Append(line, LogCategory.INFO);
        }

        file.Append("newest", LogCategory.INFO);

        Assert.True(file.PendingCharacters <= BufferedLogFile.MaxPendingCharacters);
    }

    [Theory]
    [InlineData(LogCategory.DBUG, false)]
    [InlineData(LogCategory.INFO, false)]
    [InlineData(LogCategory.WARN, true)]
    [InlineData(LogCategory.CRIT, true)]
    public void Only_warnings_and_critical_lines_are_written_at_once(LogCategory category, bool expected)
    {
        Assert.Equal(expected, BufferedLogFile.WritesAtOnce(category));
    }
}
