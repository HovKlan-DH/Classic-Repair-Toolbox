using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The lines CRT logs while a test watches, with their level - Logger writes nothing in a test (Logger.Initialize is never
// called there), so a rule about the level a failure is logged at reads Logger.SeenForTests through this. Watch() in a
// using: the hook is let go of again however the test ends.
//
// While it watches, CRT.Data logs through CRT's own adapter too, as it does in CRT (CrtLog.Sink is null in a test
// otherwise, and CRT.Data's lines went nowhere) - so a line CRT.Data logs is seen at the level CRT's log gets it.
// ###########################################################################################
internal sealed class LoggedLines : IDisposable
{
    private readonly object thisLock = new();
    private readonly List<(LogCategory Category, string Message)> thisLines = new();
    private readonly ICrtLog? thisSinkBefore;

    private LoggedLines()
    {
        this.thisSinkBefore = CrtLog.Sink;
        CrtLog.Sink = new AppLoggerAdapter();
        Logger.SeenForTests = this.Add;
    }

    public static LoggedLines Watch() => new();

    public IReadOnlyList<(LogCategory Category, string Message)> Lines
    {
        get
        {
            lock (this.thisLock)
            {
                return this.thisLines.ToList();
            }
        }
    }

    // The lines at `category` holding `text`.
    public int Count(LogCategory category, string text) =>
        this.Lines.Count(line => line.Category == category && line.Message.Contains(text, StringComparison.Ordinal));

    private void Add(LogCategory category, string message)
    {
        lock (this.thisLock)
        {
            this.thisLines.Add((category, message));
        }
    }

    public void Dispose()
    {
        if (ReferenceEquals(Logger.SeenForTests?.Target, this))
        {
            Logger.SeenForTests = null;
        }

        CrtLog.Sink = this.thisSinkBefore;
    }
}
