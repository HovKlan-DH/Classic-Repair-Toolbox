using Handlers.DataHandling;

namespace CRT.Server
{
    // ###########################################################################################
    // Forwards CRT.Data's logging seam to the server's own ILogger, so a class shared with the
    // desktop app logs into the journal here and into the app's log file there, without either
    // knowing about the other.
    //
    // This is the ServerLogAdapter that ICrtLog.cs's own header names as the intended Phase 3
    // wiring - it is installed once in Program.cs, immediately after the host is built. Before
    // that point CrtLog.Sink is null and every CRT.Data log call is inert, which is exactly what
    // the whole test suite relies on (no test installs a sink, so no test writes a log).
    //
    // MAPPING. CRT.Data's four levels map onto the nearest ILogger level. "Critical" is mapped to
    // Error rather than LogCritical deliberately: CRT.Data uses Critical for "this operation
    // failed" (an unreadable workbook, a missing file), not for "the process is doomed", and
    // journald's crit level is worth reserving for the latter.
    // ###########################################################################################
    public sealed class ServerLogAdapter : ICrtLog
    {
        private readonly ILogger logger;

        public ServerLogAdapter(ILogger<ServerLogAdapter> logger)
        {
            ArgumentNullException.ThrowIfNull(logger);
            this.logger = logger;
        }

        // The message is already a formatted string by the time it reaches here - CRT.Data builds
        // it with interpolation - so it is passed as a single argument rather than as a template.
        // Using it AS a template would make any "{" in a file name or a board label throw a format
        // exception inside logging, which is a spectacularly bad place to throw.
        public void Debug(string message) => this.logger.LogDebug("{Message}", message);

        public void Info(string message) => this.logger.LogInformation("{Message}", message);

        public void Warning(string message) => this.logger.LogWarning("{Message}", message);

        public void Critical(string message) => this.logger.LogError("{Message}", message);
    }
}
