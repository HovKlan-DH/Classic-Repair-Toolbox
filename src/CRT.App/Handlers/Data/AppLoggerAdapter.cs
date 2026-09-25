namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Forwards CRT.Data's ICrtLog gateway (CrtLog) to this app's own Logger, so a class that moved
    // into CRT.Data (BoardDataReader, BoardDataWriter, BoardComponentHighlightStorage,
    // DataValidator) writes to the exact same log file, in the exact same format, as it did before
    // the Phase 1 extraction. See CrtLog's own header comment in CRT.Data for why the seam exists.
    //
    // Installed once, from App.OnFrameworkInitializationCompleted, immediately after
    // Logger.Initialize() - CrtLog.Sink must never be set before the log file itself is ready to
    // receive writes.
    // ###########################################################################################
    internal sealed class AppLoggerAdapter : ICrtLog
    {
        public void Debug(string message) => Logger.Debug(message);
        public void Info(string message) => Logger.Info(message);
        public void Warning(string message) => Logger.Warning(message);
        public void Critical(string message) => Logger.Critical(message);
    }
}
