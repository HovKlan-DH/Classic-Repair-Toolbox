using CRT;
using System;
using System.IO;
using System.Threading;

namespace Handlers.DataHandling
{
    public enum LogCategory
    {
        DBUG,
        INFO,
        WARN,
        CRIT
    }

    public static class Logger
    {
        // How often gathered lines are written at the latest (BufferedLogFile).
        internal static readonly TimeSpan FlushInterval = TimeSpan.FromSeconds(1);

        // Null until Initialize - and then nothing is written anywhere (the test suite relies on it).
        private static BufferedLogFile? _file;
        private static Timer? _flushTimer;

        // ###########################################################################################
        // Resolves the log file path in a persistent AppData folder that survives Velopack updates,
        // and overwrites any previous log content. Must be called once at application startup.
        //
        // Lines are written in batches (BufferedLogFile): at least once a second, at once for a
        // warning or a critical line, and at exit - ProcessExit covers every way out, Velopack's
        // restart into an update included; Program.Main and the crash handlers flush as well.
        // ###########################################################################################
        public static void Initialize()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var directory = Path.Combine(appData, AppConfig.AppFolderName);
                Directory.CreateDirectory(directory);
                var logFilePath = Path.Combine(directory, AppConfig.LogFileName);
                File.WriteAllText(logFilePath, string.Empty);

                _file = new BufferedLogFile(logFilePath);
                _flushTimer = new Timer(_ => Flush(), null, FlushInterval, FlushInterval);
                AppDomain.CurrentDomain.ProcessExit += (_, _) =>
                {
                    // The timer is held in a field so it is not collected; at exit it is stopped,
                    // then the last lines written.
                    _flushTimer?.Dispose();
                    Flush();
                };
            }
            catch
            {
                _file = null;
            }
        }

        // ###########################################################################################
        // Writes every line gathered so far. Called by the timer, at exit, by the crash handlers,
        // and before anything reads the log file (the Feedback tab attaches it). Never throws.
        // ###########################################################################################
        public static void Flush() => _file?.Flush();

        // ###########################################################################################
        // Writes a Debug-level log entry.
        // ###########################################################################################
        public static void Debug(string message) => Write(LogCategory.DBUG, message);

        // ###########################################################################################
        // Writes a Debug-level log entry directly from an exception to include the raw stack trace.
        // ###########################################################################################
        public static void Debug(Exception ex, string message = "")
        {
            var msg = string.IsNullOrEmpty(message) ? ex.ToString() : $"{message} - {ex}";
            Write(LogCategory.DBUG, msg);
        }

        // ###########################################################################################
        // Writes an Info-level log entry.
        // ###########################################################################################
        public static void Info(string message) => Write(LogCategory.INFO, message);

        // ###########################################################################################
        // Writes a Warning-level log entry.
        // ###########################################################################################
        public static void Warning(string message) => Write(LogCategory.WARN, message);

        // ###########################################################################################
        // Writes a Critical-level log entry - something that went wrong, which Main's "Something went wrong" card says at
        // once (ProblemsLog). `alreadyShown`: the place it happened says it itself, so it is only listed.
        // ###########################################################################################
        public static void Critical(string message) => Write(LogCategory.CRIT, message);

        public static void Critical(string message, bool alreadyShown) => Write(LogCategory.CRIT, message, alreadyShown);

        // ###########################################################################################
        // Formats a single timestamped log line and hands it to the buffered file - thread-safe. DEBUG
        // entries too, always, since "Write a detailed log" went (owner request, 2026-11-07: "remove this
        // checkbox and then always write this detailed logfile"). A warning or an error goes on the run's
        // list as well, for the screen (ProblemsLog) - only once the log writes at all.
        // ###########################################################################################
        // ###########################################################################################
        // What a test sees of the lines logged, and at which level - Logger writes nothing in a test (Initialize is never
        // called there), so a rule about WHICH level a failure is logged at is read here (LoggedLines, in the tests). Null
        // in CRT.
        // ###########################################################################################
        internal static Action<LogCategory, string>? SeenForTests { get; set; }

        private static void Write(LogCategory category, string message, bool alreadyShown = false)
        {
            Logger.SeenForTests?.Invoke(category, message);

            var file = _file;

            if (file == null)
                return;

            DateTime now = DateTime.Now;
            file.Append($"{now:yyyy-MM-dd HH:mm:ss} {category} {message}", category);

            if (category is LogCategory.WARN or LogCategory.CRIT)
            {
                try
                {
                    ProblemsLog.Current.Add(new LoggedProblem(
                        now,
                        category == LogCategory.CRIT ? ProblemLevel.Error : ProblemLevel.Warning,
                        message,
                        ShownElsewhere: alreadyShown));
                }
                catch
                {
                    // A listener failing must never stop a line being logged.
                }
            }
        }
    }
}
