using CRT;
using System;
using System.IO;
using System.Runtime.InteropServices;

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
        private static string _logFilePath = string.Empty;
        private static readonly object _lock = new();

        // Controls whether DEBUG-level entries are written to the log file.
        // Set from UserSettings after settings are loaded.
        public static bool IsDebugEnabled { get; set; } = false;

        // ###########################################################################################
        // Resolves the log file path in a persistent AppData folder that survives Velopack updates,
        // and overwrites any previous log content. Must be called once at application startup.
        // ###########################################################################################
        public static void Initialize()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var directory = Path.Combine(appData, AppConfig.AppFolderName);
                Directory.CreateDirectory(directory);
                _logFilePath = Path.Combine(directory, AppConfig.LogFileName);
                File.WriteAllText(_logFilePath, string.Empty);
            }
            catch
            {
                _logFilePath = string.Empty;
            }
        }

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
        // Writes a Critical-level log entry.
        // ###########################################################################################
        public static void Critical(string message) => Write(LogCategory.CRIT, message);

        // ###########################################################################################
        // Formats and appends a single timestamped log line to the log file in a thread-safe manner.
        // DEBUG entries are silently suppressed when IsDebugEnabled is false.
        // ###########################################################################################
        private static void Write(LogCategory category, string message)
        {
            if (string.IsNullOrEmpty(_logFilePath))
                return;

            if (category == LogCategory.DBUG && !IsDebugEnabled)
                return;

            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {category} {message}";

            lock (_lock)
            {
                try
                {
                    File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
                catch
                {
                    // Silently absorb write failures to avoid disrupting the application
                }
            }
        }
    }
}