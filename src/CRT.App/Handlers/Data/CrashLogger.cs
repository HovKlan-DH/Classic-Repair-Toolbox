using CRT;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Captures unexpected crashes to disk so a user can send them in.
    //
    // This exists because "Logger" alone could not do the job, for two reasons that both cost a
    // real crash report:
    //
    // 1. "Logger.Initialize" TRUNCATES the log on every launch. A crash is nearly always reported
    //    after the user has relaunched the application - which is exactly the act that erases the
    //    evidence. So a crash is ALSO written to its own file (AppConfig.CrashFileName), which is
    //    only ever appended to and is never truncated at startup.
    //
    // 2. "Logger.Write" opens, appends and closes per line, which is what makes it safe here -
    //    there is no buffered stream that could still be unflushed when the process dies. But it
    //    silently swallows write failures, and it does nothing at all before "Logger.Initialize"
    //    has run. A crash during startup - before the log file exists - would therefore vanish.
    //    "WriteCrashReport" resolves its own path and so does not depend on Logger being ready.
    //
    // A crash entry is deliberately VERBOSE where the ordinary log is terse: the full exception
    // chain including every inner exception, the stack trace, the application version, the OS and
    // runtime, and which handler caught it. The whole point is that the maintainer receives it as
    // a file rather than as a description of what the user thinks they saw.
    // ###########################################################################################
    public static class CrashLogger
    {
        private static readonly object _lock = new();

        // Set once at startup so a crash report can name the build it came from without reaching
        // back into Avalonia (which may itself be the thing that is broken).
        private static string _appVersion = "unknown";

        // Guards against a crash storm writing thousands of entries. A repeating render-thread
        // exception can fire many times per second, and an unbounded writer would fill the disk
        // and bury the FIRST report - which is the one that explains the cause.
        private const int MaximumReportsPerSession = 20;
        private static int _reportsWritten;

        // The exception most recently written as a full report, held ONLY to recognise the same
        // instance coming back from another handler - see TryRecordRepeatOfLastCrash.
        //
        // This does keep one Exception object alive until the next crash replaces it. That is
        // deliberate and bounded: exactly one, for a process that is usually about to die anyway,
        // and the alternative (comparing formatted text) would match two genuinely distinct faults
        // that happened to read alike.
        private static Exception? _lastReportedException;

        // ###########################################################################################
        // Records the application version to stamp on every crash report. Safe to call more than
        // once; the last value wins.
        // ###########################################################################################
        public static void Initialize(string appVersion)
        {
            if (!string.IsNullOrWhiteSpace(appVersion))
            {
                _appVersion = appVersion;
            }
        }

        // ###########################################################################################
        // Returns the full path of the crash file, or an empty string when it cannot be resolved.
        //
        // Deliberately resolved on every call rather than cached at startup: this must keep working
        // when the crash happened BEFORE any initialisation ran.
        // ###########################################################################################
        public static string ResolveCrashFilePath()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

                if (string.IsNullOrEmpty(appData))
                {
                    return string.Empty;
                }

                var directory = Path.Combine(appData, AppConfig.AppFolderName);
                Directory.CreateDirectory(directory);
                return Path.Combine(directory, AppConfig.CrashFileName);
            }
            catch
            {
                return string.Empty;
            }
        }

        // ###########################################################################################
        // Writes one crash report to BOTH the crash file and the ordinary log.
        //
        // "source" names the handler that caught it (e.g. "Dispatcher"), which is the single most
        // useful field when reading a report: it says whether the crash killed the process, was
        // survived on the UI thread, or came from a task nobody awaited.
        //
        // "isFatal" only affects the wording. Nothing here decides whether the process lives - that
        // is the caller's business, and the handlers in App make that choice explicitly.
        // ###########################################################################################
        public static void Log(string source, Exception? exception, bool isFatal)
        {
            // ONE crash, one report - even though up to three handlers see it.
            //
            // A UI-thread exception reaches the Dispatcher filter, then the Dispatcher unhandled
            // handler, then AppDomain, and each calls this. Writing all three produced near-identical
            // entries with the same stack trace under three different "Source" lines, so a single
            // fault read as three separate ones to anyone opening the file - and it spent the
            // per-session budget three times as fast during a crash storm.
            //
            // The later handlers are still subscribed, deliberately: the filter does not run on
            // every path, so dropping one would trade a duplicate for a MISS, and this file exists
            // precisely because a missed crash is the expensive failure. Instead the repeat is
            // recorded as one extra line naming the handler that also saw it, which keeps the
            // "which handler caught it" information that made Source worth having.
            if (exception != null && TryRecordRepeatOfLastCrash(source, exception))
            {
                return;
            }

            string report;

            try
            {
                report = BuildReport(source, exception, isFatal);
            }
            catch (Exception buildFailure)
            {
                // Never let the crash reporter itself throw - it runs from handlers that are the
                // last thing standing between a crash and silence.
                report = $"Crash report could not be built for [{source}]: {buildFailure.Message}";
            }

            // The ordinary log first, so a report still lands there for anyone already reading it,
            // and so the crash file and the log agree.
            try
            {
                Logger.Critical(report);
            }
            catch
            {
                // Absorbed on purpose; the crash file below is the copy that matters.
            }

            WriteCrashFile(report);
        }

        // ###########################################################################################
        // Recognises the SAME exception instance arriving from a second (or third) handler, and
        // records it as a one-line addition to the report already written rather than as a new one.
        //
        // Returns true when the caller should write nothing further.
        //
        // Matched by REFERENCE, not by message or stack text: the handlers are all handed the very
        // same Exception object as it travels outward, so reference equality identifies exactly the
        // re-reported case and nothing else. Two genuinely separate faults that happen to share a
        // message - the same failing operation retried, say - are different objects and are each
        // reported in full, which is what a reader needs.
        //
        // Only the LAST exception is remembered. A crash is re-reported by its other handlers
        // immediately, within the same unwind, so a one-deep memory covers the real case; holding a
        // set would mean deciding when to forget entries, and a genuine recurrence of the same
        // instance much later is worth a fresh report anyway.
        //
        // The reference is held only to compare against, and is replaced by the next crash - see
        // the field's own note.
        // ###########################################################################################
        private static bool TryRecordRepeatOfLastCrash(string source, Exception exception)
        {
            lock (_lock)
            {
                if (!ReferenceEquals(_lastReportedException, exception))
                {
                    _lastReportedException = exception;
                    return false;
                }
            }

            var line = $"[Also seen by: {(string.IsNullOrWhiteSpace(source) ? "unknown" : source)}]";

            try
            {
                Logger.Critical(line);
            }
            catch
            {
                // Same reasoning as Log's own absorbed Logger call - the crash file is the copy
                // that matters.
            }

            AppendRepeatLine(line);
            return true;
        }

        // ###########################################################################################
        // Adds the "[Also seen by: ...]" line to the crash file, immediately under the report it
        // belongs to.
        //
        // Deliberately NOT routed through WriteCrashFile: this is a continuation of a report that
        // has already been counted, so it must not consume a second slot of the per-session budget
        // nor trigger the "further reports suppressed" notice. It is also skipped once the budget
        // is spent, so a suppressed storm cannot keep appending these forever.
        // ###########################################################################################
        private static void AppendRepeatLine(string line)
        {
            lock (_lock)
            {
                if (_reportsWritten >= MaximumReportsPerSession)
                {
                    return;
                }

                try
                {
                    var path = ResolveCrashFilePath();

                    if (string.IsNullOrEmpty(path))
                    {
                        return;
                    }

                    File.AppendAllText(path, line + Environment.NewLine + Environment.NewLine, Encoding.UTF8);
                }
                catch
                {
                    // A crash reporter that throws while reporting a crash helps nobody.
                }
            }
        }

        // ###########################################################################################
        // Appends a report to the crash file, creating it if needed.
        //
        // Uses one open-append-close per report for the same reason Logger does: a buffered writer
        // held open across the process's death is not guaranteed to reach the disk, and a crash
        // report that was never flushed is the exact failure this class exists to prevent.
        // ###########################################################################################
        private static void WriteCrashFile(string report)
        {
            lock (_lock)
            {
                if (_reportsWritten >= MaximumReportsPerSession)
                {
                    return;
                }

                try
                {
                    var path = ResolveCrashFilePath();

                    if (string.IsNullOrEmpty(path))
                    {
                        return;
                    }

                    var text = report + Environment.NewLine + Environment.NewLine;

                    File.AppendAllText(path, text, Encoding.UTF8);

                    // Counted only once the report is actually ON DISK. Incrementing before the
                    // attempt spent the budget on reports that were never written - a path that
                    // could not be resolved, or a write that threw - so a transient failure early
                    // in a session silently cost the reports that came after it, and a budget
                    // exhausted entirely by failed writes never even reached the suppression notice
                    // below, leaving the file with no indication anything had been lost.
                    _reportsWritten++;

                    TrimCrashFileIfOversized(path);

                    if (_reportsWritten == MaximumReportsPerSession)
                    {
                        File.AppendAllText(
                            path,
                            $"[Further crash reports in this session are suppressed after {MaximumReportsPerSession} entries]"
                                + Environment.NewLine + Environment.NewLine,
                            Encoding.UTF8);
                    }
                }
                catch
                {
                    // A crash reporter that throws while reporting a crash helps nobody.
                }
            }
        }

        // ###########################################################################################
        // Keeps the crash file from growing without limit, discarding the OLDEST reports once it
        // passes MaximumCrashFileBytes.
        //
        // MaximumReportsPerSession bounds one RUN of the application; this file deliberately
        // survives every restart, which is the whole point of it existing separately from the log.
        // So an installation with a repeating fault writes its 20 reports per launch forever, and
        // nothing here ever gave that a ceiling. It matters more than an ordinary log would,
        // because the Feedback tab attaches this file to every submission and zips it into memory -
        // an unbounded file means an ordinary feature request uploading years of crash history.
        //
        // The NEWEST reports are the ones kept. That is the opposite of what the per-session cap
        // does (it keeps the FIRST reports of a storm, since those explain the cause) and both are
        // right for their own scale: within one crash storm the first report is the informative
        // one, but across months of launches the recent crashes are the ones a user is writing in
        // about, and a years-old fault from a version long since replaced is not worth the space.
        //
        // Trimmed on a REPORT BOUNDARY, never mid-report: reports are separated by a blank line, so
        // the retained text is cut at the first separator inside the kept window. A file cut at an
        // arbitrary byte offset would open on half a stack trace with no header saying what it
        // belonged to.
        //
        // Failures are swallowed like every other write here - a trim that cannot run leaves a
        // large file, which is far better than a crash reporter that throws while reporting.
        // ###########################################################################################
        // Internal so a test can build a file that genuinely exceeds it rather than hardcoding a
        // size that silently stops exercising the trim if this value ever changes.
        internal const long MaximumCrashFileBytes = 2 * 1024 * 1024;

        // How much of an oversized file is kept, as a fraction of the cap. Well below 1 on purpose:
        // trimming to exactly the cap would put the file back over it on the very next report and
        // rewrite the whole thing again each time, so the headroom is what stops the trim running
        // on every single write.
        private const double CrashFileTrimRetainFraction = 0.5;

        // Internal rather than private so the trim can be driven against a temp file in a test.
        // The rest of this class's file writing goes through Log, which resolves the user's real
        // AppData path and which no test may call - this is the one piece of that path whose rule
        // (keep the newest, cut on a report boundary) is worth pinning, and it takes its path as an
        // argument, so a test never has to touch the real crash file. See CrashLoggerTests.
        internal static void TrimCrashFileIfOversized(string path)
        {
            try
            {
                var info = new FileInfo(path);

                if (!info.Exists || info.Length <= MaximumCrashFileBytes)
                {
                    return;
                }

                var text = File.ReadAllText(path, Encoding.UTF8);

                var keepFrom = text.Length - (int)(MaximumCrashFileBytes * CrashFileTrimRetainFraction);

                if (keepFrom <= 0)
                {
                    return;
                }

                // Forward to the next report boundary so the file never opens mid-report. The
                // separator is the blank line WriteCrashFile puts after every entry.
                var separator = Environment.NewLine + Environment.NewLine;
                var boundary = text.IndexOf(separator, keepFrom, StringComparison.Ordinal);

                if (boundary < 0)
                {
                    // No boundary in the kept window - one enormous report, or a file written by
                    // some other means. Leaving it alone beats cutting it at an arbitrary point.
                    return;
                }

                var retained = text[(boundary + separator.Length)..];

                var notice = $"[Older crash reports were trimmed to keep this file under "
                    + $"{MaximumCrashFileBytes / (1024 * 1024)} MB]"
                    + separator;

                File.WriteAllText(path, notice + retained, Encoding.UTF8);
            }
            catch
            {
                // See the header: a failed trim costs disk space, a throw costs the crash report.
            }
        }

        // ###########################################################################################
        // Formats one crash report.
        //
        // Kept as pure string work taking an Exception, so it is unit testable without provoking a
        // real crash - the handlers in App are a thin rim over this.
        // ###########################################################################################
        internal static string BuildReport(string source, Exception? exception, bool isFatal)
        {
            var builder = new StringBuilder();

            var severity = isFatal ? "FATAL" : "NON-FATAL";

            builder.Append("=== CRASH REPORT ===").Append(Environment.NewLine);
            builder.Append("Time      : ")
                .Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
                .Append(Environment.NewLine);
            builder.Append("Severity  : ").Append(severity).Append(Environment.NewLine);
            builder.Append("Source    : ").Append(string.IsNullOrWhiteSpace(source) ? "unknown" : source)
                .Append(Environment.NewLine);
            builder.Append("Version   : ").Append(_appVersion).Append(Environment.NewLine);

            // The environment lines matter because several classes of crash here are
            // platform-specific (Skia/GPU on the render thread, file paths, the MiniPro binary that
            // only ships on Windows), and a bug report rarely says which OS it came from.
            builder.Append("OS        : ").Append(SafeDescribe(() => RuntimeInformation.OSDescription))
                .Append(Environment.NewLine);
            builder.Append("Runtime   : ").Append(SafeDescribe(() => RuntimeInformation.FrameworkDescription))
                .Append(Environment.NewLine);
            builder.Append("Arch      : ").Append(SafeDescribe(() => RuntimeInformation.OSArchitecture.ToString()))
                .Append(Environment.NewLine);

            if (exception == null)
            {
                builder.Append("Exception : <none supplied>").Append(Environment.NewLine);
                return builder.ToString();
            }

            // Every inner exception is written out separately rather than relying on
            // "exception.ToString()" alone. An AggregateException - which is what an unobserved
            // Task hands over - prints its inner stack traces in a run that is genuinely hard to
            // read, and the innermost exception is usually the actual fault.
            var chain = Flatten(exception);

            for (var index = 0; index < chain.Count; index++)
            {
                var current = chain[index];

                var label = index == 0
                    ? "Exception"
                    : $"Inner exception #{index}";

                builder.Append(Environment.NewLine);
                builder.Append(label).Append(" : ").Append(current.GetType().FullName).Append(Environment.NewLine);
                builder.Append("Message   : ").Append(current.Message).Append(Environment.NewLine);

                if (!string.IsNullOrWhiteSpace(current.StackTrace))
                {
                    builder.Append("Stack trace:").Append(Environment.NewLine);
                    builder.Append(current.StackTrace).Append(Environment.NewLine);
                }
                else
                {
                    // An exception that was constructed but never thrown carries no stack trace.
                    // Saying so is better than a blank space the reader has to interpret.
                    builder.Append("Stack trace: <none>").Append(Environment.NewLine);
                }
            }

            return builder.ToString();
        }

        // ###########################################################################################
        // Walks an exception's inner chain, expanding AggregateException into each of its inner
        // exceptions rather than just the first.
        //
        // Depth limited: a malformed or self-referencing chain must not spin here, since this runs
        // while the application is already failing.
        // ###########################################################################################
        private static List<Exception> Flatten(Exception exception)
        {
            var results = new List<Exception>();
            var queue = new Queue<Exception>();
            queue.Enqueue(exception);

            var seen = new HashSet<Exception>(ReferenceEqualityComparer.Instance);

            while (queue.Count > 0 && results.Count < 25)
            {
                var current = queue.Dequeue();

                if (!seen.Add(current))
                {
                    continue;
                }

                results.Add(current);

                if (current is AggregateException aggregate)
                {
                    foreach (var inner in aggregate.InnerExceptions)
                    {
                        queue.Enqueue(inner);
                    }

                    continue;
                }

                if (current.InnerException != null)
                {
                    queue.Enqueue(current.InnerException);
                }
            }

            return results;
        }

        // ###########################################################################################
        // Reads one environment string, substituting a placeholder if it throws. These are only
        // context lines, and none of them is worth losing the stack trace over.
        // ###########################################################################################
        private static string SafeDescribe(Func<string> read)
        {
            try
            {
                return read() ?? "unknown";
            }
            catch
            {
                return "unknown";
            }
        }
    }
}
