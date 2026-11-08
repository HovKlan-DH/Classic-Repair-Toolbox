using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // *** EVERY WARNING AND ERROR OF THE RUN, WHERE A PERSON CAN READ THEM (owner request, 2026-11-07: "I think we must
    // show every ERROR and WARNING in here in the UI, and then the logfile is not so relevant for a human to read
    // anymore"; and "careful evaluate this and implement what you believe is the most optimal solution as seen from a
    // user perspective"). ***
    //
    // Logger hands every WARN and CRIT line here as it writes it - once it writes at all (Logger.Initialize), so a test
    // run fills nothing. What a person meets:
    //
    //   - A CRIT line is something that went wrong - a button whose work failed, a board that could not be read, a file
    //     that could not be saved: Main's "Something went wrong" card says it at once (Main.Problems.cs), unless the
    //     place it happened says it already (`ShownElsewhere` - Logger.Critical(message, alreadyShown: true)).
    //   - A WARN line is something worth knowing that needs nothing done - a download given up, a file skipped: listed
    //     only, as a card per warning would greet everybody working offline.
    //   - Both are listed, newest first, in "Errors and warnings" (ProblemsWindow), opened from the card and from the
    //     "Configuration" tab, with "Copy all" for the "Feedback" tab.
    //
    // The latest `Limit` are kept - a loop logging the same failure must not fill the memory.
    // ###########################################################################################
    internal enum ProblemLevel
    {
        Warning,
        Error,
    }

    internal sealed record LoggedProblem(DateTime When, ProblemLevel Level, string Message, bool ShownElsewhere = false)
    {
        // The line a person reads - see ProblemsLog.ReadableOf.
        public string Readable => ProblemsLog.ReadableOf(this.Message);

        public string LevelText => this.Level == ProblemLevel.Error ? "Error" : "Warning";

        public string TimeText => this.When.ToString("HH:mm:ss", CultureInfo.InvariantCulture);

        public bool IsError => this.Level == ProblemLevel.Error;
    }

    internal sealed class ProblemsLog
    {
        public const int Limit = 500;

        private readonly object thisLock = new();
        private readonly Queue<LoggedProblem> thisEntries = new();

        // The run's - what Logger fills.
        public static ProblemsLog Current { get; } = new();

        // The listeners, changed and read under the lock (Watch).
        private Action<LoggedProblem>? thisAdded;

        // Raised on the thread that logged it; a listener on the UI posts itself there.
        public event Action<LoggedProblem>? Added
        {
            add
            {
                lock (this.thisLock)
                {
                    this.thisAdded += value;
                }
            }

            remove
            {
                lock (this.thisLock)
                {
                    this.thisAdded -= value;
                }
            }
        }

        public void Add(LoggedProblem problem)
        {
            ArgumentNullException.ThrowIfNull(problem);

            Action<LoggedProblem>? listeners;

            lock (this.thisLock)
            {
                this.thisEntries.Enqueue(problem);

                while (this.thisEntries.Count > ProblemsLog.Limit)
                {
                    this.thisEntries.Dequeue();
                }

                // The listeners as they are when the problem is held - see Watch.
                listeners = this.thisAdded;
            }

            listeners?.Invoke(problem);
        }

        // ###########################################################################################
        // *** WHAT IS HELD, AND EVERY PROBLEM AFTER IT, AS ONE STEP (code review, 2026-11-07). *** Main listened and then
        // read what the log held: an error logged on another thread in between was in both, and its card counted it twice.
        // Here `onAdded` is listening and the log read under the same lock as Add holds a problem and takes its
        // listeners, so each problem is in what is returned (newest first) or handed to `onAdded` - never both, never
        // neither.
        // ###########################################################################################
        public IReadOnlyList<LoggedProblem> Watch(Action<LoggedProblem> onAdded)
        {
            ArgumentNullException.ThrowIfNull(onAdded);

            lock (this.thisLock)
            {
                this.thisAdded += onAdded;
                return this.thisEntries.Reverse().ToList();
            }
        }

        // What is kept, newest first.
        public IReadOnlyList<LoggedProblem> Newest()
        {
            lock (this.thisLock)
            {
                return this.thisEntries.Reverse().ToList();
            }
        }

        public int Count
        {
            get
            {
                lock (this.thisLock)
                {
                    return this.thisEntries.Count;
                }
            }
        }

        // ###########################################################################################
        // *** THE LINE A PERSON READS: what went wrong and why, never a stack trace. *** A log line is written for whoever
        // reads the log - "The Contribute tab's [Table] failed - [System.InvalidOperationException: Sequence contains no
        // elements" and the stack under it. Shown, that is "The Contribute tab's [Table] failed - [Sequence contains no
        // elements]": its first line, the exception's type name left out, a bracket left open closed. A crash report says
        // what it met.
        // ###########################################################################################
        private static readonly Regex ExceptionTypeName = new(
            @"(?<![\w.])(?:[A-Za-z_]\w*\.)*[A-Za-z_]\w*Exception:\s*",
            RegexOptions.CultureInvariant);

        private static readonly Regex CrashMessageLine = new(
            @"^Message\s*:\s*(?<message>.+)$",
            RegexOptions.CultureInvariant | RegexOptions.Multiline);

        public static string ReadableOf(string? message)
        {
            string text = message ?? string.Empty;

            if (text.StartsWith("=== CRASH REPORT ===", StringComparison.Ordinal))
            {
                Match crash = ProblemsLog.CrashMessageLine.Match(text);

                return crash.Success
                    ? "CRT met an error it did not expect: " + crash.Groups["message"].Value.Trim()
                    : "CRT met an error it did not expect";
            }

            int lineEnd = text.IndexOfAny(['\r', '\n']);
            string line = (lineEnd >= 0 ? text[..lineEnd] : text).Trim();
            line = ProblemsLog.ExceptionTypeName.Replace(line, string.Empty);

            int open = line.Count(character => character == '[');
            int close = line.Count(character => character == ']');

            return open > close ? line + new string(']', open - close) : line;
        }

        // ###########################################################################################
        // What "Copy all" puts on the clipboard: each problem in full, oldest first as a log reads, under a line saying
        // which CRT it was - to paste into the "Feedback" tab.
        // ###########################################################################################
        public static string AsText(IEnumerable<LoggedProblem> problems, string version)
        {
            ArgumentNullException.ThrowIfNull(problems);

            var text = new StringBuilder();
            text.Append("Errors and warnings of Classic Repair Toolbox ").Append(version).Append('\n');

            foreach (LoggedProblem problem in problems.OrderBy(problem => problem.When))
            {
                text.Append('\n')
                    .Append(problem.When.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
                    .Append(' ')
                    .Append(problem.LevelText)
                    .Append(": ")
                    .Append(problem.Message.Trim())
                    .Append('\n');
            }

            return text.ToString();
        }
    }
}
