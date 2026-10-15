using System;
using System.IO;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // The log file's lines, gathered and written in batches (code review, 2026-10-13).
    //
    // Logger used to open, append to and close the file for EVERY line, under a lock, on whatever
    // thread logged - the UI thread included: 655 opens in a 15-second idle session. Now a line is
    // kept in memory and written with the others at the next Flush - which Logger calls at least
    // once a second, at exit and in the crash handlers.
    //
    // *** A WARNING OR A CRITICAL LINE IS WRITTEN AT ONCE, with everything before it. *** Those are
    // the lines read after something went wrong, and the process may be about to die.
    //
    // A write that fails keeps its lines for the next try, up to MaxPendingCharacters - past that
    // the oldest are dropped, rather than a log nobody can write growing without end.
    // ###########################################################################################
    internal sealed class BufferedLogFile
    {
        internal const int MaxPendingCharacters = 1_000_000;

        private readonly object thisSync = new();
        private readonly StringBuilder thisPending = new();

        internal BufferedLogFile(string path)
        {
            this.Path = path;
        }

        internal string Path { get; }

        internal int PendingCharacters
        {
            get
            {
                lock (this.thisSync)
                {
                    return this.thisPending.Length;
                }
            }
        }

        // Whether a line of this category is written at once rather than with the next batch.
        internal static bool WritesAtOnce(LogCategory category) =>
            category is LogCategory.WARN or LogCategory.CRIT;

        internal void Append(string line, LogCategory category)
        {
            lock (this.thisSync)
            {
                this.thisPending.Append(line).Append(Environment.NewLine);

                if (BufferedLogFile.WritesAtOnce(category))
                {
                    this.FlushLocked();
                }
                else if (this.thisPending.Length > MaxPendingCharacters)
                {
                    this.thisPending.Remove(0, this.thisPending.Length - MaxPendingCharacters);
                }
            }
        }

        // Writes every line gathered so far. Never throws.
        internal void Flush()
        {
            lock (this.thisSync)
            {
                this.FlushLocked();
            }
        }

        private void FlushLocked()
        {
            if (this.thisPending.Length == 0)
            {
                return;
            }

            try
            {
                File.AppendAllText(this.Path, this.thisPending.ToString());
                this.thisPending.Clear();
            }
            catch
            {
                // Kept for the next try (see the header) - a log that cannot be written must not
                // take anything else down with it.
                if (this.thisPending.Length > MaxPendingCharacters)
                {
                    this.thisPending.Remove(0, this.thisPending.Length - MaxPendingCharacters);
                }
            }
        }
    }
}
