using System;
using System.Threading;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A FILE WRITTEN A MOMENT LATER, OFF THE THREAD THAT ASKED (2026-11-16, performance walkthrough).
    //
    // The settings file was written whole on the UI thread at every change - one to four times a
    // board switch, each a temporary file written and swapped over the real one (AtomicJsonFile),
    // which a virus scanner can hold up. Queue hands the text over and returns; the latest text is
    // written `delay` after the first one queued, on a pool thread, every text queued in between
    // replaced by the next - and none at all when it is the very text last written.
    //
    // *** WHAT IS QUEUED AND NOT WRITTEN YET IS LOST IF THE PROCESS DIES. *** So Flush writes it at
    // once, and is called on every way out: the end of Program.Main, ProcessExit, the crash handlers,
    // before an update restarts CRT, and before another file is loaded (UserSettings.LoadFrom). Only a
    // process killed outright in the moment after a change loses that change.
    //
    // Writes never overlap and never go backwards: each takes the latest text under the one write
    // lock, so a timer's write and a Flush cannot write an older text after a newer one.
    // ###########################################################################################
    internal sealed class DeferredTextFile
    {
        private readonly object thisGate = new();
        private readonly object thisWriting = new();
        private readonly TimeSpan thisDelay;
        private readonly Func<string, string, bool> thisWrite;
        private readonly Timer thisTimer;

        private (string Path, string Text)? thisPending;
        private (string Path, string Text)? thisWritten;
        private bool thisScheduled;

        // `write` writes a text to its path, answering whether it reached the disk.
        public DeferredTextFile(TimeSpan delay, Func<string, string, bool> write)
        {
            this.thisDelay = delay;
            this.thisWrite = write;
            this.thisTimer = new Timer(_ => this.Flush(), null, Timeout.Infinite, Timeout.Infinite);
        }

        // How many texts were actually written - for tests.
        internal int WritesForTests { get; private set; }

        // Whether a text waits to be written.
        public bool IsPending
        {
            get
            {
                lock (this.thisGate)
                {
                    return this.thisPending is not null;
                }
            }
        }

        // ###########################################################################################
        // `text` is what `path` should hold; written within the delay, replacing anything queued before.
        // ###########################################################################################
        public void Queue(string path, string text)
        {
            lock (this.thisGate)
            {
                this.thisPending = (path, text);

                if (!this.thisScheduled)
                {
                    this.thisScheduled = true;
                    this.thisTimer.Change(this.thisDelay, Timeout.InfiniteTimeSpan);
                }
            }
        }

        // ###########################################################################################
        // Writes what waits, at once, on this thread - after any write already under way. False only
        // when a write was made and did not reach the disk.
        // ###########################################################################################
        public bool Flush()
        {
            lock (this.thisWriting)
            {
                (string Path, string Text)? pending;

                lock (this.thisGate)
                {
                    pending = this.thisPending;
                    this.thisPending = null;
                    this.thisScheduled = false;
                    this.thisTimer.Change(Timeout.Infinite, Timeout.Infinite);
                }

                if (pending is not { } wanted || this.IsWrittenAlready(wanted))
                {
                    return true;
                }

                bool written = this.thisWrite(wanted.Path, wanted.Text);
                this.WritesForTests++;

                this.thisWritten = written ? wanted : null;
                return written;
            }
        }

        // ###########################################################################################
        // The next text queued is written even when it is the one written last - after the file was
        // read again, it may hold what somebody typed into it since.
        // ###########################################################################################
        public void ForgetWhatWasWritten()
        {
            lock (this.thisWriting)
            {
                this.thisWritten = null;
            }
        }

        private bool IsWrittenAlready((string Path, string Text) wanted) =>
            this.thisWritten is { } written &&
            string.Equals(written.Path, wanted.Path, StringComparison.Ordinal) &&
            string.Equals(written.Text, wanted.Text, StringComparison.Ordinal);
    }
}
