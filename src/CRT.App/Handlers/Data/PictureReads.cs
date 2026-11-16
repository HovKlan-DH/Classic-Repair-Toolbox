using Avalonia;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE "Schematics" TAB'S PICTURES, READ ONCE AND THE LAST FEW KEPT (2026-11-16, performance walkthrough).
    //
    // A picture was read from disk and decoded whole every time it was shown, the view blank meanwhile - going back and
    // forth between a board's top and bottom pictures, the usual way of working, decoded each at every switch (some
    // 80-400 ms apiece) - and on a board switch only after every thumbnail was decoded. Now a picture has ONE read while
    // its file stays as it is (its length and write time - BoardDataReader.FileStamp, the one stamp rule), shared by
    // whoever asks: the board load starts reading the picture it will show as the thumbnails decode, and the selection
    // takes that read. The newest Kept reads are kept, with what each measured (the picture's own empty border,
    // ImageFrameGeometry); one past them is disposed once read - never the one shown (KeepOnly is told it).
    //
    // *** NEVER MORE THAN DecodedAtOnce DECODES, THE NEWEST FIRST, AND ONE OVERTAKEN BEFORE IT BEGAN IS NEVER MADE (code
    // review, 2026-11-16). *** A board load begins its picture's read, and boards picked one after another - the wheel over
    // the "Board" drop-down - began one each, all decoding at once: several hundred MB and the CPU shared with the board
    // actually chosen. So a read waits for its turn, the newest waiting first, and one pushed past the newest Kept while
    // still waiting - or let go of by KeepOnly - comes to nothing at once. Whoever asked for it has been overtaken.
    //
    // *** A READ THAT CAME TO NOTHING IS NOT KEPT (code review, 2026-11-16): *** a picture held for a moment by a virus
    // scanner as it was shown stayed "missing" until three others were read. The next ask reads it again.
    //
    // Generic over the picture so the rule tests without a bitmap; the tab's is an Avalonia Bitmap.
    // ###########################################################################################
    internal sealed class PictureReads<TPicture>
        where TPicture : class
    {
        // Two: the picture shown can be decoding while the next one asked for begins.
        internal const int DecodedAtOnce = 2;

        private sealed class Entry(string path, string stamp)
        {
            public string Path { get; } = path;

            public string Stamp { get; } = stamp;

            public TaskCompletionSource<(TPicture? Picture, Thickness Border)> Source { get; } =
                new(TaskCreationOptions.RunContinuationsAsynchronously);

            public Task<(TPicture? Picture, Thickness Border)> Read => this.Source.Task;

            // Its decode begun (or never to be: it came to nothing before its turn).
            public bool Started { get; set; }
        }

        private readonly object thisGate = new();

        // Newest first.
        private readonly List<Entry> thisEntries = new();

        private readonly int thisKept;

        private readonly Func<string, (TPicture? Picture, Thickness Border)> thisDecode;

        private readonly Action<TPicture> thisDispose;

        private int thisReadsStarted;

        private int thisDecoding;

        private int thisDecodesMade;

        // `decode` runs on the pool; `dispose` is handed a picture let go of, once read.
        public PictureReads(int kept, Func<string, (TPicture? Picture, Thickness Border)> decode, Action<TPicture> dispose)
        {
            this.thisKept = Math.Max(1, kept);
            this.thisDecode = decode;
            this.thisDispose = dispose;
        }

        // How many reads have been asked for, how many decodes made, and how many reads are held - for the tests.
        internal int ReadsStartedForTests
        {
            get
            {
                lock (this.thisGate)
                {
                    return this.thisReadsStarted;
                }
            }
        }

        internal int DecodesMadeForTests
        {
            get
            {
                lock (this.thisGate)
                {
                    return this.thisDecodesMade;
                }
            }
        }

        internal int HeldForTests
        {
            get
            {
                lock (this.thisGate)
                {
                    return this.thisEntries.Count;
                }
            }
        }

        // ###########################################################################################
        // The picture at `path`, read: the read made already or under way while the file stays as it is, else a new one -
        // decoded on the pool in its turn (above). Safe from any thread.
        // ###########################################################################################
        public Task<(TPicture? Picture, Thickness Border)> Read(string path)
        {
            string stamp = BoardDataReader.FileStamp(path);
            List<Entry> overtaken;
            Entry entry;

            lock (this.thisGate)
            {
                int at = this.thisEntries.FindIndex(known =>
                    string.Equals(known.Path, path, StringComparison.OrdinalIgnoreCase) && known.Stamp == stamp);

                if (at >= 0)
                {
                    Entry known = this.thisEntries[at];
                    this.thisEntries.RemoveAt(at);
                    this.thisEntries.Insert(0, known);
                    return known.Read;
                }

                this.thisReadsStarted++;
                entry = new Entry(path, stamp);
                this.thisEntries.Insert(0, entry);

                // Waiting still, past the newest kept: overtaken.
                overtaken = [];

                for (int index = this.thisEntries.Count - 1; index >= this.thisKept; index--)
                {
                    if (!this.thisEntries[index].Started)
                    {
                        overtaken.Add(this.thisEntries[index]);
                        this.thisEntries[index].Started = true;
                        this.thisEntries.RemoveAt(index);
                    }
                }

                this.StartDecodes();
            }

            PictureReads<TPicture>.ComeToNothing(overtaken);

            return entry.Read;
        }

        // Under the gate: the newest reads waiting begun, while fewer than DecodedAtOnce decode.
        private void StartDecodes()
        {
            foreach (Entry entry in this.thisEntries)
            {
                if (this.thisDecoding >= PictureReads<TPicture>.DecodedAtOnce)
                {
                    return;
                }

                if (entry.Started)
                {
                    continue;
                }

                entry.Started = true;
                this.thisDecoding++;
                this.thisDecodesMade++;
                _ = Task.Run(() => this.Decode(entry));
            }
        }

        private void Decode(Entry entry)
        {
            (TPicture? Picture, Thickness Border) result;

            try
            {
                result = this.thisDecode(entry.Path);
            }
            catch (Exception)
            {
                result = (null, default);
            }

            lock (this.thisGate)
            {
                this.thisDecoding--;

                if (result.Picture is null)
                {
                    this.thisEntries.Remove(entry);
                }

                this.StartDecodes();
            }

            entry.Source.TrySetResult(result);
        }

        private static void ComeToNothing(IEnumerable<Entry> entries)
        {
            foreach (Entry entry in entries)
            {
                entry.Source.TrySetResult((null, default));
            }
        }

        // ###########################################################################################
        // Lets go of every read past the newest Kept - never the one whose picture is `current`, the one shown - each
        // disposed once it is read, or come to nothing at once while still waiting for its turn. Called as a picture is
        // shown, or none.
        // ###########################################################################################
        public void KeepOnly(TPicture? current)
        {
            var letGo = new List<Entry>();
            var neverBegun = new List<Entry>();

            lock (this.thisGate)
            {
                int kept = 0;

                foreach (Entry entry in this.thisEntries)
                {
                    bool isCurrent = current is not null
                        && entry.Read.IsCompletedSuccessfully
                        && ReferenceEquals(entry.Read.Result.Picture, current);

                    if (isCurrent || kept < this.thisKept)
                    {
                        kept++;
                        continue;
                    }

                    if (entry.Started)
                    {
                        letGo.Add(entry);
                    }
                    else
                    {
                        entry.Started = true;
                        neverBegun.Add(entry);
                    }
                }

                foreach (Entry entry in letGo)
                {
                    this.thisEntries.Remove(entry);
                }

                foreach (Entry entry in neverBegun)
                {
                    this.thisEntries.Remove(entry);
                }
            }

            PictureReads<TPicture>.ComeToNothing(neverBegun);

            foreach (Entry entry in letGo)
            {
                _ = entry.Read.ContinueWith(
                    read =>
                    {
                        if (read.IsCompletedSuccessfully && read.Result.Picture is { } picture && !ReferenceEquals(picture, current))
                        {
                            this.thisDispose(picture);
                        }
                    },
                    TaskScheduler.Default);
            }
        }
    }
}
