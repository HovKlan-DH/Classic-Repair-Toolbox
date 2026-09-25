using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A WORKBOOK CITES, READ ONCE PER VERSION OF THE FILE (code review, 2026-09-25).
    //
    // DataTreeUsage parses every master and board workbook in a tree with EPPlus - a few seconds for
    // the whole BETA tree. The maintainer application asks for it on EVERY click on a queue row (the
    // removal preview in the submission detail), and again on Approve, so a maintainer moving between
    // three submissions paid for the whole tree three times over although nothing had changed.
    //
    // Keyed on the full path, and valid only while the file's LENGTH and LAST-WRITE TIME are what
    // they were when it was read - the same rule PublishedFileHashes uses for hashes. A publish
    // writes a new workbook (a new write time), so the next read sees it.
    //
    // *** ONLY FOR SHOWING, NEVER FOR DELETING. *** The preview and the administrator's list use
    // this; UnusedFileRemover does not, and re-reads every workbook fresh before it removes
    // anything. A stale answer here can therefore only make the list shown differ from what the
    // server would remove - which the approval refuses (409) - never delete a file in use.
    //
    // A failed read is not remembered: an unreadable workbook is tried again next time, and a
    // file that changed WHILE it was read is not remembered either.
    //
    // Thread-safe: requests read it concurrently. One entry per workbook ever seen, which is a few
    // dozen - it needs no eviction.
    // ###########################################################################################
    public sealed class WorkbookReadCache
    {
        private sealed record Entry(long Length, long WriteTimeUtcTicks, IReadOnlyCollection<string> Values);

        // The two kinds of read are kept apart: a master's listing and a board's citations are
        // different answers about different files.
        private readonly ConcurrentDictionary<string, Entry> thisCitations = new(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, Entry> thisListings = new(StringComparer.Ordinal);

        private int thisReads;

        // How many times a workbook was actually opened - for tests.
        internal int Reads => Volatile.Read(ref this.thisReads);

        public delegate bool Reader(string fullPath, out IReadOnlyCollection<string> values, out string why);

        // What a board workbook cites (BoardDataReader.TryCollectReferencedLocalFiles).
        public bool TryGetCitations(string fullPath, out IReadOnlyCollection<string> cites)
        {
            return this.TryGet(
                this.thisCitations,
                fullPath,
                static (string path, out IReadOnlyCollection<string> values, out string why) =>
                {
                    why = string.Empty;
                    bool ok = BoardDataReader.TryCollectReferencedLocalFiles(path, out HashSet<string> files);
                    values = files;
                    return ok;
                },
                out cites,
                out _);
        }

        // What a master workbook lists, read by `reader` (DataTreeUsage owns the master's layout).
        internal bool TryGetListing(string fullPath, Reader reader, out IReadOnlyCollection<string> listed, out string why) =>
            this.TryGet(this.thisListings, fullPath, reader, out listed, out why);

        private bool TryGet(
            ConcurrentDictionary<string, Entry> map,
            string fullPath,
            Reader reader,
            out IReadOnlyCollection<string> values,
            out string why)
        {
            why = string.Empty;

            (long Length, long Ticks)? before = WorkbookReadCache.Stamp(fullPath);

            if (before is not null &&
                map.TryGetValue(fullPath, out Entry? cached) &&
                cached.Length == before.Value.Length &&
                cached.WriteTimeUtcTicks == before.Value.Ticks)
            {
                values = cached.Values;
                return true;
            }

            Interlocked.Increment(ref this.thisReads);

            if (!reader(fullPath, out values, out why))
            {
                map.TryRemove(fullPath, out _);
                return false;
            }

            // Remembered only when the file is the same after the read as before it.
            (long Length, long Ticks)? after = WorkbookReadCache.Stamp(fullPath);

            if (before is not null && after is not null && before.Value == after.Value)
                map[fullPath] = new Entry(before.Value.Length, before.Value.Ticks, values);
            else
                map.TryRemove(fullPath, out _);

            return true;
        }

        private static (long Length, long Ticks)? Stamp(string fullPath)
        {
            try
            {
                var info = new FileInfo(fullPath);

                return info.Exists ? (info.Length, info.LastWriteTimeUtc.Ticks) : null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException)
            {
                return null;
            }
        }
    }
}
