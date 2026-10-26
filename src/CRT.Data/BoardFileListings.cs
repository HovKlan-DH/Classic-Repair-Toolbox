using System;
using System.Collections.Generic;
using System.IO;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE BOARD FILES' LISTS AS ONE LOOKUP READS THEM - each board file read once while its stamp
    // stays the same (code review, 2026-10-25).
    //
    // DraftFiles.ListedIn trusts its shared cache only once a write has SETTLED (StampSettling, two
    // seconds), since two same-length writes inside one clock tick share a stamp. But the two seconds
    // after a save are exactly when the Drafts tab counts the draft's problems again - so each of a
    // large board's ~1,200 file lookups parsed the whole board file again, on the UI thread, after
    // every autosave. One DiskFileLookup holds one of these: within it, a board file whose stamp has
    // not moved is the file it already read. A lookup is one count, or one open table.
    // ###########################################################################################
    public sealed class BoardFileListings
    {
        private readonly object thisGate = new();
        private readonly Dictionary<string, (string Stamp, IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? Files)> thisRead =
            new(StringComparer.OrdinalIgnoreCase);

        // How many times a board file was actually read - for tests.
        internal int ReadsForTests { get; private set; }

        // The files `boardFile` lists, by name - DraftFiles.ListedIn's answer, asked again only once
        // the file's stamp has moved. Null when there is no board file there or it cannot be read.
        public IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? ListedIn(string? boardFile)
        {
            if (string.IsNullOrWhiteSpace(boardFile) || !File.Exists(boardFile))
                return null;

            string stamp = BoardDataReader.FileStamp(boardFile);

            lock (this.thisGate)
            {
                if (this.thisRead.TryGetValue(boardFile, out var known) && known.Stamp == stamp)
                    return known.Files;
            }

            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files = DraftFiles.ListedIn(boardFile);

            lock (this.thisGate)
            {
                this.ReadsForTests++;

                // Kept only when the file is still as it was when its stamp was taken.
                if (BoardDataReader.FileStamp(boardFile) == stamp)
                    this.thisRead[boardFile] = (stamp, files);
            }

            return files;
        }
    }
}
