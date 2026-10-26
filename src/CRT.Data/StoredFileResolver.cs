using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHERE A CRT 4.0 BOARD'S FILE IS ON DISK - from the board's own name for it.
    //
    // A board file names each of its files ("Component images/7406.png") with the SHA-256 of its
    // bytes (BoardDocument.Files); the bytes are stored under the checksum
    // (DataTreeLayout.StoredFilePath) in one or more STORES, each a folder holding "Files/". A
    // published board has one store, the data tree's content folder. A draft has two, tried in
    // order: its own folder (the bytes of files the published tree does not have - added or
    // replaced by the contributor), then the data tree's.
    //
    // A name the board does not list, or bytes no store holds, is null - the caller's "missing
    // file" handling, as with a missing file in the Excel trees.
    // ###########################################################################################
    public sealed class StoredFileResolver
    {
        private readonly Dictionary<string, BoardDocument.StoredFileRow> thisFiles;
        private readonly IReadOnlyList<string> thisStores;

        public StoredFileResolver(BoardDocument document, IEnumerable<string> stores)
        {
            ArgumentNullException.ThrowIfNull(document);

            this.thisFiles = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in document.Files)
                this.thisFiles.TryAdd(file.Name, file);

            this.thisStores = (stores ?? []).Where(store => !string.IsNullOrWhiteSpace(store)).ToList();
        }

        // The board's entry for a name, matched ignoring capitals as the names are; null when the
        // board has no file of that name.
        public BoardDocument.StoredFileRow? EntryOf(string? name) =>
            !string.IsNullOrWhiteSpace(name) && this.thisFiles.TryGetValue(name.Trim(), out BoardDocument.StoredFileRow? file)
                ? file
                : null;

        // The full path of the file's bytes in the first store that holds them, or null.
        public string? PathOf(string? name)
        {
            if (this.EntryOf(name) is not BoardDocument.StoredFileRow file || !DataTreeLayout.IsSha256(file.Sha256))
                return null;

            string relative = DataTreeLayout.StoredFilePath(file.Sha256, file.Name).Replace('/', Path.DirectorySeparatorChar);

            return this.thisStores
                .Select(store => Path.Combine(store, relative))
                .FirstOrDefault(File.Exists);
        }
    }
}
