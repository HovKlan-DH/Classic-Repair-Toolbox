using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A BOARD'S FILES AS BETA HOLDS THEM, for the Boards screen's Files view (owner request,
    // 2026-10-03: "Files (should not show changed files - just list all files)") -
    // POST /api/review/boards/files. The rule is CRT.Data's BoardFileEntries.ForBoard; this
    // gathers what only the server can see.
    //
    // *** IN THE CRT 4.0 TREE (2026-10-25) *** a board is its board file and the files it lists by
    // name, each one's bytes stored under their checksum (DataTreeLayout) - so the view is the
    // board file and every name it lists whose bytes the tree holds, each carrying where it lies
    // (BoardFileEntry.TreePath) and its size (TreeFileSizes, 2026-10-04). A listed file the tree
    // does not hold is left out: the view lists what is there.
    // ###########################################################################################
    //
    // THE STABLE SOURCE'S TOO (owner request, 2026-10-04: a BETA / Stable switch on Files): the same
    // look at the stable source's tree, each file opened from there (`source` Production).
    public static class BoardFilesFlow
    {
        // Null when the tree holds nothing of the board.
        public static Task<IReadOnlyList<BoardFileEntry>?> BuildAsync(
            string dataTreeRoot,
            string boardId,
            PublishedBoardReader boards,
            CancellationToken cancellationToken = default,
            BoardFileSource source = BoardFileSource.Beta)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(dataTreeRoot);
            ArgumentNullException.ThrowIfNull(boards);
            cancellationToken.ThrowIfCancellationRequested();

            string root = Path.GetFullPath(dataTreeRoot);
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(root, boardId);

            if (!location.Exists)
                return Task.FromResult<IReadOnlyList<BoardFileEntry>?>(null);

            IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? listed = PublishedBoardFiles.ByName(root, boardId);

            if (listed is null)
                return Task.FromResult<IReadOnlyList<BoardFileEntry>?>(null);

            var places = new Dictionary<string, string>(StringComparer.Ordinal);

            foreach ((string name, BoardDocument.StoredFileRow file) in listed)
            {
                string stored = PublishedBoardFiles.StoredRelative(file);

                if (stored.Length > 0 && PublishedBoardFiles.IsStored(root, stored))
                    places[name] = stored;
            }

            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForBoard(
                [DataTreeLayout.BoardFileName],
                places.Keys.ToList(),
                source);

            string boardFileRelative = DataTreeLayout.BoardFilePath(boardId);

            entries = BoardFileEntries.WithTreePaths(entries, entry =>
                entry.Path == DataTreeLayout.BoardFileName ? boardFileRelative : places.GetValueOrDefault(entry.Path));

            // Every entry opens from this one tree, so it is the root for whichever source they name.
            return Task.FromResult<IReadOnlyList<BoardFileEntry>?>(TreeFileSizes.Attach(entries, betaRoot: root, productionRoot: root));
        }
    }
}
