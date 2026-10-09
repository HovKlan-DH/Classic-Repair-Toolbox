using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A BOARD'S FILES AS BETA HOLDS THEM, for the Boards screen's Files view (owner request,
    // 2026-10-03: "Files (should not show changed files - just list all files)") -
    // POST /api/review/boards/files. The rule is CRT.Data's BoardFileEntries.ForBoard; this
    // gathers what only the server can see:
    //
    //   - every file in the board's own BETA folder (SubmissionFileTreeFlow.FilesIn - the very
    //     listing a submission's tree is built from, hidden and retired files left out);
    //   - every file its board cites that BETA holds - the shared files are the ones outside the
    //     folder. A cited file BETA does not hold is left out: the view lists what is there.
    //
    // Asked for when the view is first shown for a board, like a submission's tree: it walks the
    // board's folder and reads its workbook. Each file carries its size (TreeFileSizes, 2026-10-04).
    // ###########################################################################################
    //
    // THE STABLE SOURCE'S TOO (owner request, 2026-10-04: a BETA / Stable switch on Files): the same
    // walk over the stable source's tree, each file opened from there (`source` Production).
    public static class BoardFilesFlow
    {
        // Null when the tree holds nothing of the board.
        public static async Task<IReadOnlyList<BoardFileEntry>?> BuildAsync(
            string dataTreeRoot,
            string boardId,
            PublishedBoardReader boards,
            CancellationToken cancellationToken = default,
            BoardFileSource source = BoardFileSource.Beta)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(dataTreeRoot);
            ArgumentNullException.ThrowIfNull(boards);

            string root = Path.GetFullPath(dataTreeRoot);
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(root, boardId);

            if (!location.Exists)
                return null;

            BoardData? board = await boards.TryReadBoardAsync(root, boardId, cancellationToken).ConfigureAwait(false);

            if (board is null)
                return null;

            List<string> cited = SubmissionManifestBuilder
                .CollectReferencedFiles(board)
                .Where(path => SubmissionPathRules.TryResolve(root, path, out string resolved, out _) && File.Exists(resolved))
                .ToList();

            // Every entry opens from this one tree, so it is the root for whichever source they name.
            return TreeFileSizes.Attach(
                BoardFileEntries.ForBoard(SubmissionFileTreeFlow.FilesIn(root, location.BoardFolder), cited, source),
                betaRoot: root,
                productionRoot: root);
        }
    }
}
