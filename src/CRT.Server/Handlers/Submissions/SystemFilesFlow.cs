using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A SYSTEM'S FILES AS BETA HOLDS THEM, for the Systems screen's Files view (owner request,
    // 2026-10-03: "Files (should not show changed files - just list all files)") -
    // POST /api/review/systems/files. The rule is CRT.Data's SystemFileEntries.ForSystem; this
    // gathers what only the server can see:
    //
    //   - every file in the system's own BETA folder (SubmissionFileTreeFlow.FilesIn - the very
    //     listing a submission's tree is built from, hidden and retired files left out);
    //   - every file its board cites that BETA holds - the shared files are the ones outside the
    //     folder. A cited file BETA does not hold is left out: the view lists what is there.
    //
    // Asked for when the view is first shown for a system, like a submission's tree: it walks the
    // board's folder and reads its workbook. Each file carries its size (TreeFileSizes, 2026-10-04).
    // ###########################################################################################
    //
    // THE STABLE SOURCE'S TOO (owner request, 2026-10-04: a BETA / Stable switch on Files): the same
    // walk over the stable source's tree, each file opened from there (`source` Production).
    public static class SystemFilesFlow
    {
        // Null when the tree holds nothing of the system.
        public static async Task<IReadOnlyList<SystemFileEntry>?> BuildAsync(
            string dataTreeRoot,
            string systemId,
            PublishedBoardReader boards,
            CancellationToken cancellationToken = default,
            SystemFileSource source = SystemFileSource.Beta)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(dataTreeRoot);
            ArgumentNullException.ThrowIfNull(boards);

            string root = Path.GetFullPath(dataTreeRoot);
            PublishedBoardLocation location = PublishedBoardLocator.LocateSystem(root, systemId);

            if (!location.Exists)
                return null;

            BoardData? board = await boards.TryReadSystemAsync(root, systemId, cancellationToken).ConfigureAwait(false);

            if (board is null)
                return null;

            List<string> cited = SubmissionManifestBuilder
                .CollectReferencedFiles(board)
                .Where(path => SubmissionPathRules.TryResolve(root, path, out string resolved, out _) && File.Exists(resolved))
                .ToList();

            // Every entry opens from this one tree, so it is the root for whichever source they name.
            return TreeFileSizes.Attach(
                SystemFileEntries.ForSystem(SubmissionFileTreeFlow.FilesIn(root, location.SystemFolder), cited, source),
                betaRoot: root,
                productionRoot: root);
        }
    }
}
