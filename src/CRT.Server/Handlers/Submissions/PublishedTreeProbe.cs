using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The REAL PublishedTreeView: the questions the submission rules ask about what is published,
    // answered from the data tree on this server (security review, 2026-09-25).
    //
    // *** IN THE CRT 4.0 TREE A VIEW IS OF ONE BOARD (2026-10-25). *** A submission names its
    // files by the board's own names, so "what is published under this name" is the published
    // board's list (PublishedBoardFiles) - HashOf and SizeOf answer from it, nothing is hashed. The
    // folder listing stays the tree's own: a NEW board's folder must not differ from a published
    // one's only by capitalisation (PublishedTreeView.FindCaseVariant).
    //
    // An I/O boundary, thin on purpose - the rules that USE these answers live in CRT.Data and are
    // tested against a dictionary. Every relative path goes through SubmissionPathRules before it
    // touches the disk, the one containment rule this service has.
    // ###########################################################################################
    public static class PublishedTreeProbe
    {
        // ###########################################################################################
        // A view of the board `boardId` in the tree at `dataTreeRoot`, or NULL when there is no
        // usable root, or the board's file is there and cannot be read. Null is what the rules treat
        // as "what is published could not be consulted" - the safe answer: nothing is admitted on
        // the strength of being published already. A board the tree does not hold is an empty list.
        // ###########################################################################################
        public static PublishedTreeView? For(string? dataTreeRoot, string? boardId) =>
            PublishedTreeProbe.For(
                dataTreeRoot,
                string.IsNullOrWhiteSpace(dataTreeRoot) || !Directory.Exists(dataTreeRoot) ? null : PublishedBoardFiles.ByName(dataTreeRoot, boardId));

        // The same from the board's files already read (PublishedBoardRead.Files - an approval reads
        // the board file once); null `files` is a board file that could not be read.
        public static PublishedTreeView? For(string? dataTreeRoot, IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? files)
        {
            if (string.IsNullOrWhiteSpace(dataTreeRoot) || !Directory.Exists(dataTreeRoot) || files is null)
                return null;

            return new PublishedTreeView(
                name => PublishedTreeProbe.HashOf(files, name),
                relativeFolder => PublishedTreeProbe.List(dataTreeRoot, relativeFolder),
                name => files.TryGetValue(name.Trim(), out BoardDocument.StoredFileRow? file) ? file.Size : null);
        }

        // The checksum the board lists under this name, lower-case, or null.
        private static string? HashOf(IReadOnlyDictionary<string, BoardDocument.StoredFileRow> files, string name)
        {
            if (!files.TryGetValue(name.Trim(), out BoardDocument.StoredFileRow? file))
                return null;

            string sha = (file.Sha256 ?? string.Empty).Trim().ToLowerInvariant();

            return DataTreeLayout.IsSha256(sha) ? sha : null;
        }

        // The view a submission is judged against: its own board's.
        public static PublishedTreeView? For(string? dataTreeRoot, SubmissionManifest manifest)
        {
            ArgumentNullException.ThrowIfNull(manifest);

            return PublishedTreeProbe.For(dataTreeRoot, manifest.BoardId);
        }

        private static IReadOnlyCollection<string>? List(string dataTreeRoot, string relativeFolder)
        {
            string folder;

            if (relativeFolder.Length == 0)
            {
                folder = dataTreeRoot;
            }
            else if (!SubmissionPathRules.TryResolve(dataTreeRoot, relativeFolder, out folder, out _))
            {
                return null;
            }

            try
            {
                if (!Directory.Exists(folder))
                    return null;

                return Directory
                    .EnumerateFileSystemEntries(folder)
                    .Select(Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name))
                    .Select(name => name!)
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return null;
            }
        }
    }
}
