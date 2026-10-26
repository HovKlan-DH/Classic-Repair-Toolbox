using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A SUBMISSION'S FILE TREE: the BETA data after approving it, against BETA now (owner request,
    // 2026-09-28) - GET /api/review/submissions/{id}/files. The rule for what each file becomes is
    // CRT.Data's BoardFileEntries.ForApproval; this gathers what only the server can see.
    //
    // *** IN THE CRT 4.0 TREE THE TREE IS THE BOARD'S NAMES (2026-10-25). *** A board lists its
    // files by name, each one's bytes stored under their checksum (DataTreeLayout), so:
    //
    //   - what the submission carries, against what the published board lists under each name
    //     (SubmittedFileFacts, the same facts the submission detail sends);
    //   - what the approval removes, by name (the same FileRemovalPreview the approval is sent back);
    //   - every file the published board lists now, so the tree is the WHOLE board and not only
    //     what moves;
    //   - the board file the approval writes from the table, "board.json", which always changes (it
    //     carries the publish date).
    //
    // Every entry opened from BETA carries where its bytes lie (BoardFileEntry.TreePath) - the
    // public data is laid out by checksum - and the size of what opens from it (TreeFileSizes,
    // 2026-10-04). Asked for on demand rather than sent with every submission detail: it reads the
    // board file and builds the publish plan, and the detail is read on every click in the queue.
    // ###########################################################################################
    public static class SubmissionFileTreeFlow
    {
        public static async Task<IReadOnlyList<BoardFileEntry>> BuildAsync(
            string dataTreeRoot,
            SubmissionManifest manifest,
            PublishedBoardReader publishedBoards,
            DateTimeOffset nowUtc,
            CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(dataTreeRoot);
            ArgumentNullException.ThrowIfNull(manifest);
            ArgumentNullException.ThrowIfNull(publishedBoards);

            string root = Path.GetFullPath(dataTreeRoot);

            BoardData? published = await publishedBoards
                .TryReadAsync(root, manifest, cancellationToken)
                .ConfigureAwait(false);

            IReadOnlyDictionary<string, BoardDocument.StoredFileRow> listed =
                PublishedBoardFiles.ByName(root, manifest.BoardId) ?? new Dictionary<string, BoardDocument.StoredFileRow>();

            FileRemovalPreview removals = ApprovePublishFlow.PreviewRemovals(root, manifest, published, nowUtc);
            PublishPlanResult plan = ApprovePublishFlow.BuildPlan(root, manifest, nowUtc);

            // A plan the approval would refuse writes nothing - so nothing is shown as written.
            GeneratedFile? boardFile = plan.IsPlanned
                ? new GeneratedFile(DataTreeLayout.BoardFileName, File.Exists(plan.Plan!.BoardFilePath), Changes: true)
                : null;

            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                SubmittedFileFacts.Build(manifest, SubmissionFileTreeFlow.HashesOf(listed)),
                removals.Files.Select(removals.NameOf).ToList(),
                listed.Keys.ToList(),
                boardFile,
                sidecar: null);

            string boardFileRelative = DataTreeLayout.BoardFilePath(manifest.BoardId);

            // Where each entry's bytes lie in BETA: the board file at its place, a listed name at its
            // stored file - a removed name too, since its bytes are still there until the approval.
            entries = BoardFileEntries.WithTreePaths(entries, entry =>
            {
                if (entry.WrittenOnApproval)
                    return boardFileRelative;

                if (entry.OpenFrom == BoardFileSource.Beta && listed.TryGetValue(entry.Path, out BoardDocument.StoredFileRow? file))
                    return PublishedBoardFiles.StoredRelative(file);

                return null;
            });

            return TreeFileSizes.Attach(entries, root, submitted: TreeFileSizes.OfSubmission(manifest));
        }

        // The published board's checksum under each name - what a submitted file is compared with.
        internal static IReadOnlyDictionary<string, string> HashesOf(IReadOnlyDictionary<string, BoardDocument.StoredFileRow> listed)
        {
            var hashes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach ((string name, BoardDocument.StoredFileRow file) in listed)
            {
                if (DataTreeLayout.IsSha256(file.Sha256?.ToLowerInvariant()))
                    hashes[name] = file.Sha256!.ToLowerInvariant();
            }

            return hashes;
        }
    }
}
