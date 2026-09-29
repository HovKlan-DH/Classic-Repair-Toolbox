using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A SUBMISSION'S FILE TREE: the BETA data after approving it, against BETA now (owner request,
    // 2026-09-28) - GET /api/review/submissions/{id}/files. The rule for what each file becomes is
    // CRT.Data's SystemFileEntries.ForApproval; this gathers what only the server can see:
    //
    //   - what the submission carries, against BETA's bytes at each path (SubmittedFileFacts, the
    //     same facts the submission detail sends);
    //   - what the approval removes (the same FileRemovalPreview the approval is sent back);
    //   - every file in the system's own BETA folder now, so the tree is the WHOLE system and not
    //     only what moves - an older generation's workbook, KiCad data, anything else there;
    //   - the two files the approval writes from the table: the workbook, which always changes (it
    //     carries the publish date), and the highlight file, which changes only when its content
    //     does (BoardSidecarWriter.HoldsTheSameContent - the very check the publish makes).
    //
    // Asked for on demand rather than sent with every submission detail: it walks the board's
    // folder and builds the publish plan, which hashes files, and the detail is read on every
    // click in the queue.
    // ###########################################################################################
    public static class SubmissionFileTreeFlow
    {
        public static async Task<IReadOnlyList<SystemFileEntry>> BuildAsync(
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

            IReadOnlyDictionary<string, string> hashes = await PublishedFileHashes
                .ComputeAsync(root, manifest.Files.Select(file => file.Path).Distinct(StringComparer.Ordinal).ToList(), cancellationToken)
                .ConfigureAwait(false);

            FileRemovalPreview removals = ApprovePublishFlow.PreviewRemovals(root, manifest, published, nowUtc);
            PublishPlanResult plan = ApprovePublishFlow.BuildPlan(root, manifest, published, nowUtc);

            GeneratedFile? workbook = null;
            GeneratedFile? sidecar = null;

            // A plan the approval would refuse writes nothing - so nothing is shown as written.
            if (plan.IsPlanned)
            {
                PublishPlanDetail detail = plan.Plan!;
                BoardData board = PublishMerge.Build(manifest, published);
                bool sidecarExists = File.Exists(detail.SidecarPath);

                workbook = new GeneratedFile(
                    SubmissionFileTreeFlow.Relative(root, detail.WorkbookPath),
                    File.Exists(detail.WorkbookPath),
                    Changes: true);

                sidecar = new GeneratedFile(
                    SubmissionFileTreeFlow.Relative(root, detail.SidecarPath),
                    sidecarExists,
                    Changes: !sidecarExists || !BoardSidecarWriter.HoldsTheSameContent(
                        detail.WorkbookPath, board.ComponentHighlights, PublishMerge.CalibrationsOf(manifest)));
            }

            return SystemFileEntries.ForApproval(
                SubmittedFileFacts.Build(manifest, hashes),
                removals.Files,
                SubmissionFileTreeFlow.OwnFiles(root, manifest),
                workbook,
                sidecar);
        }

        // ###########################################################################################
        // Every file in the system's own BETA folder, data-root-relative with forward slashes. None
        // for a new system. Hidden paths (a publish's temporaries, anything dot-named) and the
        // retired system.json are left out, as the production plan leaves them out - neither is
        // ever published.
        // ###########################################################################################
        internal static IReadOnlyList<string> OwnFiles(string root, SubmissionManifest manifest)
        {
            PublishedBoardLocation location = PublishedBoardLocator.Locate(root, manifest);

            string folder = location.Exists
                ? location.SystemFolder
                : Path.Combine(root, manifest.Manufacturer.Trim(), manifest.Hardware.Trim(), manifest.Board.Trim());

            if (!Directory.Exists(folder))
                return [];

            string retired = SubmissionFileTreeFlow.Relative(root, Path.Combine(folder, SystemDescriptorStore.FileName));

            return Directory
                .EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                .Select(path => SubmissionFileTreeFlow.Relative(root, path))
                .Where(path => !path.Split('/').Any(segment => segment.StartsWith('.')))
                .Where(path => !string.Equals(path, retired, StringComparison.Ordinal))
                .Order(StringComparer.Ordinal)
                .ToList();
        }

        private static string Relative(string root, string path) =>
            Path.GetRelativePath(root, Path.GetFullPath(path)).Replace(Path.DirectorySeparatorChar, '/');
    }
}
