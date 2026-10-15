using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A SUBMISSION'S FILE TREE: the BETA data after approving it, against BETA now (owner request,
    // 2026-09-28) - GET /api/review/submissions/{id}/files. The rule for what each file becomes is
    // CRT.Data's BoardFileEntries.ForApproval; this gathers what only the server can see:
    //
    //   - what the submission carries, against BETA's bytes at each path (SubmittedFileFacts, the
    //     same facts the submission detail sends);
    //   - what the approval removes (the same FileRemovalPreview the approval is sent back);
    //   - every file in the board's own BETA folder now, so the tree is the WHOLE board and not
    //     only what moves - an older generation's workbook, KiCad data, anything else there;
    //   - the two files the approval writes from the table: the workbook, which always changes (it
    //     carries the publish date), and the highlight file, which changes only when its content
    //     does (BoardSidecarWriter.HoldsTheSameContent - the very check the publish makes).
    //
    // Asked for on demand rather than sent with every submission detail: it walks the board's
    // folder and builds the publish plan, which hashes files, and the detail is read on every
    // click in the queue. Each file carries the size of what opens from it - BETA's file, or the
    // submission's upload (TreeFileSizes, 2026-10-04).
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
                        detail.WorkbookPath,
                        board.ComponentHighlights,
                        PublishMerge.CalibrationsOf(manifest),
                        ApprovePublishFlow.RenameLogFor(detail, published, board)));
            }

            IReadOnlyList<BoardFileEntry> entries = BoardFileEntries.ForApproval(
                SubmittedFileFacts.Build(manifest, hashes),
                removals.Files,
                SubmissionFileTreeFlow.OwnFiles(root, manifest),
                workbook,
                sidecar);

            entries = SubmissionFileTreeFlow.WithAlsoUsedBy(root, manifest, entries, ApprovePublishFlow.PreviewReads);

            return TreeFileSizes.Attach(entries, root, submitted: TreeFileSizes.OfSubmission(manifest));
        }

        // ###########################################################################################
        // *** WHO ELSE SEES A REPLACED SHARED FILE (code review, 2026-10-13). *** Replacing a shared
        // file needs two approvals because every board citing it gets the new bytes - and the two
        // approvers were never told WHICH boards. Each shared file this approval replaces carries the
        // other boards whose workbooks cite it (BoardFileEntry.AlsoUsedBy), "Manufacturer / Hardware /
        // Board", sorted. Read through `cache`, so the tree's workbooks are parsed once per version,
        // not once per click - and only when the submission replaces a shared file at all.
        // ###########################################################################################
        internal static IReadOnlyList<BoardFileEntry> WithAlsoUsedBy(
            string root,
            SubmissionManifest manifest,
            IReadOnlyList<BoardFileEntry> entries,
            WorkbookReadCache cache)
        {
            var replaced = new HashSet<string>(
                entries
                    .Where(entry => entry.Change == BoardFileChange.Changed &&
                                    SubmissionFileScopes.Classify(manifest, entry.Path) is SubmissionFileScope.ManufacturerShared or SubmissionFileScope.GenericShared)
                    .Select(entry => entry.Path),
                StringComparer.OrdinalIgnoreCase);

            if (replaced.Count == 0)
                return entries;

            IReadOnlyDictionary<string, IReadOnlyList<string>> citers =
                SubmissionFileTreeFlow.BoardsCiting(root, replaced, manifest.BoardId, cache);

            return entries
                .Select(entry => replaced.Contains(entry.Path) && citers.TryGetValue(entry.Path, out IReadOnlyList<string>? boards)
                    ? entry with { AlsoUsedBy = boards }
                    : entry)
                .ToList();
        }

        // ###########################################################################################
        // For each of `paths` (data-root-relative, compared ignoring case as every client does), the
        // boards other than `ownBoardId` whose board-folder workbooks - any generation - cite it. A
        // workbook that cannot be read names nobody; this is shown, never acted on.
        // ###########################################################################################
        internal static IReadOnlyDictionary<string, IReadOnlyList<string>> BoardsCiting(
            string root,
            IReadOnlyCollection<string> paths,
            string ownBoardId,
            WorkbookReadCache cache)
        {
            var found = new Dictionary<string, SortedSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (string path in paths)
                found[path] = new SortedSet<string>(StringComparer.Ordinal);

            foreach (string workbook in SubmissionFileTreeFlow.BoardFolderWorkbooks(root))
            {
                string relative = SubmissionFileTreeFlow.Relative(root, workbook);
                string[] segments = relative.Split('/');
                string boardId = string.Join('/', segments.Take(3));

                if (string.Equals(boardId, ownBoardId, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (!cache.TryGetCitations(workbook, out IReadOnlyCollection<string> cites))
                    continue;

                foreach (string cited in cites)
                {
                    string normalised = (cited ?? string.Empty).Trim().Replace('\\', '/').TrimStart('/');

                    if (found.TryGetValue(normalised, out SortedSet<string>? boards))
                        boards.Add(string.Join(" / ", segments.Take(3)));
                }
            }

            return found
                .Where(pair => pair.Value.Count > 0)
                .ToDictionary(pair => pair.Key, pair => (IReadOnlyList<string>)pair.Value.ToList(), StringComparer.OrdinalIgnoreCase);
        }

        // Every workbook at the top of a <Manufacturer>/<Hardware>/<Board>/ folder - DataTreeUsage's
        // own rule for what a board workbook is - found three levels down without walking every file.
        private static IEnumerable<string> BoardFolderWorkbooks(string root)
        {
            if (!Directory.Exists(root))
                yield break;

            var options = new EnumerationOptions { AttributesToSkip = FileAttributes.ReparsePoint, IgnoreInaccessible = true };

            foreach (string manufacturer in Directory.EnumerateDirectories(root, "*", options))
            foreach (string hardware in Directory.EnumerateDirectories(manufacturer, "*", options))
            foreach (string board in Directory.EnumerateDirectories(hardware, "*", options))
            foreach (string file in Directory.EnumerateFiles(board, "*" + DataGenerationRules.WorkbookExtension, options))
            {
                if (DataTreeUsage.IsBoardFolderWorkbook(SubmissionFileTreeFlow.Relative(root, file)))
                    yield return file;
            }
        }

        // ###########################################################################################
        // Every file in the board's own BETA folder, data-root-relative with forward slashes. None
        // for a new board. Hidden paths (a publish's temporaries, anything dot-named) and the
        // retired system.json are left out, as the production plan leaves them out - neither is
        // ever published.
        // ###########################################################################################
        internal static IReadOnlyList<string> OwnFiles(string root, SubmissionManifest manifest)
        {
            PublishedBoardLocation location = PublishedBoardLocator.Locate(root, manifest);

            string folder = location.Exists
                ? location.BoardFolder
                : Path.Combine(root, manifest.Manufacturer.Trim(), manifest.Hardware.Trim(), manifest.Board.Trim());

            return SubmissionFileTreeFlow.FilesIn(root, folder);
        }

        // Every file in a board's folder, by the rule above - also the Boards screen's Files view
        // (BoardFilesFlow, 2026-10-03), so the two trees cannot list one folder differently.
        internal static IReadOnlyList<string> FilesIn(string root, string folder)
        {
            if (!Directory.Exists(folder))
                return [];

            string retired = SubmissionFileTreeFlow.Relative(root, Path.Combine(folder, BoardDescriptorStore.FileName));

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
