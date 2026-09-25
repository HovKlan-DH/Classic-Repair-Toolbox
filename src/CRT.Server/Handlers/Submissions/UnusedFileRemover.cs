using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // REMOVES FILES NOTHING USES from a data tree (maintainer decision, 2026-09-25: "there must be
    // no orphan files").
    //
    // Called in three places, always with a list somebody has SEEN: after a BETA publish and after
    // a production promotion (the FileRemovalPreview the reviewer approved), and from the
    // administrator's "Unused files" screen (the files the administrator chose). It never decides
    // what to remove on its own - it only ever removes FEWER:
    //
    //   - the tree is read again, NOW, and a file that is used after all is kept. Between the list
    //     being shown and this running, another publish may have started citing it;
    //   - an incomplete read removes nothing at all (DataTreeUsage fails closed);
    //   - a path is resolved through SubmissionPathRules, the one containment rule, and nothing is
    //     removed through a symbolic link (PublishPathSafety);
    //   - only files the sync manifest would list are ever candidates, so a half-written ".tmp_"
    //     file beside a publish in progress is never touched.
    //
    // A folder left empty by a removal is removed too, up to (never including) the tree's root.
    //
    // Runs under the PublishLock at every call site, so the tree does not move underneath it.
    // Nothing here throws: a file that cannot be removed is reported as kept and logged, because
    // every caller runs after an operation that has already succeeded.
    // ###########################################################################################
    public static class UnusedFileRemover
    {
        public static UnusedFileRemoval Remove(
            string dataRoot,
            IReadOnlyCollection<string> files,
            ILogger logger,
            Func<string, bool>? isLink = null)
        {
            ArgumentNullException.ThrowIfNull(files);
            ArgumentNullException.ThrowIfNull(logger);

            if (files.Count == 0)
                return UnusedFileRemoval.None;

            DataTreeUsageResult usage;

            try
            {
                usage = DataTreeUsage.Compute(dataRoot);
            }
            catch (Exception ex)
            {
                logger.LogWarning(ex, "Unused files in [{DataRoot}] were not removed: the tree could not be read.", dataRoot);
                return new UnusedFileRemoval([], files.ToList(), "The data could not be read.");
            }

            if (!usage.IsComplete)
            {
                string why = string.Join(" ", usage.Problems);
                logger.LogWarning("Unused files in [{DataRoot}] were not removed: {Why}", dataRoot, why);
                return new UnusedFileRemoval([], files.ToList(), why);
            }

            IReadOnlyList<string> removable = usage.RemovableFrom(files);
            var removed = new List<string>();
            string fullRoot = Path.GetFullPath(dataRoot);

            foreach (string path in removable)
            {
                if (UnusedFileRemover.TryRemove(fullRoot, path, isLink ?? PublishPathSafety.IsLink, logger))
                    removed.Add(path);
            }

            var removedSet = new HashSet<string>(removed, StringComparer.OrdinalIgnoreCase);

            List<string> kept = files
                .Where(path => !removedSet.Contains(path.Trim().Replace('\\', '/')))
                .ToList();

            if (removed.Count > 0)
            {
                logger.LogInformation(
                    "Removed {Count} unused file(s) from [{DataRoot}]: {Files}",
                    removed.Count, dataRoot, string.Join(" | ", removed));
            }

            return new UnusedFileRemoval(removed, kept, null);
        }

        private static bool TryRemove(string fullRoot, string relativePath, Func<string, bool> isLink, ILogger logger)
        {
            if (!SubmissionPathRules.TryResolve(fullRoot, relativePath, out string full, out string why))
            {
                logger.LogWarning("Unused file [{Path}] was not removed: {Why}", relativePath, why);
                return false;
            }

            string? link = PublishPathSafety.FindLinkOnPath(fullRoot, full, isLink);

            if (link is not null)
            {
                logger.LogWarning("Unused file [{Path}] was not removed: [{Link}] is a link.", relativePath, link);
                return false;
            }

            try
            {
                if (!File.Exists(full))
                    return false;

                File.Delete(full);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                logger.LogWarning(ex, "Unused file [{Path}] could not be removed.", relativePath);
                return false;
            }

            UnusedFileRemover.RemoveEmptyFolders(fullRoot, Path.GetDirectoryName(full));
            return true;
        }

        // Walks up from the removed file's folder, removing each that is now empty, and stops at
        // the first that is not - or at the root, which is never removed.
        private static void RemoveEmptyFolders(string fullRoot, string? folder)
        {
            string root = fullRoot.TrimEnd(Path.DirectorySeparatorChar);

            while (!string.IsNullOrEmpty(folder) &&
                   folder.Length > root.Length &&
                   folder.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
            {
                try
                {
                    if (!Directory.Exists(folder) || Directory.EnumerateFileSystemEntries(folder).Any())
                        return;

                    Directory.Delete(folder);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    return;
                }

                folder = Path.GetDirectoryName(folder);
            }
        }
    }

    // ###########################################################################################
    // What was removed, what was asked for and kept (used after all, already gone, or refused),
    // and - when nothing could be removed at all - why.
    // ###########################################################################################
    public sealed record UnusedFileRemoval(
        IReadOnlyList<string> Removed,
        IReadOnlyList<string> Kept,
        string? NotDoneBecause)
    {
        public static UnusedFileRemoval None { get; } = new([], [], null);
    }
}
