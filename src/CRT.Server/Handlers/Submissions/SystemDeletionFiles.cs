using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The DISK half of deleting a system (owner request, 2026-10-03): what one data tree holds of
    // it, whether that may go, and taking it out. SystemDeletionFlow sequences it; the rules and
    // the sentences are SystemDeletionRules'.
    //
    // *** WHAT GOES: THE SYSTEM'S OWN FOLDER, WHOLE, AND ITS ROW IN THE NEWEST MAIN EXCEL DATA
    // FILE. *** Nothing outside <Manufacturer>/<Hardware>/<Board>/ is ever touched - a shared file
    // the board cited STAYS, and becomes an unused file for Account > Unused files if nothing else
    // cites it (owner decision, 2026-09-27: AutomaticRemovalScope - automatic removal never leaves
    // a system's own folder). Hidden files in the folder go too: the folder is going.
    //
    // *** REFUSED, NOT GUESSED, when the folder cannot simply go *** (Survey's BlockedBecause):
    //
    //   - an OLDER main Excel data file lists the system. It is frozen (DataGenerationRules) and
    //     serves older CRT versions, and a master naming a missing workbook makes DataTreeUsage
    //     incomplete - which would stop every publish's removals and Account > Unused files for good.
    //   - ANOTHER BOARD uses a file inside this system's folder (the cross-board citations
    //     SubmissionRulesShippedDataTests found in the shipped data are real). Worked out with
    //     DataTreeUsage, the one rule for what a tree uses: the tree is previewed with this system's
    //     workbooks citing NOTHING, and a file of the folder still used after that is used by
    //     something else. Its own workbooks, their highlight files, its "KiCad data" folder and "!"
    //     documentation are the system's own and are not asked about.
    //   - a tree that cannot be read completely, a main Excel data file that cannot be read, or a
    //     symbolic link anywhere on the way - the same fail-closed rules every other writer keeps.
    // ###########################################################################################
    public static class SystemDeletionFiles
    {
        // ###########################################################################################
        // What `root` holds of the system. Reads every workbook in the tree when the folder holds
        // files (the "another board" question), so it takes seconds on a real tree. Writes nothing.
        // ###########################################################################################
        public static SystemTreeSurvey Survey(
            string treeName,
            string? root,
            string systemId,
            Func<string, bool>? isLink = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(treeName);
            ArgumentException.ThrowIfNullOrWhiteSpace(systemId);

            Func<string, bool> link = isLink ?? PublishPathSafety.IsLink;

            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
                return SystemTreeSurvey.Blocked(treeName, root, SystemDeletionRules.MissingTreeMessage(treeName));

            string fullRoot = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));

            if (!SubmissionPathRules.TryResolve(fullRoot, systemId, out string folder, out string why))
                return SystemTreeSurvey.Blocked(treeName, fullRoot, $"[{systemId}] cannot be a folder in the {treeName} data: {why}");

            // ---- 1. The drop-down lists: the newest may be edited, an older one may not -----------
            string? newest = MasterListing.NewestMasterPath(fullRoot);
            string? newestName = newest is null ? null : Path.GetFileName(newest);
            bool listed = false;
            var listedWorkbooks = new List<string>();
            IReadOnlyList<string> masters;

            try
            {
                masters = SystemDeletionFiles.Masters(fullRoot);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.UnreadableFolderMessage(treeName, ex.Message));
            }

            foreach (string master in masters)
            {
                string name = Path.GetFileName(master);

                if (!MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out string readWhy))
                    return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.UnreadableMasterMessage(treeName, name, readWhy));

                List<MasterListingRow> ours = rows
                    .Where(row => string.Equals(row.SystemId, systemId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (ours.Count == 0)
                    continue;

                if (!string.Equals(name, newestName, StringComparison.Ordinal))
                    return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.OlderListingMessage(treeName, name));

                listed = true;
                listedWorkbooks.AddRange(ours.Select(row => DataTreeUsage.Normalise(row.ExcelDataFile)));
            }

            // ---- 2. The folder, every file in it, no link anywhere -------------------------------
            if (PublishPathSafety.FindLinkOnPath(fullRoot, folder, link) is string folderLink)
                return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.LinkMessage(treeName, folderLink));

            var files = new List<string>();

            try
            {
                if (Directory.Exists(folder) && SystemDeletionFiles.Walk(folder, fullRoot, files, link) is string innerLink)
                    return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.LinkMessage(treeName, innerLink));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return SystemTreeSurvey.Blocked(treeName, fullRoot, SystemDeletionRules.UnreadableFolderMessage(treeName, ex.Message));
            }

            files.Sort(StringComparer.Ordinal);

            // ---- 3. Nothing else may use a file in it ---------------------------------------------
            if (files.Count > 0 &&
                SystemDeletionFiles.UsedElsewhere(treeName, fullRoot, systemId, files, listedWorkbooks) is string usedElsewhere)
            {
                return SystemTreeSurvey.Blocked(treeName, fullRoot, usedElsewhere);
            }

            return new SystemTreeSurvey(treeName, fullRoot, folder, files, listed, listed ? newest : null, BlockedBecause: null);
        }

        // ###########################################################################################
        // The folders this process may not write, among those the removal touches - each file's
        // folder, and the root for the main Excel data file's replacement. Asked BEFORE anything is
        // changed in either tree, after the survey's link checks (TreeWriteAccess). `canWrite` is
        // for tests.
        // ###########################################################################################
        public static IReadOnlyList<string> FoldersRefusing(SystemTreeSurvey survey, Func<string, bool>? canWrite = null)
        {
            ArgumentNullException.ThrowIfNull(survey);

            if (string.IsNullOrWhiteSpace(survey.Root))
                return [];

            var paths = new List<string>();

            foreach (string relative in survey.Files)
            {
                if (SubmissionPathRules.TryResolve(survey.Root, relative, out string full, out _))
                    paths.Add(full);
            }

            if (survey.ListedIn is string master)
                paths.Add(master);

            return paths.Count == 0 ? [] : TreeWriteAccess.FoldersRefusing(survey.Root, paths, canWrite);
        }

        // ###########################################################################################
        // Takes the system's row out of the tree's newest main Excel data file. Done when it was not
        // listed. The file is replaced atomically (MasterListing).
        // ###########################################################################################
        public static MasterListingEdit RemoveListing(SystemTreeSurvey survey, string systemId)
        {
            ArgumentNullException.ThrowIfNull(survey);

            return survey.ListedIn is string master
                ? MasterListing.Remove(master, systemId)
                : MasterListingEdit.Done(changed: false);
        }

        // ###########################################################################################
        // Deletes every file the survey found, then every folder that leaves empty - the system's
        // own, and its hardware and manufacturer folders when nothing else is in them, never the
        // root. A file that will not go is logged and counted as left, never thrown: the flow then
        // keeps the database record, so pressing Delete again finishes the job.
        // ###########################################################################################
        public static SystemTreeRemoval RemoveFiles(SystemTreeSurvey survey, ILogger logger, Func<string, bool>? isLink = null)
        {
            ArgumentNullException.ThrowIfNull(survey);
            ArgumentNullException.ThrowIfNull(logger);

            if (string.IsNullOrWhiteSpace(survey.Root) || survey.Folder is null)
                return new SystemTreeRemoval(0, 0);

            Func<string, bool> link = isLink ?? PublishPathSafety.IsLink;
            int removed = 0;
            int left = 0;

            foreach (string relative in survey.Files)
            {
                if (!SubmissionPathRules.TryResolve(survey.Root, relative, out string full, out string why))
                {
                    logger.LogWarning("[{Path}] was not removed from the {Tree} data: {Why}", relative, survey.TreeName, why);
                    left++;
                    continue;
                }

                if (PublishPathSafety.FindLinkOnPath(survey.Root, full, link) is string found)
                {
                    logger.LogWarning("[{Path}] was not removed from the {Tree} data: [{Link}] is a link.", relative, survey.TreeName, found);
                    left++;
                    continue;
                }

                try
                {
                    if (File.Exists(full))
                    {
                        File.Delete(full);
                        BetaRollbackFiles.Forget(full);
                        removed++;
                    }
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    logger.LogWarning(ex, "[{Path}] could not be removed from the {Tree} data.", relative, survey.TreeName);
                    left++;
                }
            }

            SystemDeletionFiles.RemoveEmptyFolders(survey.Root, survey.Folder, logger);

            return new SystemTreeRemoval(removed, left);
        }

        // Every master workbook at the top of the tree, of every generation.
        private static IReadOnlyList<string> Masters(string root) =>
            Directory
                .EnumerateFiles(root, DataGenerationRules.MasterStem + "*" + DataGenerationRules.WorkbookExtension)
                .Where(path => DataTreeUsage.IsMasterFileName(Path.GetFileName(path)))
                .Order(StringComparer.Ordinal)
                .ToList();

        // ###########################################################################################
        // Every file under `directory`, data-root-relative with "/", or the first link met - which
        // is never followed. One level at a time rather than a recursive enumeration, so a link is
        // SEEN rather than quietly walked through or skipped.
        // ###########################################################################################
        private static string? Walk(string directory, string root, List<string> files, Func<string, bool> isLink)
        {
            foreach (string entry in Directory.EnumerateFileSystemEntries(directory))
            {
                if (isLink(entry))
                    return entry;

                if (Directory.Exists(entry))
                {
                    if (SystemDeletionFiles.Walk(entry, root, files, isLink) is string found)
                        return found;
                }
                else
                {
                    files.Add(Path.GetRelativePath(root, entry).Replace(Path.DirectorySeparatorChar, '/'));
                }
            }

            return null;
        }

        // ###########################################################################################
        // The refusal for files of this folder that something else uses, or null - see the class
        // header. Its workbooks (in the folder, or listed by the newest master) are previewed as
        // citing nothing; a listed workbook the folder lacks would otherwise make the tree read as
        // incomplete.
        // ###########################################################################################
        private static string? UsedElsewhere(
            string treeName,
            string root,
            string systemId,
            IReadOnlyList<string> files,
            IReadOnlyList<string> listedWorkbooks)
        {
            var own = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string workbook in files.Where(DataTreeUsage.IsBoardFolderWorkbook).Concat(listedWorkbooks))
            {
                if (workbook.Length > 0)
                    own.Add(workbook);
            }

            var overrides = own.ToDictionary(
                workbook => workbook,
                IReadOnlyCollection<string> (_) => [],
                StringComparer.OrdinalIgnoreCase);

            DataTreeUsageResult usage;

            try
            {
                usage = DataTreeUsage.Compute(root, overrides);
            }
            catch (Exception ex)
            {
                return SystemDeletionRules.IncompleteTreeMessage(treeName, [ex.Message]);
            }

            if (!usage.IsComplete)
                return SystemDeletionRules.IncompleteTreeMessage(treeName, usage.Problems);

            var sidecars = new HashSet<string>(
                own.Select(workbook => DataTreeUsage.Normalise(BoardComponentHighlightStorage.GetJsonPath(workbook))),
                StringComparer.OrdinalIgnoreCase);

            string kicad = $"{systemId}/{DataTreeUsage.KiCadFolderName}/";

            List<string> usedElsewhere = files
                .Where(path => !own.Contains(path) && !sidecars.Contains(path))
                .Where(path => !path.StartsWith(kicad, StringComparison.OrdinalIgnoreCase))
                .Where(path => !Path.GetFileName(path).StartsWith(DataTreeUsage.DocumentationPrefix, StringComparison.Ordinal))
                .Where(usage.IsUsed)
                .ToList();

            return usedElsewhere.Count == 0 ? null : SystemDeletionRules.UsedElsewhereMessage(treeName, usedElsewhere);
        }

        // Bottom up: every empty folder inside the system's folder, the folder itself, then each
        // parent while it is empty - never the root. Best effort: an empty folder left behind holds
        // no workbook, so it is not a system and is not downloaded.
        // A system folder that was never there leaves its parents alone - an empty hardware folder
        // is somebody else's business.
        private static void RemoveEmptyFolders(string root, string folder, ILogger logger)
        {
            try
            {
                if (!Directory.Exists(folder))
                    return;

                foreach (string inner in Directory
                    .EnumerateDirectories(folder, "*", SearchOption.AllDirectories)
                    .OrderByDescending(path => path.Length))
                {
                    if (!Directory.EnumerateFileSystemEntries(inner).Any())
                        Directory.Delete(inner);
                }

                string? current = folder;

                while (!string.IsNullOrEmpty(current) &&
                       current.Length > root.Length &&
                       current.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
                {
                    if (Directory.Exists(current))
                    {
                        if (Directory.EnumerateFileSystemEntries(current).Any())
                            return;

                        Directory.Delete(current);
                    }

                    current = Path.GetDirectoryName(current);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                logger.LogWarning(ex, "An emptied folder could not be removed: {Folder}", folder);
            }
        }
    }

    // ###########################################################################################
    // What one tree holds of a system. Root is the full tree root; Folder the system's folder (null
    // when it could not be resolved); Files every file in it, data-root-relative, ordinal order;
    // ListedIn the newest main Excel data file when it lists the system. BlockedBecause says why
    // nothing may be deleted, in SystemDeletionRules' words.
    // ###########################################################################################
    public sealed record SystemTreeSurvey(
        string TreeName,
        string? Root,
        string? Folder,
        IReadOnlyList<string> Files,
        bool IsListed,
        string? ListedIn,
        string? BlockedBecause)
    {
        public bool HoldsAnything => this.Files.Count > 0 || this.IsListed;

        public static SystemTreeSurvey Blocked(string treeName, string? root, string why) =>
            new(treeName, root, null, [], false, null, why);
    }

    public sealed record SystemTreeRemoval(int Removed, int Left);
}
