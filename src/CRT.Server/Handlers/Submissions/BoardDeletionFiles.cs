using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The DISK half of deleting a board (owner request, 2026-10-03): what one data tree holds of
    // it, whether that may go, and taking it out. BoardDeletionFlow sequences it; the rules and
    // the sentences are BoardDeletionRules'.
    //
    // *** WHAT GOES: THE BOARD'S OWN FOLDER, WHOLE, AND ITS ROW IN THE NEWEST MAIN EXCEL DATA
    // FILE. *** Nothing outside <Manufacturer>/<Hardware>/<Board>/ is ever touched - a shared file
    // the board cited STAYS, and becomes an unused file for Account > Unused files if nothing else
    // cites it (owner decision, 2026-09-27: AutomaticRemovalScope - automatic removal never leaves
    // a board's own folder). Hidden files in the folder go too: the folder is going.
    //
    // *** REFUSED, NOT GUESSED, when the folder cannot simply go *** (Survey's BlockedBecause):
    //
    //   - an OLDER main Excel data file lists the board. It is frozen (DataGenerationRules) and
    //     serves older CRT versions, and a master naming a missing workbook makes DataTreeUsage
    //     incomplete - which would stop every publish's removals and Account > Unused files for good.
    //   - ANOTHER BOARD uses a file inside this board's folder (the cross-board citations
    //     SubmissionRulesShippedDataTests found in the shipped data are real). Worked out with
    //     DataTreeUsage, the one rule for what a tree uses: the tree is previewed with this board's
    //     workbooks citing NOTHING, and a file of the folder still used after that is used by
    //     something else. Its own workbooks, their highlight files, its "KiCad data" folder and "!"
    //     documentation are the board's own and are not asked about.
    //   - a tree that cannot be read completely, a main Excel data file that cannot be read, or a
    //     symbolic link anywhere on the way - the same fail-closed rules every other writer keeps.
    // ###########################################################################################
    public static class BoardDeletionFiles
    {
        // ###########################################################################################
        // What `root` holds of the board. Reads every workbook in the tree when the folder holds
        // files (the "another board" question), so it takes seconds on a real tree. Writes nothing.
        // ###########################################################################################
        public static BoardTreeSurvey Survey(
            string treeName,
            string? root,
            string boardId,
            Func<string, bool>? isLink = null)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(treeName);
            ArgumentException.ThrowIfNullOrWhiteSpace(boardId);

            Func<string, bool> link = isLink ?? PublishPathSafety.IsLink;

            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
                return BoardTreeSurvey.Blocked(treeName, root, BoardDeletionRules.MissingTreeMessage(treeName));

            string fullRoot = Path.TrimEndingDirectorySeparator(Path.GetFullPath(root));

            if (!SubmissionPathRules.TryResolve(fullRoot, boardId, out string folder, out string why))
                return BoardTreeSurvey.Blocked(treeName, fullRoot, $"[{boardId}] cannot be a folder in the {treeName} data: {why}");

            // ---- 1. The drop-down lists: the newest may be edited, an older one may not -----------
            string? newest = MasterListing.NewestMasterPath(fullRoot);
            string? newestName = newest is null ? null : Path.GetFileName(newest);
            bool listed = false;
            var listedWorkbooks = new List<string>();
            IReadOnlyList<string> masters;

            try
            {
                masters = BoardDeletionFiles.Masters(fullRoot);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.UnreadableFolderMessage(treeName, ex.Message));
            }

            foreach (string master in masters)
            {
                string name = Path.GetFileName(master);

                if (!MasterListing.TryRead(master, out IReadOnlyList<MasterListingRow> rows, out string readWhy))
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.UnreadableMasterMessage(treeName, name, readWhy));

                List<MasterListingRow> ours = rows
                    .Where(row => string.Equals(row.BoardId, boardId, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (ours.Count == 0)
                    continue;

                if (!string.Equals(name, newestName, StringComparison.Ordinal))
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.OlderListingMessage(treeName, name));

                listed = true;
                listedWorkbooks.AddRange(ours.Select(row => DataTreeUsage.Normalise(row.ExcelDataFile)));
            }

            // ---- 2. The folder, every file in it, no link anywhere -------------------------------
            if (PublishPathSafety.FindLinkOnPath(fullRoot, folder, link) is string folderLink)
                return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.LinkMessage(treeName, folderLink));

            var files = new List<string>();

            try
            {
                if (Directory.Exists(folder) && BoardDeletionFiles.Walk(folder, fullRoot, files, link) is string innerLink)
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.LinkMessage(treeName, innerLink));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.UnreadableFolderMessage(treeName, ex.Message));
            }

            files.Sort(StringComparer.Ordinal);

            // ---- 3. Nothing else may use a file in it ---------------------------------------------
            if (files.Count > 0 &&
                BoardDeletionFiles.UsedElsewhere(treeName, fullRoot, boardId, files, listedWorkbooks) is string usedElsewhere)
            {
                return BoardTreeSurvey.Blocked(treeName, fullRoot, usedElsewhere);
            }

            return new BoardTreeSurvey(treeName, fullRoot, folder, files, listed, listed ? newest : null, BlockedBecause: null);
        }

        // ###########################################################################################
        // The folders this process may not write, among those the removal touches - each file's
        // folder, and the root for the main Excel data file's replacement. Asked BEFORE anything is
        // changed in either tree, after the survey's link checks (TreeWriteAccess). `canWrite` is
        // for tests.
        // ###########################################################################################
        public static IReadOnlyList<string> FoldersRefusing(BoardTreeSurvey survey, Func<string, bool>? canWrite = null)
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
        // Takes the board's row out of the tree's newest main Excel data file. Done when it was not
        // listed. The file is replaced atomically (MasterListing).
        // ###########################################################################################
        public static MasterListingEdit RemoveListing(BoardTreeSurvey survey, string boardId)
        {
            ArgumentNullException.ThrowIfNull(survey);

            return survey.ListedIn is string master
                ? MasterListing.Remove(master, boardId)
                : MasterListingEdit.Done(changed: false);
        }

        // ###########################################################################################
        // Deletes every file the survey found, then every folder that leaves empty - the board's
        // own, and its hardware and manufacturer folders when nothing else is in them, never the
        // root. A file that will not go is logged and counted as left, never thrown: the flow then
        // keeps the database record, so pressing Delete again finishes the job.
        // ###########################################################################################
        public static BoardTreeRemoval RemoveFiles(BoardTreeSurvey survey, ILogger logger, Func<string, bool>? isLink = null)
        {
            ArgumentNullException.ThrowIfNull(survey);
            ArgumentNullException.ThrowIfNull(logger);

            if (string.IsNullOrWhiteSpace(survey.Root) || survey.Folder is null)
                return new BoardTreeRemoval(0, 0);

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

            BoardDeletionFiles.RemoveEmptyFolders(survey.Root, survey.Folder, logger);

            return new BoardTreeRemoval(removed, left);
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
                    if (BoardDeletionFiles.Walk(entry, root, files, isLink) is string found)
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
            string boardId,
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
                return BoardDeletionRules.IncompleteTreeMessage(treeName, [ex.Message]);
            }

            if (!usage.IsComplete)
                return BoardDeletionRules.IncompleteTreeMessage(treeName, usage.Problems);

            var sidecars = new HashSet<string>(
                own.Select(workbook => DataTreeUsage.Normalise(BoardComponentHighlightStorage.GetJsonPath(workbook))),
                StringComparer.OrdinalIgnoreCase);

            string kicad = $"{boardId}/{DataTreeUsage.KiCadFolderName}/";

            List<string> usedElsewhere = files
                .Where(path => !own.Contains(path) && !sidecars.Contains(path))
                .Where(path => !path.StartsWith(kicad, StringComparison.OrdinalIgnoreCase))
                .Where(path => !Path.GetFileName(path).StartsWith(DataTreeUsage.DocumentationPrefix, StringComparison.Ordinal))
                .Where(usage.IsUsed)
                .ToList();

            return usedElsewhere.Count == 0 ? null : BoardDeletionRules.UsedElsewhereMessage(treeName, usedElsewhere);
        }

        // Bottom up: every empty folder inside the board's folder, the folder itself, then each
        // parent while it is empty - never the root. Best effort: an empty folder left behind holds
        // no workbook, so it is not a board and is not downloaded.
        // A board folder that was never there leaves its parents alone - an empty hardware folder
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
    // What one tree holds of a board. Root is the full tree root; Folder the board's folder (null
    // when it could not be resolved); Files every file in it, data-root-relative, ordinal order;
    // ListedIn the newest main Excel data file when it lists the board. BlockedBecause says why
    // nothing may be deleted, in BoardDeletionRules' words.
    // ###########################################################################################
    public sealed record BoardTreeSurvey(
        string TreeName,
        string? Root,
        string? Folder,
        IReadOnlyList<string> Files,
        bool IsListed,
        string? ListedIn,
        string? BlockedBecause)
    {
        public bool HoldsAnything => this.Files.Count > 0 || this.IsListed;

        public static BoardTreeSurvey Blocked(string treeName, string? root, string why) =>
            new(treeName, root, null, [], false, null, why);
    }

    public sealed record BoardTreeRemoval(int Removed, int Left);
}
