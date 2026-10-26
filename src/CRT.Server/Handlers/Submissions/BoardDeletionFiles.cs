using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // The DISK half of deleting a board (owner request, 2026-10-03): what one data tree holds of
    // it, whether that may go, and taking it out. BoardDeletionFlow sequences it; the rules and
    // the sentences are BoardDeletionRules'.
    //
    // *** WHAT GOES, IN THE CRT 4.0 TREE (2026-10-25): THE BOARD'S OWN FOLDER, WHOLE, THE STORED
    // FILES ONLY IT LISTS, AND ITS ROW IN THE MAIN FILE. *** A board's folder holds its board file;
    // its files' bytes are stored once under their checksum ("Files/<ab>/<sha256>.<ext>") for every
    // board naming them - so bytes another board lists STAY, and the rest go with the board
    // (AutomaticRemovalScope - a board's own files). Hidden files in the folder go too: the folder is
    // going. The Excel generations older CRTs read are trees of their own, never touched here.
    //
    // *** REFUSED, NOT GUESSED, when it cannot simply go *** (Survey's BlockedBecause): a tree that
    // cannot be read completely, a main file or board file that cannot be read, or a symbolic link
    // anywhere on the way - the same fail-closed rules every other writer keeps.
    // ###########################################################################################
    public static class BoardDeletionFiles
    {
        // An exception's own text names the server's paths: it goes to the log, and the refusal the
        // administrator reads says where to look instead (code review, 2026-10-13).
        internal static string LoggedOnly(Exception ex)
        {
            CrtLog.Warning($"Reading a board's folder for its deletion failed - [{ex}]");
            return "the server's log says why";
        }

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

            // ---- 1. The drop-down lists: the tree's main file ------------------------------------
            //
            // The CRT 4.0 tree has ONE main file (DataTreeLayout.MainFileName) - the Excel generations
            // older CRTs read are trees of their own, never touched here.
            string? newest = MasterListing.NewestMasterPath(fullRoot);
            bool listed = false;

            if (newest is not null)
            {
                if (!MasterListing.TryRead(newest, out IReadOnlyList<MasterListingRow> rows, out string readWhy))
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.UnreadableMasterMessage(treeName, Path.GetFileName(newest), readWhy));

                listed = MasterListing.IndexOfBoard(rows, boardId) >= 0;
            }

            // ---- 2. The folder, every file in it, no link anywhere -------------------------------
            if (PublishPathSafety.FindLinkOnPath(fullRoot, folder, link) is string folderLink)
                return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.LinkMessage(treeName, fullRoot, folderLink));

            var files = new List<string>();

            try
            {
                if (Directory.Exists(folder) && BoardDeletionFiles.Walk(folder, fullRoot, files, link) is string innerLink)
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.LinkMessage(treeName, fullRoot, innerLink));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.UnreadableFolderMessage(treeName, BoardDeletionFiles.LoggedOnly(ex)));
            }

            // ---- 3. The stored files only this board lists -----------------------------------------
            //
            // In the CRT 4.0 tree a board's files' bytes are stored once under their checksum, for
            // every board naming them. Those no OTHER board lists go with the board - left, they
            // would be files every CRT downloads and nothing shows; those another board lists stay.
            IReadOnlyList<string>? stored = null;
            string? incomplete = null;

            if (files.Count > 0)
                stored = BoardDeletionFiles.StoredOnlyHere(treeName, fullRoot, boardId, out incomplete);

            if (incomplete is not null)
                return BoardTreeSurvey.Blocked(treeName, fullRoot, incomplete);

            foreach (string path in stored ?? [])
            {
                if (!SubmissionPathRules.TryResolve(fullRoot, path, out string full, out _))
                    continue;

                if (PublishPathSafety.FindLinkOnPath(fullRoot, full, link) is string storedLink)
                    return BoardTreeSurvey.Blocked(treeName, fullRoot, BoardDeletionRules.LinkMessage(treeName, fullRoot, storedLink));

                files.Add(path);
            }

            files.Sort(StringComparer.Ordinal);

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
        // The stored files the board lists that no OTHER board of the tree lists - what goes with it.
        // Worked out with DataTreeUsage, the one rule for what a tree uses: the tree is previewed
        // with this board's file listing NOTHING, and its stored files then unused are its alone.
        // Null, with the refusal in `incomplete`, when the tree cannot be read completely - failing
        // closed, as every writer does; an empty list when the board file lists nothing.
        // ###########################################################################################
        private static IReadOnlyList<string>? StoredOnlyHere(string treeName, string root, string boardId, out string? incomplete)
        {
            incomplete = null;

            string boardFile = DataTreeLayout.BoardFilePath(boardId);
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(root, boardId);

            if (!location.Exists)
                return [];

            BoardDocument? document = BoardFiles.ReadDocument(location.WorkbookPath);

            if (document is null)
            {
                incomplete = BoardDeletionRules.IncompleteTreeMessage(treeName, [$"The board file [{boardFile}] could not be read."]);
                return null;
            }

            List<string> listed = document.Files
                .Select(PublishedBoardFiles.StoredRelative)
                .Where(path => path.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (listed.Count == 0)
                return [];

            DataTreeUsageResult usage;

            try
            {
                usage = DataTreeUsage.Compute(root, new Dictionary<string, IReadOnlyCollection<string>> { [boardFile] = [] });
            }
            catch (Exception ex)
            {
                incomplete = BoardDeletionRules.IncompleteTreeMessage(treeName, [BoardDeletionFiles.LoggedOnly(ex)]);
                return null;
            }

            if (!usage.IsComplete)
            {
                incomplete = BoardDeletionRules.IncompleteTreeMessage(treeName, usage.Problems);
                return null;
            }

            return usage.RemovableFrom(listed);
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
