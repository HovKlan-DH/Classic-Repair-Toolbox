using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHICH FILES IN A DATA TREE ARE USED - the rule behind "there must be no orphan files"
    // (maintainer decision, 2026-09-25). Every file the rule does not call used is an orphan: a
    // file every CRT downloads and nothing ever shows.
    //
    // A file is USED when it is:
    //   1. a MASTER workbook ("Classic-Repair-Toolbox.xlsx", "Classic-Repair-Toolbox.v2.0.0.xlsx",
    //      at the top of the tree);
    //   2. a BOARD workbook - one a master lists in its "Excel data file" column, OR one found at
    //      the top of a board folder (<Manufacturer>/<Hardware>/<Board>/) - of EVERY generation,
    //      and its .json sidecar;
    //   3. cited by any of those board workbooks (the four sheets BoardDataReader collects);
    //   4. inside a board's "KiCad data" folder;
    //   5. inside a folder CRT reads by NAME rather than through a workbook (FoldersReadByName -
    //      today only the MiniPro IC tests, which a plain "no workbook cites it" count calls unused
    //      and a sweep on it would delete);
    //   6. a documentation file whose name starts with "!" (the "!README.txt" convention).
    //
    // *** A BOARD FOUND IN THE TREE COUNTS EVEN WHEN NO MASTER LISTS IT. *** A new system the
    // server publishes is not added to the master workbook (a known gap, done by hand), so a rule
    // that trusted the masters alone would call the whole new board unused and delete it.
    //
    // *** AN OLDER GENERATION IS USED FOR AS LONG AS IT EXISTS. *** Older workbooks serve older
    // CRT builds and are never written (DataGenerationRules), so what they cite stays.
    //
    // *** FAIL CLOSED. *** A master or board workbook that cannot be read, a master that lists a
    // board workbook the tree does not have, or a folder that cannot be walked, makes the result
    // INCOMPLETE - and an incomplete result removes nothing (RemovableFrom and UnusedFiles are
    // empty). An unreadable workbook cites an UNKNOWN set of files, which is not the same as none.
    //
    // *** MATCHED IGNORING CASE, deliberately the safe way round. *** A row citing "U8.PNG" for a
    // file stored as "U8.png" works on every Windows and macOS client, so that file is used.
    //
    // Only files the sync manifest would list are considered at all (DataChecksumManifest.
    // IsSyncable): a dot-file or a half-written ".tmp_" file is never an orphan to delete.
    //
    // CRT's OWN cleanup on a user's PC (DataManager.DeleteOrphanAndUnusedFilesAsync) is a
    // different question - it keeps whatever the server's manifest lists - and does not use this.
    // ###########################################################################################
    public static class DataTreeUsage
    {
        public const string MasterSheetName = "Hardware & Board";
        public const string ExcelDataFileColumn = "Excel data file";
        public const string KiCadFolderName = "KiCad data";
        public const string DocumentationPrefix = "!";

        // The MiniPro IC tests: CRT reads the catalogue and the vectors by folder
        // (IcTestCatalogue builds its paths from this constant, so there is one spelling).
        public const string MiniProTestsFolder = "Generic shared files/MiniPro/IC tests";

        // Every folder CRT reads by name. A CRT.App test fails if the app reads one not listed.
        public static IReadOnlyList<string> FoldersReadByName { get; } = [DataTreeUsage.MiniProTestsFolder];

        // How far down a master's header row is looked for.
        private const int HeaderSearchRows = 20;

        // ###########################################################################################
        // Works out what the tree at `dataRoot` uses.
        //
        // `citationsAfter` PREVIEWS a publish: board workbook (data-root-relative) -> what it WILL
        // cite, used instead of what the file on disk cites - or as an extra workbook when the
        // publish creates it. That is how a reviewer is shown what a publish would remove BEFORE
        // anything is written, from the same rule that then removes it.
        //
        // `cache` skips re-reading a workbook that has not changed since it was last read - for
        // what is SHOWN only (see WorkbookReadCache). Without one every workbook is read now, which
        // is what anything that deletes must do.
        // ###########################################################################################
        public static DataTreeUsageResult Compute(
            string dataRoot,
            IReadOnlyDictionary<string, IReadOnlyCollection<string>>? citationsAfter = null,
            WorkbookReadCache? cache = null)
        {
            var problems = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var usedFolders = new HashSet<string>(DataTreeUsage.FoldersReadByName, StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(dataRoot) || !Directory.Exists(dataRoot))
            {
                problems.Add("The data tree could not be found.");
                return new DataTreeUsageResult([], used, usedFolders, problems, 0, 0);
            }

            string root = Path.GetFullPath(dataRoot);
            var overrides = new Dictionary<string, IReadOnlyCollection<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, IReadOnlyCollection<string>> pair in citationsAfter ?? new Dictionary<string, IReadOnlyCollection<string>>())
            {
                string key = DataTreeUsage.Normalise(pair.Key);

                if (key.Length > 0)
                    overrides[key] = pair.Value ?? [];
            }

            IReadOnlyList<string> files = DataTreeUsage.ListFiles(root, problems);

            // ---- 1. Masters, and the board workbooks they list ---------------------------------
            var masters = files
                .Where(path => !path.Contains('/') && DataTreeUsage.IsMasterFileName(path))
                .ToList();

            // Listed workbook -> the master that listed it, for the "lists a missing file" message.
            var boardWorkbooks = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

            foreach (string master in masters)
            {
                used.Add(master);

                if (!DataTreeUsage.TryReadListed(Path.Combine(root, master), cache, out IReadOnlyCollection<string> listed, out string why))
                {
                    problems.Add($"The master workbook [{master}] could not be read: {why}");
                    continue;
                }

                foreach (string workbook in listed)
                    boardWorkbooks.TryAdd(workbook, master);
            }

            // ---- 2. Board workbooks found in board folders, whoever lists them -----------------
            foreach (string path in files.Where(DataTreeUsage.IsBoardFolderWorkbook))
                boardWorkbooks.TryAdd(path, null);

            // ...and the ones a previewed publish creates.
            foreach (string path in overrides.Keys)
                boardWorkbooks.TryAdd(path, null);

            var existing = new HashSet<string>(files, StringComparer.OrdinalIgnoreCase);

            foreach ((string workbook, string? listedBy) in boardWorkbooks)
            {
                used.Add(workbook);
                used.Add(DataTreeUsage.Normalise(BoardComponentHighlightStorage.GetJsonPath(workbook)));

                string folder = DataTreeUsage.FolderOf(workbook);
                usedFolders.Add(folder.Length == 0 ? DataTreeUsage.KiCadFolderName : $"{folder}/{DataTreeUsage.KiCadFolderName}");

                // ---- 3. What it cites ------------------------------------------------------------
                if (overrides.TryGetValue(workbook, out IReadOnlyCollection<string>? after))
                {
                    foreach (string cited in after)
                        DataTreeUsage.AddNormalised(used, cited);

                    continue;
                }

                if (!existing.Contains(workbook))
                {
                    problems.Add(listedBy is null
                        ? $"The board workbook [{workbook}] is not in the data."
                        : $"The master workbook [{listedBy}] lists [{workbook}], which is not in the data.");
                    continue;
                }

                string full = Path.Combine(root, workbook.Replace('/', Path.DirectorySeparatorChar));

                if (!DataTreeUsage.TryReadCitations(full, cache, out IReadOnlyCollection<string> cites))
                {
                    problems.Add($"The board workbook [{workbook}] could not be read.");
                    continue;
                }

                foreach (string cited in cites)
                    DataTreeUsage.AddNormalised(used, cited);
            }

            return new DataTreeUsageResult(files, used, usedFolders, problems, masters.Count, boardWorkbooks.Count);
        }

        // ###########################################################################################
        // What a board stopped citing: in `before` and not in `after`, ignoring case (a file cited
        // again in another case is still used). The CANDIDATES for removal by one publish; only
        // those the whole tree no longer uses are actually removed (RemovableFrom).
        // ###########################################################################################
        public static IReadOnlyList<string> NoLongerCited(IEnumerable<string>? before, IEnumerable<string>? after)
        {
            var stillCited = new HashSet<string>(
                (after ?? []).Select(DataTreeUsage.Normalise).Where(path => path.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            return (before ?? [])
                .Select(DataTreeUsage.Normalise)
                .Where(path => path.Length > 0 && !stillCited.Contains(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.Ordinal)
                .ToList();
        }

        // "Classic-Repair-Toolbox.xlsx" and "Classic-Repair-Toolbox.v<version>.xlsx".
        public static bool IsMasterFileName(string? fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) ||
                !fileName.EndsWith(DataGenerationRules.WorkbookExtension, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string stem = fileName[..^DataGenerationRules.WorkbookExtension.Length];

            return string.Equals(stem, DataGenerationRules.MasterStem, StringComparison.OrdinalIgnoreCase) ||
                   stem.StartsWith(DataGenerationRules.MasterStem + ".v", StringComparison.OrdinalIgnoreCase);
        }

        // A workbook at the top of <Manufacturer>/<Hardware>/<Board>/, outside the shared folders,
        // and not an office lock file ("~$Data C64 250407.xlsx" while Excel has it open).
        public static bool IsBoardFolderWorkbook(string path)
        {
            string[] segments = path.Split('/');

            if (segments.Length != 4 ||
                !segments[3].EndsWith(DataGenerationRules.WorkbookExtension, StringComparison.OrdinalIgnoreCase) ||
                segments[3].StartsWith("~$", StringComparison.Ordinal))
            {
                return false;
            }

            return !SubmissionFileScopes.IsSharedFolderName(segments[0]) && !SubmissionFileScopes.IsSharedFolderName(segments[1]);
        }

        // A master's listing and a board's citations, through the cache when there is one.
        private static bool TryReadListed(string masterPath, WorkbookReadCache? cache, out IReadOnlyCollection<string> listed, out string why)
        {
            if (cache is not null)
                return cache.TryGetListing(masterPath, DataTreeUsage.ReadListed, out listed, out why);

            return DataTreeUsage.ReadListed(masterPath, out listed, out why);
        }

        private static bool ReadListed(string masterPath, out IReadOnlyCollection<string> listed, out string why)
        {
            bool ok = DataTreeUsage.TryReadListedWorkbooks(masterPath, out IReadOnlyList<string> paths, out why);
            listed = paths;
            return ok;
        }

        private static bool TryReadCitations(string workbookPath, WorkbookReadCache? cache, out IReadOnlyCollection<string> cites)
        {
            if (cache is not null)
                return cache.TryGetCitations(workbookPath, out cites);

            bool ok = BoardDataReader.TryCollectReferencedLocalFiles(workbookPath, out HashSet<string> files);
            cites = files;
            return ok;
        }

        // ###########################################################################################
        // The master's "Excel data file" column, normalised. False (with a reason) when the sheet or
        // its header is missing or the file cannot be opened - never an empty list, which would read
        // as "this master lists nothing" and fail open.
        // ###########################################################################################
        private static bool TryReadListedWorkbooks(string masterPath, out IReadOnlyList<string> listed, out string why)
        {
            listed = [];
            why = string.Empty;

            EpplusLicense.Ensure();

            try
            {
                using var stream = new FileStream(masterPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);
                ExcelWorksheet? sheet = package.Workbook.Worksheets[DataTreeUsage.MasterSheetName];

                if (sheet?.Dimension is null)
                {
                    why = $"it has no [{DataTreeUsage.MasterSheetName}] sheet.";
                    return false;
                }

                int lastRow = sheet.Dimension.End.Row;
                int lastColumn = sheet.Dimension.End.Column;
                int headerRow = -1;
                int column = -1;

                for (int row = 1; row <= Math.Min(DataTreeUsage.HeaderSearchRows, lastRow) && column < 0; row++)
                {
                    for (int col = 1; col <= lastColumn; col++)
                    {
                        if (string.Equals(sheet.Cells[row, col].Text?.Trim(), DataTreeUsage.ExcelDataFileColumn, StringComparison.OrdinalIgnoreCase))
                        {
                            headerRow = row;
                            column = col;
                            break;
                        }
                    }
                }

                if (column < 0)
                {
                    why = $"its [{DataTreeUsage.MasterSheetName}] sheet has no [{DataTreeUsage.ExcelDataFileColumn}] column.";
                    return false;
                }

                var paths = new List<string>();

                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    string path = DataTreeUsage.Normalise(sheet.Cells[row, column].Text);

                    if (path.Length > 0)
                        paths.Add(path);
                }

                listed = paths;
                return true;
            }
            catch (Exception ex)
            {
                why = ex.Message;
                return false;
            }
        }

        // ###########################################################################################
        // Every file the sync manifest would list, data-root-relative with "/", sorted. A folder
        // that cannot be walked is a PROBLEM, not a gap: files not seen are files not known to be
        // used. Links are not followed - a link is never part of the data.
        // ###########################################################################################
        private static IReadOnlyList<string> ListFiles(string root, List<string> problems)
        {
            try
            {
                var options = new EnumerationOptions
                {
                    RecurseSubdirectories = true,
                    IgnoreInaccessible = false,
                    AttributesToSkip = FileAttributes.ReparsePoint
                };

                return Directory
                    .EnumerateFiles(root, "*", options)
                    .Select(path => Path.GetRelativePath(root, path).Replace(Path.DirectorySeparatorChar, '/'))
                    .Where(DataChecksumManifest.IsSyncable)
                    .OrderBy(path => path, StringComparer.Ordinal)
                    .ToList();
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                problems.Add($"The data tree could not be read completely: {ex.Message}");
                return [];
            }
        }

        private static void AddNormalised(HashSet<string> set, string? path)
        {
            string normalised = DataTreeUsage.Normalise(path);

            if (normalised.Length > 0)
                set.Add(normalised);
        }

        // The app's own normalisation (DataManager, BoardDataReader): trimmed, "/" separators, no
        // leading slash.
        internal static string Normalise(string? path) =>
            string.IsNullOrWhiteSpace(path) ? string.Empty : path.Trim().Replace('\\', '/').TrimStart('/');

        private static string FolderOf(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? string.Empty : path[..slash];
        }
    }

    // ###########################################################################################
    // The answer for one tree. `IsUsed` is the rule; `UnusedFiles` and `RemovableFrom` are what may
    // be deleted, and both are EMPTY when the result is incomplete - see the class above.
    // ###########################################################################################
    public sealed class DataTreeUsageResult
    {
        private readonly HashSet<string> thisUsed;
        private readonly IReadOnlyList<string> thisUsedFolderPrefixes;
        private readonly HashSet<string> thisFiles;

        internal DataTreeUsageResult(
            IReadOnlyList<string> files,
            HashSet<string> used,
            HashSet<string> usedFolders,
            IReadOnlyList<string> problems,
            int masterCount,
            int boardWorkbookCount)
        {
            this.Files = files;
            this.thisFiles = new HashSet<string>(files, StringComparer.OrdinalIgnoreCase);
            this.thisUsed = used;
            this.thisUsedFolderPrefixes = usedFolders.Select(folder => folder.TrimEnd('/') + "/").ToList();
            this.Problems = problems;
            this.MasterCount = masterCount;
            this.BoardWorkbookCount = boardWorkbookCount;
        }

        // Every file considered, data-root-relative, sorted.
        public IReadOnlyList<string> Files { get; }

        // Why nothing may be removed; empty when the result is complete.
        public IReadOnlyList<string> Problems { get; }

        public bool IsComplete => this.Problems.Count == 0;

        public int MasterCount { get; }

        public int BoardWorkbookCount { get; }

        public bool IsUsed(string? relativePath)
        {
            string path = DataTreeUsage.Normalise(relativePath);

            if (path.Length == 0)
                return false;

            if (this.thisUsed.Contains(path))
                return true;

            string name = path[(path.LastIndexOf('/') + 1)..];

            if (name.StartsWith(DataTreeUsage.DocumentationPrefix, StringComparison.Ordinal))
                return true;

            return this.thisUsedFolderPrefixes.Any(prefix => path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
        }

        // Every file in the tree nothing uses - the administrator's "Unused files" list.
        public IReadOnlyList<string> UnusedFiles =>
            this.IsComplete ? this.Files.Where(path => !this.IsUsed(path)).ToList() : [];

        // Those of `candidates` that are in the tree and unused, in the tree's own spelling.
        public IReadOnlyList<string> RemovableFrom(IEnumerable<string>? candidates)
        {
            if (!this.IsComplete)
                return [];

            var wanted = new HashSet<string>(
                (candidates ?? []).Select(DataTreeUsage.Normalise).Where(path => path.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            return this.Files
                .Where(path => wanted.Contains(path) && !this.IsUsed(path))
                .ToList();
        }

        public bool Contains(string? relativePath) => this.thisFiles.Contains(DataTreeUsage.Normalise(relativePath));
    }
}
