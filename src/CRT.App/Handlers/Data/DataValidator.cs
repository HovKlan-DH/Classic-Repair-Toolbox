using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE LAUNCH CHECK OF EVERY BOARD, WRITTEN TO THE LOG FILE.
    //
    // *** ITS RULES ARE BoardDataChecks' SINCE 2026-10-02. *** They were written out here, logging
    // warnings nowhere else, until the Drafts tab's table started showing the same checks on the
    // cells they are about (owner request: "integrate the existing validation check into the
    // "Draft" system table view"). The project owner chose to KEEP this log as well - it is the only
    // place a problem in a board nobody is editing shows up - so it now runs those same rules
    // (BoardCheckScope.Everything, files looked for in the downloaded data), and the log and the
    // table can never disagree about what is wrong.
    //
    // The master workbook's own board files ("Hardware & Board") are not in any board's rows, so
    // that one check stays here.
    // ###########################################################################################
    public static class DataValidator
    {
        // ###########################################################################################
        // Validates all data definitions and paths across the main Excel file and all board-specific
        // files in the background, emitting warnings to the log for any inconsistencies found.
        // ###########################################################################################
        public static async Task ValidateAllDataAsync()
        {
            Logger.Info("Starting background data validation");

            foreach (var entry in DataManager.HardwareBoards)
            {
                if (entry.IsDraftOnly)
                {
                    // A board that exists only as a local draft (session 2c, task 9) has no files
                    // under the data root at all - its ExcelDataFile is an identity key naming a
                    // file that is never created, and its schematics and attachments live under
                    // "Drafts/". Validating it here would report every one of them as missing on
                    // every launch, which is noise that would train the reader to ignore this log.
                    // Its problems are shown in its table on the Drafts tab instead.
                    continue;
                }

                // Check main excel board file path
                ValidateMainWorkbookFile(entry.ExcelDataFile);

                if (string.IsNullOrWhiteSpace(entry.ExcelDataFile))
                    continue;

                // Load board data to validate its internal paths (this also effectively pre-warms the cache)
                var boardData = await DataManager.LoadBoardDataAsync(entry);
                if (boardData == null) continue;

                // The board as loaded - a draft's, when the board has one - so its files are looked
                // for as a submit looks for them: the draft's own folder first, then the downloaded data.
                IReadOnlyList<BoardDataProblem> problems = BoardDataChecks.Check(
                    BoardCheckRows.From(boardData),
                    new DiskFileLookup(
                        DataManager.DataRoot,
                        DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, entry.ExcelDataFile)),
                    BoardCheckScope.Everything);

                foreach (BoardDataProblem problem in problems)
                {
                    Logger.Warning(DataValidator.Describe(entry.ExcelDataFile, problem));
                }
            }

            Logger.Info("Background data validation complete");
        }

        // ###########################################################################################
        // One problem as a log line: which file, which sheet and row (1-based, as Excel numbers its
        // data rows under the header) and column - or the JSON beside it, for a highlight - then the
        // problem in the table's own words.
        // ###########################################################################################
        internal static string Describe(string excelDataFile, BoardDataProblem problem)
        {
            string where = problem.IsInSheet
                ? $"sheet [{problem.Sheet}] row [{problem.Index + 1}]{(problem.Column is null ? string.Empty : $" column [{problem.Column}]")}"
                : $"JSON file [{Path.ChangeExtension(excelDataFile, ".json")}]";

            string level = problem.Level == BoardProblemLevel.Error ? "error" : "warning";

            return $"Excel data file [{excelDataFile}] {where} has {level} [{problem.Code}]: {problem.Message} - please fix!";
        }

        // ###########################################################################################
        // The main workbook's board file - in no board's rows, so not one of BoardDataChecks' rules:
        // named at all, with forward slashes, there, and spelled exactly.
        // ###########################################################################################
        private static void ValidateMainWorkbookFile(string? file)
        {
            const string sheetName = "Hardware & Board";

            if (string.IsNullOrWhiteSpace(file))
            {
                Logger.Warning($"Main Excel file sheet [{sheetName}] has an entry with an empty file name - please fix!");
                return;
            }

            if (file.Contains('\\'))
            {
                Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] uses backslash instead of forward slash - please fix!");
            }

            // Clean the path characters so the existence check works regardless of the format issue
            var safeFile = file.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            var fullPath = Path.Combine(DataManager.DataRoot, safeFile);

            if (!File.Exists(fullPath))
            {
                Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] does not exist - please fix!");
            }
            else if (!HasExactCaseMatch(DataManager.DataRoot, safeFile))
            {
                Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] has incorrect casing (UPPER/lowercase) - please fix!");
            }
        }

        // ###########################################################################################
        // Verifies that a relative path perfectly matches the case of the folders/files on the disk.
        // Necessary because Windows File.Exists is case-insensitive, but Linux/web-hosts are not.
        // ###########################################################################################
        private static bool HasExactCaseMatch(string rootDir, string relativePath)
        {
            var segments = relativePath.Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
            var currentPath = rootDir;

            foreach (var segment in segments)
            {
                if (!Directory.Exists(currentPath))
                    return true; // Handled by File.Exists

                bool foundMatch = false;

                foreach (var entry in Directory.EnumerateFileSystemEntries(currentPath))
                {
                    var entryName = Path.GetFileName(entry);
                    if (string.Equals(entryName, segment, StringComparison.OrdinalIgnoreCase))
                    {
                        if (!string.Equals(entryName, segment, StringComparison.Ordinal))
                        {
                            return false; // Case mismatch detected
                        }

                        currentPath = entry; // Advance deeper using real casing
                        foundMatch = true;
                        break;
                    }
                }

                if (!foundMatch)
                    return true; // Handled by File.Exists
            }

            return true;
        }
    }
}
