using System;
using System.IO;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
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
                // Check main excel board file path
                ValidateFile(string.Empty, "Hardware & Board", entry.ExcelDataFile);

                if (string.IsNullOrWhiteSpace(entry.ExcelDataFile))
                    continue;

                // Load board data to validate its internal paths (this also effectively pre-warms the cache)
                var boardData = await DataManager.LoadBoardDataAsync(entry);
                if (boardData == null) continue;

                string contextName = entry.ExcelDataFile;

                foreach (var schematic in boardData.Schematics)
                {
                    ValidateFile(contextName, "Board schematics", schematic.SchematicImageFile);
                }

                foreach (var image in boardData.ComponentImages)
                {
                    string componentImageEntry = $"[{image.BoardLabel.Trim()}] pin [{image.Pin.Trim()}]";
                    ValidateFile(contextName, "Component images", image.File, componentImageEntry);
                }

                foreach (var localFile in boardData.ComponentLocalFiles)
                {
                    ValidateFile(contextName, "Component local files", localFile.File);
                }

                foreach (var boardLocalFile in boardData.BoardLocalFiles)
                {
                    ValidateFile(contextName, "Board local files", boardLocalFile.File);
                }
            }

            Logger.Info("Background data validation complete");
        }

        // ###########################################################################################
        // Validates a single path for empty values, backslashes, existence on disk, and exact case match.
        // ###########################################################################################
        private static void ValidateFile(string excelDataFile, string sheetName, string? file, string? entryLabel = null)
        {
            bool isMain = string.IsNullOrEmpty(excelDataFile);

            if (string.IsNullOrWhiteSpace(file))
            {
                if (isMain)
                {
                    Logger.Warning(
                        string.IsNullOrWhiteSpace(entryLabel)
                            ? $"Main Excel file sheet [{sheetName}] has an entry with an empty file name - please fix!"
                            : $"Main Excel file sheet [{sheetName}] has an entry {entryLabel} with an empty file name - please fix!");
                }
                else
                {
                    Logger.Warning(
                        string.IsNullOrWhiteSpace(entryLabel)
                            ? $"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry with an empty file name - please fix!"
                            : $"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry {entryLabel} with an empty file name - please fix!");
                }

                return;
            }

            if (file.Contains('\\'))
            {
                if (isMain)
                    Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] uses backslash instead of forward slash - please fix!");
                else
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] and file [{file}] uses backslash instead of forward slash - please fix!");
            }

            // Clean the path characters so the existence check works regardless of the format issue
            var safeFile = file.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
            var fullPath = Path.Combine(DataManager.DataRoot, safeFile);

            if (!File.Exists(fullPath))
            {
                if (isMain)
                    Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] does not exist - please fix!");
                else
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] and file [{file}] does not exist - please fix!");
            }
            else if (!HasExactCaseMatch(DataManager.DataRoot, safeFile))
            {
                if (isMain)
                    Logger.Warning($"Main Excel file sheet [{sheetName}] and file [{file}] has incorrect casing (UPPER/lowercase) - please fix!");
                else
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] and file [{file}] has incorrect casing (UPPER/lowercase) - please fix!");
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