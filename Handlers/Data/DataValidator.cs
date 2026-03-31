using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    public static class DataValidator
    {
        private static readonly HashSet<string> AllowedTimeDivValues = new(StringComparer.OrdinalIgnoreCase)
        {
            "2nS", "5nS", "10nS", "20nS", "50nS", "100nS", "200nS", "500nS",
            "1uS", "2uS", "5uS", "10uS", "20uS", "50uS", "100uS", "200uS", "500uS",
            "1mS", "2mS", "5mS", "10mS", "20mS", "50mS", "100mS", "200mS", "500mS",
            "1S", "2S", "5S", "10S", "20S", "50S", "100S", "200S", "500S", "1000S"
        };

        private static readonly HashSet<string> AllowedVoltsDivValues = new(StringComparer.OrdinalIgnoreCase)
        {
            "5mV", "10mV", "20mV", "50mV", "100mV", "200mV", "500mV",
            "1V", "2V", "5V", "10V", "20V", "50V", "100V"
        };

        private static readonly Regex TriggerLevelRegex = new(
            @"^-?\d+(?:\.\d+)?[Vv]$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

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
                    ValidateComponentImageTimeDiv(contextName, componentImageEntry, image.TimeDiv);
                    ValidateComponentImageVoltsDiv(contextName, componentImageEntry, image.VoltsDiv);
                    ValidateComponentImageTriggerLevel(contextName, componentImageEntry, image.TriggerLevelVolts);
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
                /*
                                                else
                                                {
                                                    Logger.Warning(
                                                        string.IsNullOrWhiteSpace(entryLabel)
                                                            ? $"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry with an empty file name - please fix!"
                                                            : $"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry {entryLabel} with an empty file name - please fix!");
                                                }
                */
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
        // Validates T/DIV against the supported oscilloscope timing list.
        // Empty values are allowed and ignored.
        // ###########################################################################################
        private static void ValidateComponentImageTimeDiv(string excelDataFile, string entryLabel, string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            if (AllowedTimeDivValues.Contains(value))
                return;

            Logger.Warning($"Excel data file [{excelDataFile}] sheet [Component images] has an entry {entryLabel} with invalid [T/DIV] value [{value}] - please fix!");
        }

        // ###########################################################################################
        // Validates V/DIV against the supported oscilloscope voltage-per-division list.
        // Empty values are allowed and ignored.
        // ###########################################################################################
        private static void ValidateComponentImageVoltsDiv(string excelDataFile, string entryLabel, string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            if (AllowedVoltsDivValues.Contains(value))
                return;

            Logger.Warning($"Excel data file [{excelDataFile}] sheet [Component images] has an entry {entryLabel} with invalid [V/DIV] value [{value}] - please fix!");
        }

        // ###########################################################################################
        // Validates T.LVL as a signed or unsigned integer/decimal voltage ending with V or v.
        // Empty values are allowed and ignored.
        // ###########################################################################################
        private static void ValidateComponentImageTriggerLevel(string excelDataFile, string entryLabel, string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return;

            if (TriggerLevelRegex.IsMatch(value))
                return;

            Logger.Warning($"Excel data file [{excelDataFile}] sheet [Component images] has an entry {entryLabel} with invalid [T.LVL] value [{value}] - please fix!");
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