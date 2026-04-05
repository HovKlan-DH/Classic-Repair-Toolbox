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

        private static readonly Regex UuidV4Regex = new(
            @"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-4[0-9a-fA-F]{3}-[89aAbB][0-9a-fA-F]{3}-[0-9a-fA-F]{12}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        // ###########################################################################################
        // Validates all data definitions and paths across the main Excel file and all board-specific
        // files in the background, emitting warnings to the log for any inconsistencies found.
        // ###########################################################################################
        public static async Task ValidateAllDataAsync()
        {
            Logger.Info("Starting background data validation");

            var seenUuids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

                ValidateBoardUuids(contextName, boardData, seenUuids);
                ValidateOrphanComponents(contextName, boardData);
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
        // Validates UUID v4 values across all UUID-bearing sheets in the board Excel data file.
        // Logs empty UUIDs, invalid UUID v4 format, and duplicates across all loaded boards/sheets.
        // ###########################################################################################
        private static void ValidateBoardUuids(string excelDataFile, BoardData boardData, Dictionary<string, string> seenUuids)
        {
            ValidateUuidsInSheet(
                excelDataFile,
                "Board schematics",
                boardData.Schematics,
                schematic => schematic.UuidV4,
                schematic => $"schematic [{schematic.SchematicName.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Components",
                boardData.Components,
                component => component.UuidV4,
                component => $"component [{component.BoardLabel.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Component images",
                boardData.ComponentImages,
                image => image.UuidV4,
                image => $"component image [{image.BoardLabel.Trim()}] pin [{image.Pin.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Component local files",
                boardData.ComponentLocalFiles,
                localFile => localFile.UuidV4,
                localFile => $"component local file [{localFile.BoardLabel.Trim()}] name [{localFile.Name.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Component links",
                boardData.ComponentLinks,
                link => link.UuidV4,
                link => $"component link [{link.BoardLabel.Trim()}] name [{link.Name.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Board local files",
                boardData.BoardLocalFiles,
                localFile => localFile.UuidV4,
                localFile => $"board local file [{localFile.Category.Trim()}] name [{localFile.Name.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Board links",
                boardData.BoardLinks,
                link => link.UuidV4,
                link => $"board link [{link.Category.Trim()}] name [{link.Name.Trim()}]",
                seenUuids);

            ValidateUuidsInSheet(
                excelDataFile,
                "Credits",
                boardData.Credits,
                credit => credit.UuidV4,
                credit => $"credit [{credit.Category.Trim()}] name [{credit.NameOrHandle.Trim()}]",
                seenUuids);
        }

        // ###########################################################################################
        // Validates one sheet's UUID values for presence, strict UUID v4 format, and global uniqueness.
        // Duplicate checks only run after the UUID has passed format validation.
        // ###########################################################################################
        private static void ValidateUuidsInSheet<T>(
            string excelDataFile,
            string sheetName,
            IEnumerable<T> entries,
            Func<T, string> getUuid,
            Func<T, string> getEntryLabel,
            Dictionary<string, string> seenUuids)
        {
            foreach (var entry in entries)
            {
                string entryLabel = getEntryLabel(entry);
                string uuid = (getUuid(entry) ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(uuid))
                {
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry {entryLabel} with an empty [UUID v4] value - please fix!");
                    continue;
                }

                if (!UuidV4Regex.IsMatch(uuid))
                {
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] has an entry {entryLabel} with invalid [UUID v4] value [{uuid}] - please fix!");
                    continue;
                }

                string currentLocation = $"Excel data file [{excelDataFile}] sheet [{sheetName}] entry {entryLabel}";

                if (seenUuids.TryGetValue(uuid, out var firstLocation))
                {
                    Logger.Warning($"Duplicate [UUID v4] value [{uuid}] found in {currentLocation}; already used in {firstLocation} - please fix!");
                    continue;
                }

                seenUuids[uuid] = currentLocation;
            }
        }

        // ###########################################################################################
        // Detects orphan component references between the component-related board Excel sheets.
        // Uses [BoardLabel] as the shared component key and emits warnings for any missing links.
        // ###########################################################################################
        private static void ValidateOrphanComponents(string excelDataFile, BoardData boardData)
        {
            var componentLabels = CreateNormalizedLabelSet(boardData.Components, component => component.BoardLabel);
            var highlightLabels = CreateNormalizedLabelSet(boardData.ComponentHighlights, highlight => highlight.BoardLabel);

            foreach (var component in boardData.Components)
            {
                string boardLabel = NormalizeLabel(component.BoardLabel);
                if (string.IsNullOrWhiteSpace(boardLabel))
                    continue;

                if (!highlightLabels.Contains(boardLabel))
                {
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [Components] has an orphan component [{component.BoardLabel.Trim()}] that does not exist in sheet [Component highlights] - please fix!");
                }
            }

            ValidateComponentReferencesInSheet(
                excelDataFile,
                "Component images",
                boardData.ComponentImages,
                image => image.BoardLabel,
                image => $"component image [{image.BoardLabel.Trim()}] pin [{image.Pin.Trim()}]",
                componentLabels);

            ValidateComponentReferencesInSheet(
                excelDataFile,
                "Component highlights",
                boardData.ComponentHighlights,
                highlight => highlight.BoardLabel,
                highlight => $"component highlight [{highlight.BoardLabel.Trim()}] schematic [{highlight.SchematicName.Trim()}]",
                componentLabels);

            ValidateComponentReferencesInSheet(
                excelDataFile,
                "Component local files",
                boardData.ComponentLocalFiles,
                localFile => localFile.BoardLabel,
                localFile => $"component local file [{localFile.BoardLabel.Trim()}] name [{localFile.Name.Trim()}]",
                componentLabels);

            ValidateComponentReferencesInSheet(
                excelDataFile,
                "Component links",
                boardData.ComponentLinks,
                link => link.BoardLabel,
                link => $"component link [{link.BoardLabel.Trim()}] name [{link.Name.Trim()}]",
                componentLabels);
        }

        // ###########################################################################################
        // Validates that entries in a component-related sheet reference a component label that
        // actually exists in the [Components] sheet.
        // ###########################################################################################
        private static void ValidateComponentReferencesInSheet<T>(
            string excelDataFile,
            string sheetName,
            IEnumerable<T> entries,
            Func<T, string> getBoardLabel,
            Func<T, string> getEntryLabel,
            HashSet<string> componentLabels)
        {
            foreach (var entry in entries)
            {
                string boardLabel = NormalizeLabel(getBoardLabel(entry));
                if (string.IsNullOrWhiteSpace(boardLabel))
                    continue;

                if (!componentLabels.Contains(boardLabel))
                {
                    Logger.Warning($"Excel data file [{excelDataFile}] sheet [{sheetName}] has an orphan entry {getEntryLabel(entry)} because component [{boardLabel}] does not exist in sheet [Components] - please fix!");
                }
            }
        }

        // ###########################################################################################
        // Builds a case-insensitive label set from a sequence of entries using trimmed board labels.
        // Empty labels are ignored.
        // ###########################################################################################
        private static HashSet<string> CreateNormalizedLabelSet<T>(IEnumerable<T> entries, Func<T, string> getLabel)
        {
            var labels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in entries)
            {
                string label = NormalizeLabel(getLabel(entry));
                if (!string.IsNullOrWhiteSpace(label))
                {
                    labels.Add(label);
                }
            }

            return labels;
        }

        // ###########################################################################################
        // Normalizes a component label by trimming surrounding whitespace.
        // ###########################################################################################
        private static string NormalizeLabel(string? label)
        {
            return (label ?? string.Empty).Trim();
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