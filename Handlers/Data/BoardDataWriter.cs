using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // DTO used by the schematics label editor save pipeline to persist rectangles and any newly
    // introduced component labels back into the board-specific Excel workbook.
    // ###########################################################################################
    internal sealed class LabelEditorSaveRow
    {
        public string SchematicName { get; init; } = string.Empty;
        public string BoardLabel { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string Region { get; init; } = string.Empty;
        public double X { get; init; }
        public double Y { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
    }

    // ###########################################################################################
    // Writes label-editor changes back into the board Excel workbook while preserving unrelated
    // schematics and unrelated component rows.
    // ###########################################################################################
    internal static class BoardDataWriter
    {
        private const string SheetComponents = "Components";
        private const string SheetComponentHighlights = "Component highlights";

        private const string ColUuidV4 = "UUID v4";
        private const string ColBoardLabel = "Board label";
        private const string ColCategory = "Category";
        private const string ColRegion = "Region";
        private const string ColFriendlyName = "Friendly name";
        private const string ColTechnicalNameOrValue = "Technical name or value";
        private const string ColPartNumber = "Part-number";
        private const string ColDescription = "Short one-liner description (one short line only!)";

        private const string ColSchematicName = "Schematic name";
        private const string ColX = "X";
        private const string ColY = "Y";
        private const string ColWidth = "Width";
        private const string ColHeight = "Height";

        private static readonly string[] ComponentHighlightsHeaders = new[]
        {
            ColSchematicName,
            ColBoardLabel,
            ColX,
            ColY,
            ColWidth,
            ColHeight
        };

        private static readonly string[] ComponentsHeaders = new[]
        {
            ColBoardLabel,
            ColCategory,
            ColRegion
        };

        // ###########################################################################################
        // Saves the current schematic editor state to the workbook by replacing only the edited
        // schematic's highlight rows and appending any newly introduced component rows if needed.
        // Uses a FileInfo-backed EPPlus package so workbook changes are committed directly to disk.
        // ###########################################################################################
        public static async Task<(bool Success, string ErrorMessage)> SaveLabelEditorChangesAsync(
            string excelPath,
            string schematicName,
            IReadOnlyList<LabelEditorSaveRow> rows,
            string region)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                Logger.Warning($"BoardDataWriter save aborted because Excel file does not exist - [{excelPath}]");
                return (false, $"Board Excel file not found - [{excelPath}]");
            }

            if (string.IsNullOrWhiteSpace(schematicName))
            {
                Logger.Warning("BoardDataWriter save aborted because no schematic name was provided");
                return (false, "No schematic is currently selected");
            }

            return await Task.Run(() =>
            {
                ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");

                try
                {
                    var fileInfo = new FileInfo(excelPath);

                    Logger.Debug($"BoardDataWriter opening Excel file for write: [{excelPath}]");
                    Logger.Debug($"BoardDataWriter target schematic: [{schematicName}]");
                    Logger.Debug($"BoardDataWriter target region: [{region}]");
                    Logger.Debug($"BoardDataWriter received save rows: [{rows.Count}]");
                    Logger.Debug($"BoardDataWriter file exists: [{fileInfo.Exists}]");
                    Logger.Debug($"BoardDataWriter file is read-only: [{fileInfo.IsReadOnly}]");
                    Logger.Debug($"BoardDataWriter file last write time before save: [{fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss.fff}]");

                    using (var package = new ExcelPackage(fileInfo))
                    {
                        var highlightsSheet = package.Workbook.Worksheets[SheetComponentHighlights];
                        if (highlightsSheet == null)
                        {
                            Logger.Warning($"BoardDataWriter could not find sheet [{SheetComponentHighlights}] in [{excelPath}]");
                            return (false, $"Sheet [{SheetComponentHighlights}] not found in board Excel file");
                        }

                        var highlightsColMap = FindHeaderRow(highlightsSheet, ComponentHighlightsHeaders, out int highlightsHeaderRow);
                        if (highlightsColMap == null)
                        {
                            Logger.Warning($"BoardDataWriter could not find header row in sheet [{SheetComponentHighlights}]");
                            return (false, $"Header row not found in sheet [{SheetComponentHighlights}]");
                        }

                        Logger.Info($"BoardDataWriter replacing highlight block for schematic [{schematicName}]");
                        ReplaceHighlightsForSchematic(highlightsSheet, highlightsColMap, highlightsHeaderRow, schematicName, rows);

                        var componentsSheet = package.Workbook.Worksheets[SheetComponents];
                        if (componentsSheet == null)
                        {
                            Logger.Warning($"BoardDataWriter could not find sheet [{SheetComponents}] in [{excelPath}]");
                            return (false, $"Sheet [{SheetComponents}] not found in board Excel file");
                        }

                        var componentsColMap = FindHeaderRow(componentsSheet, ComponentsHeaders, out int componentsHeaderRow);
                        if (componentsColMap == null)
                        {
                            Logger.Warning($"BoardDataWriter could not find header row in sheet [{SheetComponents}]");
                            return (false, $"Header row not found in sheet [{SheetComponents}]");
                        }

//                        Logger.Info("BoardDataWriter appending any missing component rows");
                        AppendMissingComponents(componentsSheet, componentsColMap, componentsHeaderRow, rows, region);

//                        Logger.Info($"BoardDataWriter about to save workbook changes for schematic [{schematicName}]");
                        package.Save();
                    }

                    fileInfo.Refresh();
                    Logger.Info($"BoardDataWriter completed workbook save for Excel file [{excelPath}]");
//                    Logger.Info($"BoardDataWriter file last write time after save: [{fileInfo.LastWriteTime:yyyy-MM-dd HH:mm:ss.fff}]");

                    return (true, string.Empty);
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to save label editor changes to Excel file [{excelPath}] - [{ex}]");
                    return (false, BuildSaveFailureMessage(ex, excelPath));
                }
            });
        }

        // ###########################################################################################
        // Replaces only the current schematic's highlight rows in the worksheet and preserves all
        // highlight rows belonging to other schematics. New rows are written in stable grouped
        // order so labels stay together by category and board label.
        // ###########################################################################################
        private static void ReplaceHighlightsForSchematic(
            ExcelWorksheet sheet,
            Dictionary<string, int> colMap,
            int headerRow,
            string schematicName,
            IReadOnlyList<LabelEditorSaveRow> rows)
        {
            int maxRow = sheet.Dimension?.End.Row ?? headerRow;
            int deletedRowCount = 0;

            for (int row = maxRow; row > headerRow; row--)
            {
                string existingSchematicName = GetCellText(sheet, row, colMap[ColSchematicName]);
                if (string.Equals(existingSchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                {
                    sheet.DeleteRow(row);
                    deletedRowCount++;
                }
            }

            var orderedRows = rows
                .Where(item => string.Equals(item.SchematicName, schematicName, StringComparison.OrdinalIgnoreCase))
                .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                .OrderBy(item => item.Category?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.BoardLabel?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Y)
                .ThenBy(item => item.X)
                .ToList();

//            Logger.Debug($"BoardDataWriter deleted highlight rows for schematic [{schematicName}]: [{deletedRowCount}]");
            Logger.Debug($"BoardDataWriter will write highlight rows for schematic [{schematicName}]: [{orderedRows.Count}]");

            int appendRow = (sheet.Dimension?.End.Row ?? headerRow) + 1;

            foreach (var item in orderedRows)
            {
//                Logger.Debug(
//                    $"BoardDataWriter writing highlight row at Excel row [{appendRow}] -> Label=[{item.BoardLabel}] Category=[{item.Category}] X=[{item.X}] Y=[{item.Y}] Width=[{item.Width}] Height=[{item.Height}]");

                sheet.Cells[appendRow, colMap[ColSchematicName]].Value = item.SchematicName.Trim();
                sheet.Cells[appendRow, colMap[ColBoardLabel]].Value = item.BoardLabel.Trim();
                sheet.Cells[appendRow, colMap[ColX]].Value = FormatRoundedInteger(item.X);
                sheet.Cells[appendRow, colMap[ColY]].Value = FormatRoundedInteger(item.Y);
                sheet.Cells[appendRow, colMap[ColWidth]].Value = FormatRoundedInteger(item.Width);
                sheet.Cells[appendRow, colMap[ColHeight]].Value = FormatRoundedInteger(item.Height);
                appendRow++;
            }
        }

        // ###########################################################################################
        // Inserts only truly missing component rows. A blank existing region is treated as a
        // wildcard match so shared component rows are not duplicated for PAL or NTSC saves.
        // ###########################################################################################
        private static void AppendMissingComponents(
            ExcelWorksheet sheet,
            Dictionary<string, int> colMap,
            int headerRow,
            IReadOnlyList<LabelEditorSaveRow> rows,
            string region)
        {
            var rowsToInsert = rows
                .Where(item => !string.IsNullOrWhiteSpace(item.BoardLabel))
                .GroupBy(item => item.BoardLabel.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.First())
                .Where(item => !ComponentExistsInSheet(sheet, colMap, headerRow, item, region))
                .OrderBy(item => item.Category?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.BoardLabel?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

//            Logger.Info($"BoardDataWriter missing component rows to insert: [{rowsToInsert.Count}]");

            foreach (var item in rowsToInsert)
            {
                int insertRow = FindInsertRowForComponent(sheet, colMap, headerRow, item);
                int currentEndRow = sheet.Dimension?.End.Row ?? headerRow;

                if (insertRow <= currentEndRow)
                {
                    sheet.InsertRow(insertRow, 1);
                }
                else
                {
                    insertRow = currentEndRow + 1;
                }

                Logger.Info(
                    $"BoardDataWriter inserting component row at Excel row [{insertRow}] -> Label=[{item.BoardLabel}] Category=[{item.Category}] Region=[{(string.IsNullOrWhiteSpace(item.Region) ? region : item.Region.Trim())}]");

                WriteComponentRow(sheet, colMap, insertRow, item, region);
            }
        }

        // ###########################################################################################
        // Finds the first row that contains all required headers and returns the full header map
        // for that row, resolved by normalized header text.
        // ###########################################################################################
        private static Dictionary<string, int>? FindHeaderRow(ExcelWorksheet sheet, string[] requiredHeaders, out int headerRow)
        {
            headerRow = -1;

            int maxRow = sheet.Dimension?.End.Row ?? 0;
            int maxCol = sheet.Dimension?.End.Column ?? 0;

            for (int row = 1; row <= maxRow; row++)
            {
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (int col = 1; col <= maxCol; col++)
                {
                    string text = NormalizeHeader(GetCellText(sheet, row, col));
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        colMap[text] = col;
                    }
                }

                if (requiredHeaders.All(header => colMap.ContainsKey(header)))
                {
                    headerRow = row;
                    return colMap;
                }
            }

            return null;
        }

        // ###########################################################################################
        // Collapses Excel header line breaks into single spaces so matching stays robust.
        // ###########################################################################################
        private static string NormalizeHeader(string text)
        {
            var parts = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", parts).Trim();
        }

        // ###########################################################################################
        // Returns the trimmed worksheet cell text or an empty string if no text is present.
        // ###########################################################################################
        private static string GetCellText(ExcelWorksheet sheet, int row, int col)
        {
            return sheet.Cells[row, col].Text?.Trim() ?? string.Empty;
        }

        // ###########################################################################################
        // Finds the best insertion row so a new component is placed alongside other rows in the
        // same category and ordered by board label within that category.
        // ###########################################################################################
        private static int FindInsertRowForComponent(
            ExcelWorksheet sheet,
            Dictionary<string, int> colMap,
            int headerRow,
            LabelEditorSaveRow item)
        {
            string targetCategory = item.Category?.Trim() ?? string.Empty;
            string targetBoardLabel = item.BoardLabel?.Trim() ?? string.Empty;

            int maxRow = sheet.Dimension?.End.Row ?? headerRow;
            int? firstRowInCategory = null;
            int? insertBeforeRowInCategory = null;
            int? lastRowInCategory = null;

            for (int row = headerRow + 1; row <= maxRow; row++)
            {
                string existingBoardLabel = GetCellText(sheet, row, colMap[ColBoardLabel]);
                string existingCategory = GetCellText(sheet, row, colMap[ColCategory]);

                if (string.IsNullOrWhiteSpace(existingBoardLabel) && string.IsNullOrWhiteSpace(existingCategory))
                {
                    continue;
                }

                if (!string.Equals(existingCategory, targetCategory, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                firstRowInCategory ??= row;
                lastRowInCategory = row;

                if (string.Compare(existingBoardLabel, targetBoardLabel, StringComparison.OrdinalIgnoreCase) > 0)
                {
                    insertBeforeRowInCategory = row;
                    break;
                }
            }

            if (insertBeforeRowInCategory.HasValue)
            {
                return insertBeforeRowInCategory.Value;
            }

            if (lastRowInCategory.HasValue)
            {
                return lastRowInCategory.Value + 1;
            }

            return maxRow + 1;
        }

        // ###########################################################################################
        // Writes one component row using the available mapped columns and default blank values for
        // optional metadata fields that are not yet known for the newly created label.
        // ###########################################################################################
        private static void WriteComponentRow(
            ExcelWorksheet sheet,
            Dictionary<string, int> colMap,
            int row,
            LabelEditorSaveRow item,
            string region)
        {
            if (colMap.TryGetValue(ColUuidV4, out int uuidCol))
            {
                sheet.Cells[row, uuidCol].Value = Guid.NewGuid().ToString();
            }

            sheet.Cells[row, colMap[ColBoardLabel]].Value = item.BoardLabel.Trim();
            sheet.Cells[row, colMap[ColCategory]].Value = item.Category.Trim();
            sheet.Cells[row, colMap[ColRegion]].Value = string.IsNullOrWhiteSpace(item.Region) ? region : item.Region.Trim();

            if (colMap.TryGetValue(ColFriendlyName, out int friendlyNameCol))
            {
                sheet.Cells[row, friendlyNameCol].Value = string.Empty;
            }

            if (colMap.TryGetValue(ColTechnicalNameOrValue, out int technicalNameCol))
            {
                sheet.Cells[row, technicalNameCol].Value = string.Empty;
            }

            if (colMap.TryGetValue(ColPartNumber, out int partNumberCol))
            {
                sheet.Cells[row, partNumberCol].Value = string.Empty;
            }

            if (colMap.TryGetValue(ColDescription, out int descriptionCol))
            {
                sheet.Cells[row, descriptionCol].Value = string.Empty;
            }
        }

        // ###########################################################################################
        // Builds a stable uniqueness key for a component row using board label and region.
        // ###########################################################################################
        private static string BuildComponentKey(string boardLabel, string region)
        {
            return $"{boardLabel.Trim()}\u001F{region.Trim()}";
        }

        // ###########################################################################################
        // Returns true when the Components sheet already contains a row for the given board label.
        // A blank existing region is treated as a shared/global component row that matches any
        // requested region, preventing accidental duplication of existing components.
        // ###########################################################################################
        private static bool ComponentExistsInSheet(
            ExcelWorksheet sheet,
            Dictionary<string, int> colMap,
            int headerRow,
            LabelEditorSaveRow item,
            string defaultRegion)
        {
            string targetBoardLabel = item.BoardLabel?.Trim() ?? string.Empty;
            string targetRegion = string.IsNullOrWhiteSpace(item.Region) ? defaultRegion.Trim() : item.Region.Trim();

            int maxRow = sheet.Dimension?.End.Row ?? headerRow;

            for (int row = headerRow + 1; row <= maxRow; row++)
            {
                string existingBoardLabel = GetCellText(sheet, row, colMap[ColBoardLabel]);
                string existingRegion = GetCellText(sheet, row, colMap[ColRegion]);

                if (!string.Equals(existingBoardLabel, targetBoardLabel, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(existingRegion) ||
                    string.IsNullOrWhiteSpace(targetRegion) ||
                    string.Equals(existingRegion, targetRegion, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        // ###########################################################################################
        // Builds a user-facing save error message and specializes the common Excel file-lock case
        // so the schematics editor can show a clear action to the user.
        // ###########################################################################################
        private static string BuildSaveFailureMessage(Exception ex, string excelPath)
        {
            if (IsFileLockedForOverwrite(ex))
            {
                return "The Excel file is already open in another program. Close it and try saving again.";
            }

            return $"Error saving file {excelPath}";
        }

        // ###########################################################################################
        // Detects the common overwrite failure path raised when Excel or another process keeps the
        // workbook locked while EPPlus tries to save the updated file.
        // ###########################################################################################
        private static bool IsFileLockedForOverwrite(Exception ex)
        {
            for (Exception? current = ex; current != null; current = current.InnerException)
            {
                if (current is IOException ioException)
                {
                    if (ioException.HResult == unchecked((int)0x80070020) ||
                        ioException.Message.IndexOf("being used by another process", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        // ###########################################################################################
        // Formats a highlight coordinate or size as an integer string using midpoint rounding away
        // from zero so the saved Excel value stays compact while remaining predictable.
        // ###########################################################################################
        private static string FormatRoundedInteger(double value)
        {
            return Math.Round(value, 0, MidpointRounding.AwayFromZero)
                .ToString(CultureInfo.InvariantCulture);
        }

    }
}