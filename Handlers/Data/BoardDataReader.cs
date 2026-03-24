using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    internal static class BoardDataReader
    {
        // Sheet names
        private const string SheetBoardSchematics = "Board schematics";
        private const string SheetComponents = "Components";
        private const string SheetComponentImages = "Component images";
        private const string SheetComponentHighlights = "Component highlights";
        private const string SheetComponentLocalFiles = "Component local files";
        private const string SheetComponentLinks = "Component links";
        private const string SheetBoardLocalFiles = "Board local files";
        private const string SheetBoardLinks = "Board links";
        private const string SheetCredits = "Credits";

        // Board schematics columns
        private const string ColSchematicName = "Schematic name";
        private const string ColSchematicImageFile = "Schematic image file";
        private const string ColMainImageHighlightColor = "Main image highlight color";
        private const string ColMainHighlightOpacity = "Main highlight opacity";
        private const string ColThumbnailImageHighlightColor = "Thumbnail image highlight color";
        private const string ColThumbnailHighlightOpacity = "Thumbnail highlight opacity";

        // Shared columns (appear in multiple sheets)
        private const string ColBoardLabel = "Board label";
        private const string ColCategory = "Category";
        private const string ColName = "Name";
        private const string ColFile = "File";
        private const string ColUrl = "URL";
        private const string ColRegion = "Region";

        // Components columns
        private const string ColFriendlyName = "Friendly name";
        private const string ColTechnicalNameOrValue = "Technical name or value";
        private const string ColPartNumber = "Part-number";
        private const string ColDescription = "Short one-liner description (one short line only!)";

        // Component images columns
        private const string ColPin = "Pin";
        private const string ColExpectedOscilloscopeReading = "Expected oscilloscope reading";
        private const string ColNote = "Note";
        private const string ColTimeDiv = "T/DIV";
        private const string ColVoltsDiv = "V/DIV";
        private const string ColTriggerLevelVolts = "T.Level";

        // Component highlights columns
        private const string ColX = "X";
        private const string ColY = "Y";
        private const string ColWidth = "Width";
        private const string ColHeight = "Height";

        // Credits columns
        private const string ColSubCategory = "Sub-category";
        private const string ColNameOrHandle = "Name or handle";
        private const string ColContact = "Contact (email or web page)";

        // Required header sets per sheet - used for robust, order-independent column mapping
        // Compliant with .NET6
        private static readonly string[] SchematicsHeaders = new[] { ColSchematicName, ColSchematicImageFile, ColMainImageHighlightColor, ColMainHighlightOpacity, ColThumbnailImageHighlightColor, ColThumbnailHighlightOpacity };
        private static readonly string[] ComponentsHeaders = new[] { ColBoardLabel, ColFriendlyName, ColTechnicalNameOrValue, ColPartNumber, ColCategory, ColRegion, ColDescription };
        private static readonly string[] ComponentImagesHeaders = new[] { ColBoardLabel, ColRegion, ColPin, ColName, ColExpectedOscilloscopeReading, ColFile, ColNote, ColTimeDiv, ColVoltsDiv, ColTriggerLevelVolts };
        private static readonly string[] ComponentHighlightsHeaders = new[] { ColSchematicName, ColBoardLabel, ColX, ColY, ColWidth, ColHeight };
        private static readonly string[] ComponentLocalFilesHeaders = new[] { ColBoardLabel, ColName, ColFile };
        private static readonly string[] ComponentLinksHeaders = new[] { ColBoardLabel, ColName, ColUrl };
        private static readonly string[] BoardLocalFilesHeaders = new[] { ColCategory, ColName, ColFile };
        private static readonly string[] BoardLinksHeaders = new[] { ColCategory, ColName, ColUrl };
        private static readonly string[] CreditsHeaders = new[] { ColCategory, ColSubCategory, ColNameOrHandle, ColContact };

        private static readonly Dictionary<string, BoardData> _cache = new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // Lazily loads and caches all sheets from a board-specific Excel file.
        // Returns null if the file cannot be opened or parsed.
        // Subsequent calls for the same cacheKey return the cached instance instantly.
        // ###########################################################################################
        public static async Task<BoardData?> LoadAsync(string excelPath, string cacheKey)
        {
            if (_cache.TryGetValue(cacheKey, out var cached))
                return cached;

            if (!File.Exists(excelPath))
            {
                Logger.Warning($"Board Excel file not found: [{excelPath}]");
                return null;
            }

            return await Task.Run(() =>
            {
                ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");

                try
                {
                    using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var package = new ExcelPackage(stream);

                    string revisionDate = string.Empty;
                    var schematicsSheet = package.Workbook.Worksheets[SheetBoardSchematics];
                    if (schematicsSheet != null)
                    {
                        int limitRow = Math.Min(10, schematicsSheet.Dimension?.End.Row ?? 0);
                        int limitCol = Math.Min(10, schematicsSheet.Dimension?.End.Column ?? 0);
                        for (int r = 1; r <= limitRow; r++)
                        {
                            for (int c = 1; c <= limitCol; c++)
                            {
                                string text = GetCellText(schematicsSheet, r, c);
                                if (text.StartsWith("# Revision date:", StringComparison.OrdinalIgnoreCase))
                                {
                                    revisionDate = text.Substring("# Revision date:".Length).Trim();
                                    break;
                                }
                            }
                            if (!string.IsNullOrWhiteSpace(revisionDate))
                                break;
                        }
                    }

                    var data = new BoardData
                    {
                        RevisionDate = revisionDate,
                        Schematics = MapSchematics(ReadSheetRows(package, SheetBoardSchematics, SchematicsHeaders)),
                        Components = MapComponents(ReadSheetRows(package, SheetComponents, ComponentsHeaders)),
                        ComponentImages = MapComponentImages(ReadSheetRows(package, SheetComponentImages, ComponentImagesHeaders)),
                        ComponentHighlights = MapComponentHighlights(ReadSheetRows(package, SheetComponentHighlights, ComponentHighlightsHeaders)),
                        ComponentLocalFiles = MapComponentLocalFiles(ReadSheetRows(package, SheetComponentLocalFiles, ComponentLocalFilesHeaders)),
                        ComponentLinks = MapComponentLinks(ReadSheetRows(package, SheetComponentLinks, ComponentLinksHeaders)),
                        BoardLocalFiles = MapBoardLocalFiles(ReadSheetRows(package, SheetBoardLocalFiles, BoardLocalFilesHeaders)),
                        BoardLinks = MapBoardLinks(ReadSheetRows(package, SheetBoardLinks, BoardLinksHeaders)),
                        Credits = MapCredits(ReadSheetRows(package, SheetCredits, CreditsHeaders)),
                        IsLoaded = true
                    };

                    _cache[cacheKey] = data;
                    Logger.Info($"Board Excel data loaded and cached for [{cacheKey}]");
                    return (BoardData?)data;
                }
                catch (Exception ex)
                {
                    Logger.Critical($"Failed to load board Excel file [{cacheKey}] - [{ex.Message}]");
                    return null;
                }
            });
        }

        // ###########################################################################################
        // Scans the named sheet for the header row containing all required headers, then reads
        // all data rows below it. Each row is a case-insensitive dictionary keyed by header name.
        // Multi-line cell headers (Alt+Enter in Excel) are normalized to single-space strings.
        // ###########################################################################################
        private static List<Dictionary<string, string>> ReadSheetRows(ExcelPackage package, string sheetName, string[] requiredHeaders)
        {
            var rows = new List<Dictionary<string, string>>();
            var sheet = package.Workbook.Worksheets[sheetName];

            if (sheet == null)
            {
                Logger.Warning($"Sheet [{sheetName}] not found in board Excel file");
                return rows;
            }

            int maxRow = sheet.Dimension?.End.Row ?? 0;
            int maxCol = sheet.Dimension?.End.Column ?? 0;

            // Locate the header row by matching all required headers
            int headerRow = -1;
            var headerColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int row = 1; row <= maxRow; row++)
            {
                var colMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                for (int col = 1; col <= maxCol; col++)
                {
                    string text = NormalizeHeader(GetCellText(sheet, row, col));
                    if (!string.IsNullOrWhiteSpace(text))
                        colMap[text] = col;
                }

                if (requiredHeaders.All(h => colMap.ContainsKey(h)))
                {
                    headerRow = row;
                    headerColMap = colMap;
                    break;
                }
            }

            if (headerRow < 1)
            {
                Logger.Warning($"Header row not found in sheet [{sheetName}] - verify column header names");
                return rows;
            }

            // Read data rows
            for (int row = headerRow + 1; row <= maxRow; row++)
            {
                var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                bool hasData = false;

                foreach (var (header, col) in headerColMap)
                {
                    string value = GetCellText(sheet, row, col);
                    rowData[header] = value;
                    if (!string.IsNullOrWhiteSpace(value))
                        hasData = true;
                }

                if (hasData)
                    rows.Add(rowData);
            }

            return rows;
        }

        // ###########################################################################################
        // Typed mapping methods - convert raw string dictionaries to model instances.
        // ###########################################################################################
        private static string Val(Dictionary<string, string> row, string key)
            => row.TryGetValue(key, out var v) ? v : string.Empty;

        private static List<BoardSchematicEntry> MapSchematics(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardSchematicEntry
            {
                SchematicName = Val(r, ColSchematicName),
                SchematicImageFile = Val(r, ColSchematicImageFile),
                MainImageHighlightColor = Val(r, ColMainImageHighlightColor),
                MainHighlightOpacity = Val(r, ColMainHighlightOpacity),
                ThumbnailImageHighlightColor = Val(r, ColThumbnailImageHighlightColor),
                ThumbnailHighlightOpacity = Val(r, ColThumbnailHighlightOpacity)
            }).ToList();

        private static List<ComponentEntry> MapComponents(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentEntry
            {
                BoardLabel = Val(r, ColBoardLabel),
                FriendlyName = Val(r, ColFriendlyName),
                TechnicalNameOrValue = Val(r, ColTechnicalNameOrValue),
                PartNumber = Val(r, ColPartNumber),
                Category = Val(r, ColCategory),
                Region = Val(r, ColRegion),
                Description = Val(r, ColDescription)
            }).ToList();

        private static List<ComponentImageEntry> MapComponentImages(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentImageEntry
            {
                BoardLabel = Val(r, ColBoardLabel),
                Region = Val(r, ColRegion),
                Pin = Val(r, ColPin),
                Name = Val(r, ColName),
                ExpectedOscilloscopeReading = Val(r, ColExpectedOscilloscopeReading),
                File = Val(r, ColFile),
                Note = Val(r, ColNote),
                TimeDiv = Val(r, ColTimeDiv),
                VoltsDiv = Val(r, ColVoltsDiv),
                TriggerLevelVolts = Val(r, ColTriggerLevelVolts)
            }).ToList();

        private static List<ComponentHighlightEntry> MapComponentHighlights(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentHighlightEntry
            {
                SchematicName = Val(r, ColSchematicName),
                BoardLabel = Val(r, ColBoardLabel),
                X = Val(r, ColX),
                Y = Val(r, ColY),
                Width = Val(r, ColWidth),
                Height = Val(r, ColHeight)
            }).ToList();

        private static List<ComponentLocalFileEntry> MapComponentLocalFiles(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentLocalFileEntry
            {
                BoardLabel = Val(r, ColBoardLabel),
                Name = Val(r, ColName),
                File = Val(r, ColFile)
            }).ToList();

        private static List<ComponentLinkEntry> MapComponentLinks(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentLinkEntry
            {
                BoardLabel = Val(r, ColBoardLabel),
                Name = Val(r, ColName),
                Url = Val(r, ColUrl)
            }).ToList();

        private static List<BoardLocalFileEntry> MapBoardLocalFiles(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardLocalFileEntry
            {
                Category = Val(r, ColCategory),
                Name = Val(r, ColName),
                File = Val(r, ColFile)
            }).ToList();

        private static List<BoardLinkEntry> MapBoardLinks(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardLinkEntry
            {
                Category = Val(r, ColCategory),
                Name = Val(r, ColName),
                Url = Val(r, ColUrl)
            }).ToList();

        private static List<CreditEntry> MapCredits(List<Dictionary<string, string>> rows)
            => rows.Select(r => new CreditEntry
            {
                Category = Val(r, ColCategory),
                SubCategory = Val(r, ColSubCategory),
                NameOrHandle = Val(r, ColNameOrHandle),
                Contact = Val(r, ColContact)
            }).ToList();

        // ###########################################################################################
        // Collapses Alt+Enter line breaks in Excel cell headers into a single space, then trims.
        // ###########################################################################################
        private static string NormalizeHeader(string text)
        {
            //var parts = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            var parts = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); // compliant with .NET6
            return string.Join(" ", parts).Trim();
        }

        // ###########################################################################################
        // Returns the trimmed text value of a worksheet cell, or an empty string if null or blank.
        // ###########################################################################################
        private static string GetCellText(ExcelWorksheet sheet, int row, int col)
            => sheet.Cells[row, col].Text?.Trim() ?? string.Empty;
    }
}