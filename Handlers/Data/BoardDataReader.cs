using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
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
        private const string SheetKiCadImportantSignals = "Important signals";

        // Board schematics columns
        private const string ColSchematicName = "Schematic name";
        private const string ColCadName = "CAD name";
        private const string ColSchematicImageFile = "Schematic image file";
        private const string ColSchematicHighlightColor = "Schematic highlight color";
        private const string ColSchematicHighlightOpacity = "Schematic highlight opacity";
        private const string ColOppositeTraceHighlightColor = "Opposite trace highlight color";
        private const string ColThumbnailHighlightColor = "Thumbnail highlight color";
        private const string ColThumbnailHighlightOpacity = "Thumbnail highlight opacity";

        // Shared columns
        private const string ColUuidV4 = "UUID v4";
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
        private const string ColTriggerLevelVolts = "T.LVL";

        // Component highlights columns
        private const string ColX = "X";
        private const string ColY = "Y";
        private const string ColWidth = "Width";
        private const string ColHeight = "Height";

        // Credits columns
        private const string ColSubCategory = "Sub-category";
        private const string ColNameOrHandle = "Name or handle";
        private const string ColContact = "Contact (email or web page)";

        // Important signals columns
        private const string ColDisplayName = "Display name";
        private const string ColKiCadNetName = "KiCad net name";

        private static readonly string[] SchematicsHeaders = new[] { ColSchematicName, ColSchematicImageFile };
        private static readonly string[] ComponentsHeaders = new[] { ColBoardLabel, ColFriendlyName, ColTechnicalNameOrValue };
        private static readonly string[] ComponentImagesHeaders = new[] { ColBoardLabel, ColPin, ColName, ColFile };
        private static readonly string[] ComponentHighlightsHeaders = new[] { ColSchematicName, ColBoardLabel, ColX, ColY };
        private static readonly string[] ComponentLocalFilesHeaders = new[] { ColBoardLabel, ColName, ColFile };
        private static readonly string[] ComponentLinksHeaders = new[] { ColBoardLabel, ColName, ColUrl };
        private static readonly string[] BoardLocalFilesHeaders = new[] { ColCategory, ColName, ColFile };
        private static readonly string[] BoardLinksHeaders = new[] { ColCategory, ColName, ColUrl };
        private static readonly string[] CreditsHeaders = new[] { ColCategory, ColNameOrHandle };
        private static readonly string[] KiCadImportantSignalsHeaders = new[] { ColDisplayName, ColKiCadNetName };

        // Concurrent on purpose: LoadAsync writes this from inside Task.Run while the UI thread
        // reads it, and DataValidator's startup walk loads every board from a background task in
        // parallel with the UI loading the selected one. With a plain Dictionary those races
        // corrupt the table, and the corruption surfaces as a board silently loading as null.
        // Two concurrent first-loads of the same key may both parse the file; the last write wins,
        // which is the same benign outcome the old code had on its lucky days.
        private static readonly ConcurrentDictionary<string, BoardData> _cache = new(StringComparer.OrdinalIgnoreCase);

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
                        Schematics = MapSchematics(ReadSheetRows(package, cacheKey, SheetBoardSchematics, SchematicsHeaders)),
                        Components = MapComponents(ReadSheetRows(package, cacheKey, SheetComponents, ComponentsHeaders)),
                        ComponentImages = MapComponentImages(ReadSheetRows(package, cacheKey, SheetComponentImages, ComponentImagesHeaders)),
                        ComponentHighlights = BoardComponentHighlightStorage.LoadComponentHighlights(excelPath),
                        ComponentLocalFiles = MapComponentLocalFiles(ReadSheetRows(package, cacheKey, SheetComponentLocalFiles, ComponentLocalFilesHeaders)),
                        ComponentLinks = MapComponentLinks(ReadSheetRows(package, cacheKey, SheetComponentLinks, ComponentLinksHeaders)),
                        BoardLocalFiles = MapBoardLocalFiles(ReadSheetRows(package, cacheKey, SheetBoardLocalFiles, BoardLocalFilesHeaders)),
                        BoardLinks = MapBoardLinks(ReadSheetRows(package, cacheKey, SheetBoardLinks, BoardLinksHeaders)),
                        Credits = MapCredits(ReadSheetRows(package, cacheKey, SheetCredits, CreditsHeaders)),
                        KiCadImportantSignals = MapKiCadImportantSignals(ReadSheetRows(package, cacheKey, SheetKiCadImportantSignals, KiCadImportantSignalsHeaders)),
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
        // Removes one cached board-data instance so the next load re-reads the Excel file from disk.
        // ###########################################################################################
        public static void ClearCache(string cacheKey)
        {
            if (string.IsNullOrWhiteSpace(cacheKey))
            {
                return;
            }

            BoardDataReader._cache.TryRemove(cacheKey, out _);
        }

        // ###########################################################################################
        // Clears all cached board-data instances so future loads re-read all board Excel files from disk.
        // ###########################################################################################
        public static void ClearAllCache()
        {
            BoardDataReader._cache.Clear();
        }

        // ###########################################################################################
        // Scans the named sheet for the header row containing all required headers, then reads
        // all data rows below it. Each row is a case-insensitive dictionary keyed by header name.
        // Multi-line cell headers (Alt+Enter in Excel) are normalized to single-space strings.
        // ###########################################################################################
        private static List<Dictionary<string, string>> ReadSheetRows(
    ExcelPackage package,
    string workbookDisplayPath,
    string sheetName,
    string[] requiredHeaders)
        {
            var rows = new List<Dictionary<string, string>>();
            var sheet = package.Workbook.Worksheets[sheetName];

            if (sheet == null)
            {
                Logger.Warning($"Sheet [{sheetName}] missing in [{workbookDisplayPath}]");
                return rows;
            }

            int maxRow = sheet.Dimension?.End.Row ?? 0;
            int maxCol = sheet.Dimension?.End.Column ?? 0;

            int headerRow = -1;
            var headerColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

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

                if (requiredHeaders.All(h => colMap.ContainsKey(h)))
                {
                    headerRow = row;
                    headerColMap = colMap;
                    break;
                }
            }

            if (headerRow < 1)
            {
                Logger.Warning($"Header row not found in sheet [{sheetName}] in [{workbookDisplayPath}] - verify column header names");
                return rows;
            }

            for (int row = headerRow + 1; row <= maxRow; row++)
            {
                var rowData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                bool hasData = false;

                foreach (var (header, col) in headerColMap)
                {
                    string value = GetCellText(sheet, row, col);
                    rowData[header] = value;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        hasData = true;
                    }
                }

                if (hasData)
                {
                    rows.Add(rowData);
                }
            }

            return rows;
        }

        // ###########################################################################################
        // Typed mapping methods - convert raw string dictionaries to model instances.
        // ###########################################################################################
        private static string Val(Dictionary<string, string> row, string key)
            => row.TryGetValue(key, out var v) ? v : string.Empty;

        // ###########################################################################################
        // Maps board schematic rows from Excel into the runtime schematic model.
        // KiCad calibration is no longer read from Excel and is instead loaded from the board JSON file.
        // Opposite trace highlight opacity is no longer loaded from Excel and now follows the main
        // schematic highlight opacity behavior in the renderer.
        // ###########################################################################################
        private static List<BoardSchematicEntry> MapSchematics(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardSchematicEntry
            {
                UuidV4 = Val(r, ColUuidV4),
                SchematicName = Val(r, ColSchematicName),
                CadName = Val(r, ColCadName),
                SchematicImageFile = Val(r, ColSchematicImageFile),
                SchematicHighlightColor = Val(r, ColSchematicHighlightColor),
                SchematicHighlightOpacity = Val(r, ColSchematicHighlightOpacity),
                OppositeTraceHighlightColor = Val(r, ColOppositeTraceHighlightColor),
                ThumbnailHighlightColor = Val(r, ColThumbnailHighlightColor),
                ThumbnailHighlightOpacity = Val(r, ColThumbnailHighlightOpacity)
            }).ToList();

        private static List<ComponentEntry> MapComponents(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentEntry
            {
                UuidV4 = Val(r, ColUuidV4),
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
                UuidV4 = Val(r, ColUuidV4),
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

        private static List<ComponentLocalFileEntry> MapComponentLocalFiles(List<Dictionary<string, string>> rows)
            => rows.Select(r => new ComponentLocalFileEntry
            {
                UuidV4 = Val(r, ColUuidV4),
                BoardLabel = Val(r, ColBoardLabel),
                Name = Val(r, ColName),
                File = Val(r, ColFile)
            }).ToList();

        private static List<ComponentLinkEntry> MapComponentLinks(List<Dictionary<string, string>> rows)
           => rows.Select(r => new ComponentLinkEntry
           {
               UuidV4 = Val(r, ColUuidV4),
               BoardLabel = Val(r, ColBoardLabel),
               Name = Val(r, ColName),
               Url = Val(r, ColUrl)
           }).ToList();

        private static List<BoardLocalFileEntry> MapBoardLocalFiles(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardLocalFileEntry
            {
                UuidV4 = Val(r, ColUuidV4),
                Category = Val(r, ColCategory),
                Name = Val(r, ColName),
                File = Val(r, ColFile)
            }).ToList();

        private static List<BoardLinkEntry> MapBoardLinks(List<Dictionary<string, string>> rows)
            => rows.Select(r => new BoardLinkEntry
            {
                UuidV4 = Val(r, ColUuidV4),
                Category = Val(r, ColCategory),
                Name = Val(r, ColName),
                Url = Val(r, ColUrl)
            }).ToList();

        private static List<CreditEntry> MapCredits(List<Dictionary<string, string>> rows)
            => rows.Select(r => new CreditEntry
            {
                UuidV4 = Val(r, ColUuidV4),
                Category = Val(r, ColCategory),
                SubCategory = Val(r, ColSubCategory),
                NameOrHandle = Val(r, ColNameOrHandle),
                Contact = Val(r, ColContact)
            }).ToList();

        // ###########################################################################################
        // Maps the grouped important KiCad signal references from the board Excel sheet.
        // ###########################################################################################
        private static List<KiCadImportantSignalEntry> MapKiCadImportantSignals(List<Dictionary<string, string>> rows)
            => rows
                .Select(r => new KiCadImportantSignalEntry
                {
                    DisplayName = Val(r, ColDisplayName),
                    KiCadNetName = Val(r, ColKiCadNetName)
                })
                .Where(entry =>
                    !string.IsNullOrWhiteSpace(entry.DisplayName) &&
                    !string.IsNullOrWhiteSpace(entry.KiCadNetName))
                .ToList();

        // ###########################################################################################
        // Collapses Alt+Enter line breaks in Excel cell headers into a single space, then trims.
        // ###########################################################################################
        private static string NormalizeHeader(string text)
        {
            var parts = text.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", parts).Trim();
        }

        // ###########################################################################################
        // Returns the trimmed text value of a worksheet cell, or an empty string if null or blank.
        // ###########################################################################################
        private static string GetCellText(ExcelWorksheet sheet, int row, int col)
            => sheet.Cells[row, col].Text?.Trim() ?? string.Empty;
    

        // ###########################################################################################
        // Collects all relative local file paths referenced by the board workbook so contributed
        // board assets can be protected from online overwrite.
        // ###########################################################################################
        public static HashSet<string> CollectReferencedLocalFiles(string excelPath)
        {
            TryCollectReferencedLocalFiles(excelPath, out var files);
            return files;
        }

        // ###########################################################################################
        // Same collection, but reports whether the workbook could actually be read. Orphan cleanup
        // uses this: a workbook that FAILS to read references an unknown set of files, which is not
        // the same as referencing none - treating it as empty would orphan that board's assets.
        // A missing workbook returns true with an empty set (an absent file references nothing).
        // ###########################################################################################
        public static bool TryCollectReferencedLocalFiles(string excelPath, out HashSet<string> files)
        {
            files = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return true;
            }

            ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);

                foreach (var row in ReadSheetRows(package, excelPath, SheetBoardSchematics, SchematicsHeaders))
                {
                    string file = thisNormalizeRelativePath(Val(row, ColSchematicImageFile));
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        files.Add(file);
                    }
                }

                foreach (var row in ReadSheetRows(package, excelPath, SheetComponentImages, ComponentImagesHeaders))
                {
                    string file = thisNormalizeRelativePath(Val(row, ColFile));
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        files.Add(file);
                    }
                }

                foreach (var row in ReadSheetRows(package, excelPath, SheetComponentLocalFiles, ComponentLocalFilesHeaders))
                {
                    string file = thisNormalizeRelativePath(Val(row, ColFile));
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        files.Add(file);
                    }
                }

                foreach (var row in ReadSheetRows(package, excelPath, SheetBoardLocalFiles, BoardLocalFilesHeaders))
                {
                    string file = thisNormalizeRelativePath(Val(row, ColFile));
                    if (!string.IsNullOrWhiteSpace(file))
                    {
                        files.Add(file);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to collect referenced local files from board Excel file [{excelPath}] - [{ex.Message}]");
                files.Clear();
                return false;
            }
        }

        // ###########################################################################################
        // Normalizes a relative file path for comparison with manifest entries.
        // ###########################################################################################
        private static string thisNormalizeRelativePath(string path)
        {
            return string.IsNullOrWhiteSpace(path)
                ? string.Empty
                : path.Trim().Replace('\\', '/').TrimStart('/');
        }





    }
}