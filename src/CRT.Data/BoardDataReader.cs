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
        // ###########################################################################################
        // The workbook's shape lives in BoardWorkbookSchema, NOT here. These are aliases onto it,
        // kept so the mapping code below reads as it always did.
        //
        // It moved when Phase 5 added a WRITER (publishing): two copies of a column name, one per
        // direction, is the defect that ComponentContributionPayload had and that
        // SubmissionContract.cs was created to fix. A workbook written "Part number" and read
        // "Part-number" loses every part number silently - nothing throws, the data is just gone.
        // ###########################################################################################

        // Sheet names
        private const string SheetBoardSchematics = BoardWorkbookSchema.SheetBoardSchematics;
        private const string SheetComponents = BoardWorkbookSchema.SheetComponents;
        private const string SheetComponentImages = BoardWorkbookSchema.SheetComponentImages;
        private const string SheetComponentLocalFiles = BoardWorkbookSchema.SheetComponentLocalFiles;
        private const string SheetComponentLinks = BoardWorkbookSchema.SheetComponentLinks;
        private const string SheetBoardLocalFiles = BoardWorkbookSchema.SheetBoardLocalFiles;
        private const string SheetBoardLinks = BoardWorkbookSchema.SheetBoardLinks;
        private const string SheetCredits = BoardWorkbookSchema.SheetCredits;
        private const string SheetKiCadImportantSignals = BoardWorkbookSchema.SheetKiCadImportantSignals;

        // The two columns CollectReferencedLocalFiles reads directly. Every other column is read
        // through BoardWorkbookSchema's row mappers, which moved there on 2026-09-24 so the Drafts
        // tab's table editor could save through the same mapping this class loads through.
        private const string ColSchematicImageFile = BoardWorkbookSchema.ColSchematicImageFile;
        private const string ColFile = BoardWorkbookSchema.ColFile;

        private static string Val(Dictionary<string, string> row, string key)
            => row.TryGetValue(key, out var v) ? v : string.Empty;

        // The required-header subsets, also from the schema - these are what ReadSheetRows scans
        // for to find a header row, so they are the names that must never drift.
        private static readonly string[] SchematicsHeaders = [.. BoardWorkbookSchema.BoardSchematics.RequiredHeaders];
        private static readonly string[] ComponentsHeaders = [.. BoardWorkbookSchema.Components.RequiredHeaders];
        private static readonly string[] ComponentImagesHeaders = [.. BoardWorkbookSchema.ComponentImages.RequiredHeaders];
        private static readonly string[] ComponentLocalFilesHeaders = [.. BoardWorkbookSchema.ComponentLocalFiles.RequiredHeaders];
        private static readonly string[] ComponentLinksHeaders = [.. BoardWorkbookSchema.ComponentLinks.RequiredHeaders];
        private static readonly string[] BoardLocalFilesHeaders = [.. BoardWorkbookSchema.BoardLocalFiles.RequiredHeaders];
        private static readonly string[] BoardLinksHeaders = [.. BoardWorkbookSchema.BoardLinks.RequiredHeaders];
        private static readonly string[] CreditsHeaders = [.. BoardWorkbookSchema.Credits.RequiredHeaders];
        private static readonly string[] KiCadImportantSignalsHeaders = [.. BoardWorkbookSchema.KiCadImportantSignals.RequiredHeaders];

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
        //
        // draft, when given, is a local BoardDraft (NewContributeStrategy.md Phase 2) overlaid onto
        // the OFFICIAL result via BoardDraftApplier.ApplyDraft before it is returned. Deliberately
        // applied here, on the way out, rather than folded into what gets cached: _cache must only
        // ever hold the officially synced data, or (a) a draft edit would need to invalidate the
        // cache to take effect, and (b) the cache would need a second key per draft state, both far
        // more complex than re-applying a cheap in-memory merge on every call. The Excel read stays
        // a pure "what does the official file say" question; drafting is a concern layered on top.
        //
        // allowMissingOfficialFile is for a system that exists ONLY as a local draft (session 2c,
        // task 9) - it has no .xlsx under Data/ and never will, so the ordinary "file not found"
        // answer of null would make the board unopenable. The caller passes true ONLY when it has
        // established that from the draft's own NewSystemRegistration; "the file happens to be
        // missing" is deliberately NOT the trigger, because for a system the main workbook DOES
        // list, a missing file is a genuine sync failure that must keep logging and returning null
        // rather than silently rendering as an empty board.
        // ###########################################################################################
        public static async Task<BoardData?> LoadAsync(
            string excelPath,
            string cacheKey,
            bool allowMissingOfficialFile = false)
        {
            if (_cache.TryGetValue(cacheKey, out var cached))
                return cached;

            if (!File.Exists(excelPath))
            {
                if (allowMissingOfficialFile)
                {
                    // A system that exists only as a local draft has no published workbook by
                    // construction, and an empty board is the truthful answer for it: officially,
                    // none of it exists yet. Its own rows come from the DRAFT workbook, which the
                    // caller resolves separately (DraftBoardSource).
                    //
                    // Nothing is cached: there is no file behind this, so caching it would pin an
                    // empty board against a path that may later hold a real one.
                    return new BoardData();
                }

                CrtLog.Warning($"Board Excel file not found: [{excelPath}]");
                return null;
            }

            return await Task.Run(() =>
            {
                BoardData? data = ParseWorkbook(excelPath, cacheKey);
                if (data is null)
                {
                    return null;
                }

                _cache[cacheKey] = data;
                CrtLog.Info($"Board Excel data loaded and cached for [{cacheKey}]");

                return data;
            });
        }

        // ###########################################################################################
        // Reads a workbook from disk WITHOUT consulting or populating the cache
        // (NewContributeStrategy.md Phase 6).
        //
        // *** THIS EXISTS FOR READ-MODIFY-WRITE ON A DRAFT. *** DraftWorkbookStore has to see what
        // is on disk right now, including an edit the contributor made in Excel while the
        // application was running - a cached read there would silently write back a board built on
        // a stale copy, destroying that edit. It also must not POPULATE the cache, or the
        // pre-modification state would become what the next board load displays.
        //
        // Shares ParseWorkbook with LoadAsync, so there is one parser rather than two that can
        // drift apart on a schema change.
        // ###########################################################################################
        public static BoardData? ReadWorkbookUncached(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return null;
            }

            return ParseWorkbook(excelPath, excelPath);
        }

        // ###########################################################################################
        // The parse itself, shared by the cached and uncached paths. Caches nothing and applies no
        // draft - both are the caller's business.
        //
        // `logKey` names the file in log lines only; sheet reads take it for the same reason.
        // ###########################################################################################
        private static BoardData? ParseWorkbook(string excelPath, string logKey)
        {
            EpplusLicense.Ensure();

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);

                return new BoardData
                {
                    RevisionDate = ScanRevisionDate(package),
                    HardwareName = ScanPreambleValue(package, BoardWorkbookSchema.HardwareMarkerPrefix),
                    BoardName = ScanPreambleValue(package, BoardWorkbookSchema.BoardMarkerPrefix),
                    Schematics = BoardWorkbookSchema.MapSchematics(ReadSheetRows(package, logKey, SheetBoardSchematics, SchematicsHeaders)),
                    Components = BoardWorkbookSchema.MapComponents(ReadSheetRows(package, logKey, SheetComponents, ComponentsHeaders)),
                    ComponentImages = BoardWorkbookSchema.MapComponentImages(ReadSheetRows(package, logKey, SheetComponentImages, ComponentImagesHeaders)),
                    ComponentHighlights = BoardComponentHighlightStorage.LoadComponentHighlights(excelPath),
                    ComponentLocalFiles = BoardWorkbookSchema.MapComponentLocalFiles(ReadSheetRows(package, logKey, SheetComponentLocalFiles, ComponentLocalFilesHeaders)),
                    ComponentLinks = BoardWorkbookSchema.MapComponentLinks(ReadSheetRows(package, logKey, SheetComponentLinks, ComponentLinksHeaders)),
                    BoardLocalFiles = BoardWorkbookSchema.MapBoardLocalFiles(ReadSheetRows(package, logKey, SheetBoardLocalFiles, BoardLocalFilesHeaders)),
                    BoardLinks = BoardWorkbookSchema.MapBoardLinks(ReadSheetRows(package, logKey, SheetBoardLinks, BoardLinksHeaders)),
                    Credits = BoardWorkbookSchema.MapCredits(ReadSheetRows(package, logKey, SheetCredits, CreditsHeaders)),
                    KiCadImportantSignals = BoardWorkbookSchema.MapKiCadImportantSignals(ReadSheetRows(package, logKey, SheetKiCadImportantSignals, KiCadImportantSignalsHeaders)),
                };
            }
            catch (Exception ex)
            {
                CrtLog.Critical($"Failed to load board Excel file [{logKey}] - [{ex.Message}]");
                return null;
            }
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
                CrtLog.Warning($"Sheet [{sheetName}] missing in [{workbookDisplayPath}]");
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
                CrtLog.Warning($"Header row not found in sheet [{sheetName}] in [{workbookDisplayPath}] - verify column header names");
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
        // A preamble value from the "Board schematics" sheet - "# Hardware: Commodore 64" yields
        // "Commodore 64" (added 2026-09-23).
        //
        // The same top-left 10x10 scan ScanRevisionDate uses, and for the same reason: the block
        // is hand-maintained, so a value may sit a row or two from where a generated file puts it.
        // Presentation only; nothing keys off the result.
        // ###########################################################################################
        private static string ScanPreambleValue(ExcelPackage package, string markerPrefix)
        {
            var schematicsSheet = package.Workbook.Worksheets[SheetBoardSchematics];
            if (schematicsSheet == null)
            {
                return string.Empty;
            }

            int limitRow = Math.Min(10, schematicsSheet.Dimension?.End.Row ?? 0);
            int limitCol = Math.Min(10, schematicsSheet.Dimension?.End.Column ?? 0);

            for (int r = 1; r <= limitRow; r++)
            {
                for (int c = 1; c <= limitCol; c++)
                {
                    string text = GetCellText(schematicsSheet, r, c);
                    if (text.StartsWith(markerPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        return text.Substring(markerPrefix.Length).Trim();
                    }
                }
            }

            return string.Empty;
        }

        // ###########################################################################################
        // The board's revision date, as the ONE place that knows how to find it: the first cell in
        // the top-left 10x10 of the "Board schematics" sheet whose text starts "# Revision date:",
        // with the rest of that cell taken and trimmed.
        //
        // Extracted so LoadAsync and ReadRevisionDateOnly cannot drift apart - if they ever
        // disagreed, the drift check would compare a value against itself read a different way,
        // which is the one thing that must not happen here. BoardDataReaderTests pins them equal.
        //
        // The value is hand-maintained free text (real boards say "2026-August-21"), so nothing is
        // validated or reformatted - see DraftRevisionComparer for how it is interpreted.
        // ###########################################################################################
        private static string ScanRevisionDate(ExcelPackage package)
        {
            var schematicsSheet = package.Workbook.Worksheets[SheetBoardSchematics];
            if (schematicsSheet == null)
            {
                return string.Empty;
            }

            int limitRow = Math.Min(10, schematicsSheet.Dimension?.End.Row ?? 0);
            int limitCol = Math.Min(10, schematicsSheet.Dimension?.End.Column ?? 0);

            for (int r = 1; r <= limitRow; r++)
            {
                for (int c = 1; c <= limitCol; c++)
                {
                    string text = GetCellText(schematicsSheet, r, c);
                    if (text.StartsWith(BoardWorkbookSchema.RevisionDateMarkerPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        return text.Substring(BoardWorkbookSchema.RevisionDateMarkerPrefix.Length).Trim();
                    }
                }
            }

            return string.Empty;
        }

        // ###########################################################################################
        // The board's revision date WITHOUT parsing the rest of the workbook - for the drift check
        // (NewContributeStrategy.md Phase 2, session 2d), which needs this one value per drafted
        // system and has no use for its ten sheets or its JSON sidecar.
        //
        // Deliberately does NOT touch _cache: it produces no BoardData, and a board opened for its
        // revision alone must not look "loaded" to anything else. Returns empty for a missing or
        // unreadable file, so a board that cannot be read reports no drift rather than throwing
        // inside a UI refresh.
        //
        // Try TryGetCachedRevisionDate first - a board already loaded (always true of the selected
        // one) needs no disk read at all.
        // ###########################################################################################
        public static string ReadRevisionDateOnly(string excelPath)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                return string.Empty;
            }

            EpplusLicense.Ensure();

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);

                return ScanRevisionDate(package);
            }
            catch (Exception ex)
            {
                CrtLog.Warning($"Could not read the revision date from [{excelPath}] - [{ex.Message}]");
                return string.Empty;
            }
        }

        // ###########################################################################################
        // The revision date of a board that is ALREADY loaded, or null when it is not cached. Lets
        // the drift check skip the disk entirely for the board on screen, which is the common case.
        // ###########################################################################################
        public static string? TryGetCachedRevisionDate(string cacheKey) =>
            _cache.TryGetValue(cacheKey, out var cached) ? cached.RevisionDate : null;



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

            EpplusLicense.Ensure();

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
                CrtLog.Warning($"Failed to collect referenced local files from board Excel file [{excelPath}] - [{ex.Message}]");
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