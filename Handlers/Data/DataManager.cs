using CRT;
using Handlers.OnlineHandling;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    public static class DataManager
    {
        private const string DataRootArg = "--data-root=";
        private const string SheetHardwareBoard = "Hardware & Board";
        private const string SheetOscilloscope = "Oscilloscope";

        // Column header names used for robust, order-independent column mapping
        private const string ColHardwareName = "Hardware name in drop-down";
        private const string ColBoardName = "Board name in drop-down";
        private const string ColExcelDataFile = "Excel data file";
        private const string ColHardwareNotes = "Hardware notes in \"Overview\" tab";

        // Column headers for Oscilloscope
        private const string ColBrand = "Brand";
        private const string ColSeriesOrModel = "Series or model";
        private const string ColPort = "Port";
        private const string ColIdentify = "Identify";
        private const string ColDrainErrorQueue = "DrainErrorQueue";
        private const string ColOperationComplete = "Operation-Complete";
        private const string ColClearStatistics = "Clear-Statistics";
        private const string ColQueryActiveTrigger = "QueryActiveTrigger";
        private const string ColStop = "Stop";
        private const string ColSingle = "Single";
        private const string ColRun = "Run";
        private const string ColQueryTriggerMode = "QueryTriggerMode";
        private const string ColQueryTriggerLevel = "QueryTriggerLevel";
        private const string ColSetTriggerLevel = "SetTriggerLevel";
        private const string ColQueryTimeDiv = "QueryTimeDiv";
        private const string ColSetTimeDiv = "SetTimeDiv";
        private const string ColQueryVoltsDiv = "QueryVoltsDiv";
        private const string ColSetVoltsDiv = "SetVoltsDiv";
        private const string ColDumpImage = "DumpImage";
        private const string ColTimeDiv = "TIME/DIV";
        private const string ColVoltsDiv = "VOLTS/DIV";
        private const string ColDebounceTime = "Debounce-Time";

        private static string _dataRoot = string.Empty;
        private static List<DataFileEntry>? _syncManifest;

        public static string DataRoot => _dataRoot;
        //        public static List<HardwareBoardEntry> HardwareBoards { get; private set; } = [];
        public static List<HardwareBoardEntry> HardwareBoards { get; private set; } = new(); // compliant with .NET6
        public static List<OscilloscopeEntry> Oscilloscopes { get; private set; } = new();

        public static string ResolvedMainExcelFileName { get; private set; } = string.Empty;
        public static bool DataUpdateRequiresAppUpdate { get; private set; }

        // Raised with a general status message (e.g. "Checking files...", "Sync complete")
        public static event Action<string>? StatusChanged;

        // Raised with the relative file path of whichever file is currently being processed
        public static event Action<string>? FileDownloadChanged;

        // ###########################################################################################
        // Resolves the data root, ensures the folder exists, syncs all Excel files against the
        // online manifest, then loads hardware definitions. Images are left for SyncRemainingAsync.
        // ###########################################################################################
        public static async Task InitializeAsync(string[] args)
        {
            Logger.Info(args.Length > 0
                ? $"Commandline parameters: [{string.Join(" ", args)}]"
                : "No commandline parameters given");
            _dataRoot = ResolveDataRoot(args);
            Logger.Info($"Data root is [{_dataRoot}]");

            // Create the data root folder, if this is either first-run or using a custom --data-root
            bool isNewRoot = !Directory.Exists(_dataRoot);
            Directory.CreateDirectory(_dataRoot);
            if (isNewRoot)
            {
                // Check if this is a first-run with the default "AppData" path
                var bundledData = Path.Combine(AppContext.BaseDirectory, "Data");
                if (Directory.Exists(bundledData))
                {
                    Logger.Info("Data root folder created — seeding from install package");
                    RaiseStatus("Seeding data from install package...");
                    await Task.Run(() => CopyDirectory(bundledData, _dataRoot));
                    Logger.Info("Data seeded from install package");
                }
                else
                {
                    Logger.Info("Data root folder created — all files will be downloaded from online source");
                }
            }
            else if (UserSettings.CheckDataOnLaunch)
            {
                Logger.Info("Checking online source for new or updated files");
            }

            // Determine from local files initially (fallback if offline)
            var localFiles = Directory.EnumerateFiles(_dataRoot)
                .Select(Path.GetFileName)
                .Where(f => !string.IsNullOrEmpty(f))
                .Cast<string>();

            DetermineResolvedMainExcel(localFiles);

            // Preserve the local "newer version exists" state.
            // The online manifest may refine which file should be synced,
            // but it must not clear a newer version already detected locally.
            bool localDataUpdateRequiresAppUpdate = DataUpdateRequiresAppUpdate;

#if DEBUG
            if (!AppConfig.DebugSimulateSync)
            {
                Logger.Info("DEBUG build - skipping online sync");
            }
            else
            {
#endif
                if (UserSettings.CheckDataOnLaunch)
                {
                    RaiseStatus("Fetching online file manifest...");
                    _syncManifest = await OnlineServices.FetchManifestAsync(RaiseStatus);

                    if (_syncManifest != null)
                    {
                        // Use a dummy filter run to harvest the file names natively through the delegate 
                        // (ensures we get the exact string field the Sync tool operates on without guessing properties)
                        var onlineFileNames = new List<string>();
                        await OnlineServices.SyncFilesAsync(_syncManifest, string.Empty, f => { onlineFileNames.Add(f); return false; }, null, null);

                        DetermineResolvedMainExcel(onlineFileNames);

                        // Keep the local warning state even if the online manifest does not list
                        // the same newer main Excel file already present on disk.
                        DataUpdateRequiresAppUpdate |= localDataUpdateRequiresAppUpdate;

                        // Sync only the resolved versioned main Excel file first — one exact match
                        RaiseStatus("Checking main data file...");
                        await OnlineServices.SyncFilesAsync(
                            _syncManifest, _dataRoot,
                            f => string.Equals(f, ResolvedMainExcelFileName, StringComparison.OrdinalIgnoreCase),
                            RaiseStatus, RaiseFileDownload,
                            label: "Main Excel data file");
                    }
                }
                else
                {
                    Logger.Info("Online data sync skipped - disabled in settings");
                }
#if DEBUG
            }
#endif

            if (DataUpdateRequiresAppUpdate)
            {
                Logger.Warning("Newer main Excel data file versions exist, but an application update is required to utilize them");
                RaiseStatus("Data components outdated - application update required!");
            }

            Logger.Info($"Resolved main Excel data file to be used: [{ResolvedMainExcelFileName}]");

            // Load main Excel now — HardwareBoards is populated from this point on
            RaiseStatus("Loading hardware definitions...");
            await Task.Run(LoadMainExcel);

#if DEBUG
            if (AppConfig.DebugSimulateSync)
            {
#endif
                if (_syncManifest != null && HardwareBoards.Count > 0)
                {
                    // Build an exact set of board Excel paths from the loaded hardware list —
                    // this is a handful of files, not a broad extension scan of the full manifest
                    var boardExcelFiles = HardwareBoards
                        .Select(e => e.ExcelDataFile)
                        .Where(f => !string.IsNullOrWhiteSpace(f))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                    RaiseStatus("Checking board Excel data files...");
                    await OnlineServices.SyncFilesAsync(
                        _syncManifest, _dataRoot,
                        f => boardExcelFiles.Contains(f),
                        RaiseStatus, RaiseFileDownload,
                        label: "board Excel data files");
                }
#if DEBUG
            }
#endif
        }

        // Returns true when a background sync manifest is queued and ready to process.
        public static bool HasPendingSync => _syncManifest != null;

        // ###########################################################################################
        // Syncs all non-Excel files (images etc.) using the manifest already fetched at startup.
        // Intended to run silently in the background after the UI has opened.
        // onStatus: optional callback for general progress messages, fired on the caller's thread.
        // Returns the number of files that were successfully new or updated.
        // ###########################################################################################
        public static async Task<int> SyncRemainingAsync(Action<string>? onStatus = null)
        {
            if (_syncManifest == null)
                return 0;

            var manifest = _syncManifest;
            _syncManifest = null;

            return await OnlineServices.SyncFilesAsync(
                manifest, _dataRoot,
                f => !f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase),
                onStatus);
        }

        // ###########################################################################################
        // Parses available file lists to find the highest compatible version of the main Excel file.
        // ###########################################################################################
        private static void DetermineResolvedMainExcel(IEnumerable<string> availableFiles)
        {
            Version appVer = Version.TryParse(AppConfig.AppVersionString, out var v) ? v : new Version(0, 0, 0, 0);

            Version? bestVer = null;
            string bestFile = string.Empty;
            bool newerExists = false;

            foreach (var fullPath in availableFiles)
            {
                // Ensure any directory prefixes from manifest strings are ignored
                string file = Path.GetFileName(fullPath);

                if (file.StartsWith(AppConfig.MainExcelFileNamePrefix, StringComparison.OrdinalIgnoreCase) &&
                    file.EndsWith(AppConfig.MainExcelFileSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    string versionPart = file.Substring(
                        AppConfig.MainExcelFileNamePrefix.Length,
                        file.Length - AppConfig.MainExcelFileNamePrefix.Length - AppConfig.MainExcelFileSuffix.Length);

                    if (Version.TryParse(versionPart, out var fileVer))
                    {
                        if (fileVer <= appVer)
                        {
                            if (bestVer == null || fileVer > bestVer)
                            {
                                bestVer = fileVer;
                                bestFile = fullPath; // Preserve the original path for sync
                            }
                        }
                        else
                        {
                            newerExists = true;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(bestFile))
            {
                // Fallback looking for legacy unversioned file if no versioned ones exist
                bestFile = AppConfig.MainExcelFileName;
            }

            ResolvedMainExcelFileName = bestFile;
            DataUpdateRequiresAppUpdate = newerExists;
        }

        // ###########################################################################################
        // Recursively copies all files and subdirectories from source into destination.
        // ###########################################################################################
        private static void CopyDirectory(string source, string destination)
        {
            foreach (var file in Directory.EnumerateFiles(source, "*", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(source, file);
                var dest = Path.Combine(destination, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(dest)!);
                File.Copy(file, dest, overwrite: true);
            }
        }

        // ###########################################################################################
        // Parses --data-root from args, or falls back to a persistent AppData folder that survives
        // Velopack updates (which replace the install directory but leave AppData untouched).
        // ###########################################################################################
        private static string ResolveDataRoot(string[] args)
        {
            foreach (var arg in args)
            {
                if (arg.StartsWith(DataRootArg, StringComparison.OrdinalIgnoreCase))
                    return arg[DataRootArg.Length..].Trim('"', '\'');
            }

            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(appData, AppConfig.AppFolderName, "Data");
        }

        // ###########################################################################################
        // Reads hardware and board definitions from the main Excel file.
        // Column positions are resolved by header name so reordering columns is handled gracefully.
        // Hardware name is carried forward across rows where the cell is empty (merged cell pattern).
        // ###########################################################################################
        private static void LoadMainExcel()
        {
            if (string.IsNullOrEmpty(ResolvedMainExcelFileName))
            {
                Logger.Warning("No compatible main Excel data file matched the current application version");
                return;
            }

            var excelPath = Path.Combine(_dataRoot, ResolvedMainExcelFileName);

            if (!File.Exists(excelPath))
            {
                Logger.Warning($"Main Excel data file not found - [{excelPath}]");
                return;
            }

            ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);
                var sheet = package.Workbook.Worksheets[SheetHardwareBoard];

                if (sheet == null)
                {
                    Logger.Warning($"Sheet [{SheetHardwareBoard}] not found in main Excel data file");
                    return;
                }

                var hardwareRequiredCols = new[] { ColHardwareName, ColBoardName, ColExcelDataFile, ColHardwareNotes };
                var colMap = FindHeaderRow(sheet, hardwareRequiredCols, out int headerRow);

                if (colMap == null)
                {
                    Logger.Warning($"Header row not found in [{SheetHardwareBoard}] sheet - verify column header names match expected values");
                    return;
                }

                var entries = new List<HardwareBoardEntry>();
                int maxRow = sheet.Dimension?.End.Row ?? 0;
                string lastHwName = string.Empty;

                for (int row = headerRow + 1; row <= maxRow; row++)
                {
                    string hardwareName = GetCellText(sheet, row, colMap[ColHardwareName]);
                    string boardName = GetCellText(sheet, row, colMap[ColBoardName]);
                    string excelFile = GetCellText(sheet, row, colMap[ColExcelDataFile]);
                    string notes = GetCellText(sheet, row, colMap[ColHardwareNotes]);

                    // Carry forward hardware name for merged/empty cells in that column
                    if (!string.IsNullOrWhiteSpace(hardwareName))
                        lastHwName = hardwareName;
                    else
                        hardwareName = lastHwName;

                    // Skip rows that have neither a board name nor an Excel file reference
                    if (string.IsNullOrWhiteSpace(boardName) && string.IsNullOrWhiteSpace(excelFile))
                        continue;

                    entries.Add(new HardwareBoardEntry
                    {
                        HardwareName = hardwareName,
                        BoardName = boardName,
                        ExcelDataFile = excelFile,
                        HardwareNotes = notes
                    });
                }

                HardwareBoards = entries;
                Logger.Info($"Checking [{entries.Count}] board Excel data files, extracted from main Excel data file:");
                foreach (var boardEntry in entries)
                {
                    Logger.Info($"    [{boardEntry.ExcelDataFile}]");
                }

                var oscSheet = package.Workbook.Worksheets[SheetOscilloscope];
                if (oscSheet == null)
                {
                    Logger.Warning($"Sheet [{SheetOscilloscope}] not found in main Excel data file");
                }
                else
                {
                    var oscRequiredCols = new[] { ColBrand, ColSeriesOrModel, ColPort };
                    var oscColMap = FindHeaderRow(oscSheet, oscRequiredCols, out int oscHeaderRow);

                    if (oscColMap == null)
                    {
                        Logger.Warning($"Header row not found in [{SheetOscilloscope}] sheet - verify column header names match expected values");
                    }
                    else
                    {
                        var oscEntries = new List<OscilloscopeEntry>();
                        int oscMaxRow = oscSheet.Dimension?.End.Row ?? 0;
                        string lastBrandName = string.Empty;

                        for (int row = oscHeaderRow + 1; row <= oscMaxRow; row++)
                        {
                            string brand = GetCellTextSafe(oscSheet, row, oscColMap, ColBrand);
                            string series = GetCellTextSafe(oscSheet, row, oscColMap, ColSeriesOrModel);

                            if (!string.IsNullOrWhiteSpace(brand))
                                lastBrandName = brand;
                            else
                                brand = lastBrandName;

                            if (string.IsNullOrWhiteSpace(series))
                                continue;

                            oscEntries.Add(new OscilloscopeEntry
                            {
                                Brand = brand,
                                SeriesOrModel = series,
                                Port = GetCellTextSafe(oscSheet, row, oscColMap, ColPort),
                                Identify = GetCellTextSafe(oscSheet, row, oscColMap, ColIdentify),
                                DrainErrorQueue = GetCellTextSafe(oscSheet, row, oscColMap, ColDrainErrorQueue),
                                OperationComplete = GetCellTextSafe(oscSheet, row, oscColMap, ColOperationComplete),
                                ClearStatistics = GetCellTextSafe(oscSheet, row, oscColMap, ColClearStatistics),
                                QueryActiveTrigger = GetCellTextSafe(oscSheet, row, oscColMap, ColQueryActiveTrigger),
                                Stop = GetCellTextSafe(oscSheet, row, oscColMap, ColStop),
                                Single = GetCellTextSafe(oscSheet, row, oscColMap, ColSingle),
                                Run = GetCellTextSafe(oscSheet, row, oscColMap, ColRun),
                                QueryTriggerMode = GetCellTextSafe(oscSheet, row, oscColMap, ColQueryTriggerMode),
                                QueryTriggerLevel = GetCellTextSafe(oscSheet, row, oscColMap, ColQueryTriggerLevel),
                                SetTriggerLevel = GetCellTextSafe(oscSheet, row, oscColMap, ColSetTriggerLevel),
                                QueryTimeDiv = GetCellTextSafe(oscSheet, row, oscColMap, ColQueryTimeDiv),
                                SetTimeDiv = GetCellTextSafe(oscSheet, row, oscColMap, ColSetTimeDiv),
                                QueryVoltsDiv = GetCellTextSafe(oscSheet, row, oscColMap, ColQueryVoltsDiv),
                                SetVoltsDiv = GetCellTextSafe(oscSheet, row, oscColMap, ColSetVoltsDiv),
                                DumpImage = GetCellTextSafe(oscSheet, row, oscColMap, ColDumpImage),
                                TimeDivList = GetCellTextSafe(oscSheet, row, oscColMap, ColTimeDiv),
                                VoltsDivList = GetCellTextSafe(oscSheet, row, oscColMap, ColVoltsDiv),
                                DebounceTime = GetCellTextSafe(oscSheet, row, oscColMap, ColDebounceTime)
                            });
                        }

                        Oscilloscopes = oscEntries;
                        Logger.Info($"Loaded [{oscEntries.Count}] oscilloscope definitions from main Excel data file");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Failed to load main Excel data file - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Lazily loads and caches all sheets from the board-specific Excel file linked to the entry.
        // Delegates to BoardDataReader for parsing and caching. Returns null on failure.
        // ###########################################################################################
        public static async Task<BoardData?> LoadBoardDataAsync(HardwareBoardEntry entry)
        {
            var excelPath = Path.Combine(_dataRoot, entry.ExcelDataFile.Replace('/', Path.DirectorySeparatorChar));
            return await BoardDataReader.LoadAsync(excelPath, entry.ExcelDataFile);
        }

        // ###########################################################################################
        // Scans the worksheet for the first row containing all required column header names.
        // Matching is case-insensitive. Returns a header-name-to-column-index map on success,
        // or null when not all required headers are found.
        // ###########################################################################################
        private static Dictionary<string, int>? FindHeaderRow(ExcelWorksheet sheet, string[] required, out int headerRow)
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
                        colMap[text] = col;
                }

                if (required.All(h => colMap.ContainsKey(h)))
                {
                    headerRow = row;
                    return colMap;
                }
            }

            return null;
        }

        // ###########################################################################################
        // Collapses Alt+Enter line breaks in Excel cell headers into a single space, then trims.
        // ###########################################################################################
        private static string NormalizeHeader(string text)
        {
//            var parts = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
            var parts = text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries); // .NET6 compliant
            return string.Join(" ", parts).Trim();
        }

        // ###########################################################################################
        // Safely retrieves the trimmed text value of a worksheet cell if the column exists in the map.
        // ###########################################################################################
        private static string GetCellTextSafe(ExcelWorksheet sheet, int row, Dictionary<string, int> colMap, string columnName)
        {
            if (colMap.TryGetValue(columnName, out int col))
                return GetCellText(sheet, row, col);

            return string.Empty;
        }

        // ###########################################################################################
        // Returns the trimmed text value of a worksheet cell, or an empty string if null or blank.
        // ###########################################################################################
        private static string GetCellText(ExcelWorksheet sheet, int row, int col)
            => sheet.Cells[row, col].Text?.Trim() ?? string.Empty;

        // ###########################################################################################
        // Fires the StatusChanged event with the given message.
        // ###########################################################################################
        private static void RaiseStatus(string message) => StatusChanged?.Invoke(message);

        // ###########################################################################################
        // Fires the FileDownloadChanged event with the given file path.
        // ###########################################################################################
        private static void RaiseFileDownload(string filePath) => FileDownloadChanged?.Invoke(filePath);
    }
}