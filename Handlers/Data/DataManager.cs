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
        private static List<DataFileEntry>? _lastFetchedManifest;
        private static readonly System.Threading.SemaphoreSlim OrphanAndUnusedFileCleanupSemaphore = new(1, 1);

        public static string DataRoot => _dataRoot;
        public static List<HardwareBoardEntry> HardwareBoards { get; private set; } = [];
        public static List<OscilloscopeEntry> Oscilloscopes { get; private set; } = new();

        public static string ResolvedMainExcelFileName { get; private set; } = string.Empty;
        public static bool DataUpdateRequiresAppUpdate { get; private set; }

        // Raised with a general status message (e.g. "Checking files...", "Sync complete")
        public static event Action<string>? StatusChanged;

        // Raised with the relative file path of whichever file is currently being processed
        public static event Action<string>? FileDownloadChanged;

        private static readonly HashSet<string> _protectedContributionFiles = new(StringComparer.OrdinalIgnoreCase);

        public static int ProtectedContributionFileCount => _protectedContributionFiles.Count;

        // ###########################################################################################
        // Result details returned from an immediate manual data sync.
        // ###########################################################################################
        public sealed class ManualDataSyncResult
        {
            public int ChangedCount { get; init; }
            public bool MainExcelChanged { get; init; }
            public bool AnyExcelChanged { get; init; }
            public int ProtectedFilesCount { get; init; }
        }

        private sealed class MainExcelResolution
        {
            public string FileName { get; init; } = string.Empty;
            public bool RequiresAppUpdate { get; init; }
            public string SourceDescription { get; init; } = string.Empty;
        }

        // ###########################################################################################
        // Resolves the best main Excel file from a file list, preferring the newest compatible
        // versioned file and optionally falling back to the legacy unversioned file name.
        // ###########################################################################################
        private static MainExcelResolution ResolveMainExcelCandidate(
            IEnumerable<string> availableFiles,
            bool allowSyntheticLegacyFallback,
            string sourceDescription)
        {
            Version appVersion = Version.TryParse(AppConfig.AppNumericVersionString, out var parsedVersion)
                ? parsedVersion
                : new Version(0, 0, 0, 0);

            Version? bestCompatibleVersion = null;
            string bestCompatibleFile = string.Empty;
            string legacyFile = string.Empty;
            bool newerExists = false;

            foreach (string rawPath in availableFiles ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(rawPath))
                {
                    continue;
                }

                string normalizedPath = thisNormalizeRelativePath(rawPath);
                string fileName = Path.GetFileName(normalizedPath);

                if (string.Equals(fileName, AppConfig.MainExcelFileName, StringComparison.OrdinalIgnoreCase))
                {
                    legacyFile = normalizedPath;
                }

                if (!fileName.StartsWith(AppConfig.MainExcelFileNamePrefix, StringComparison.OrdinalIgnoreCase) ||
                    !fileName.EndsWith(AppConfig.MainExcelFileSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string versionPart = fileName.Substring(
                    AppConfig.MainExcelFileNamePrefix.Length,
                    fileName.Length - AppConfig.MainExcelFileNamePrefix.Length - AppConfig.MainExcelFileSuffix.Length);

                if (!Version.TryParse(versionPart, out Version? fileVersion))
                {
                    continue;
                }

                if (fileVersion <= appVersion)
                {
                    if (bestCompatibleVersion == null || fileVersion > bestCompatibleVersion)
                    {
                        bestCompatibleVersion = fileVersion;
                        bestCompatibleFile = normalizedPath;
                    }
                }
                else
                {
                    newerExists = true;
                }
            }

            if (!string.IsNullOrWhiteSpace(bestCompatibleFile))
            {
                return new MainExcelResolution
                {
                    FileName = bestCompatibleFile,
                    RequiresAppUpdate = newerExists,
                    SourceDescription = sourceDescription
                };
            }

            if (!string.IsNullOrWhiteSpace(legacyFile))
            {
                return new MainExcelResolution
                {
                    FileName = legacyFile,
                    RequiresAppUpdate = newerExists,
                    SourceDescription = sourceDescription
                };
            }

            if (allowSyntheticLegacyFallback)
            {
                return new MainExcelResolution
                {
                    FileName = AppConfig.MainExcelFileName,
                    RequiresAppUpdate = newerExists,
                    SourceDescription = sourceDescription
                };
            }

            return new MainExcelResolution
            {
                FileName = string.Empty,
                RequiresAppUpdate = newerExists,
                SourceDescription = sourceDescription
            };
        }

        // ###########################################################################################
        // Applies the selected main Excel resolution to runtime state and emits a detailed log line.
        // ###########################################################################################
        private static void ApplyMainExcelResolution(MainExcelResolution resolution, string context)
        {
            ResolvedMainExcelFileName = resolution.FileName;
            DataUpdateRequiresAppUpdate = resolution.RequiresAppUpdate;

            Logger.Info(
                $"Main Excel resolution [{context}] source=[{resolution.SourceDescription}] file=[{(string.IsNullOrWhiteSpace(resolution.FileName) ? "none" : resolution.FileName)}] requires_app_update=[{resolution.RequiresAppUpdate}]");
        }

        // ###########################################################################################
        // Returns whether the given relative data file currently exists inside the data root.
        // ###########################################################################################
        private static bool DoesRelativeDataFileExist(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(_dataRoot) || string.IsNullOrWhiteSpace(relativePath))
            {
                return false;
            }

            string fullPath = Path.Combine(_dataRoot, relativePath.Replace('/', Path.DirectorySeparatorChar));
            return File.Exists(fullPath);
        }

        // ###########################################################################################
        // Ensures the selected main Excel exists locally after sync and falls back to the local
        // cached candidate only when the preferred online candidate could not be obtained.
        // ###########################################################################################
        private static MainExcelResolution EnsureUsableMainExcelAfterSync(
            MainExcelResolution preferredResolution,
            MainExcelResolution localFallbackResolution,
            string context,
            Action<string>? onStatus = null)
        {
            if (!string.IsNullOrWhiteSpace(preferredResolution.FileName) &&
                DoesRelativeDataFileExist(preferredResolution.FileName))
            {
                return preferredResolution;
            }

            if (!string.IsNullOrWhiteSpace(localFallbackResolution.FileName) &&
                DoesRelativeDataFileExist(localFallbackResolution.FileName) &&
                !string.Equals(localFallbackResolution.FileName, preferredResolution.FileName, StringComparison.OrdinalIgnoreCase))
            {
                Logger.Warning(
                    $"Resolved main Excel file [{preferredResolution.FileName}] was not available after sync - falling back to local cached [{localFallbackResolution.FileName}] in [{context}]");

                onStatus?.Invoke("Latest main data file could not be downloaded - using local cached data");

                return new MainExcelResolution
                {
                    FileName = localFallbackResolution.FileName,
                    RequiresAppUpdate = preferredResolution.RequiresAppUpdate || localFallbackResolution.RequiresAppUpdate,
                    SourceDescription = $"local fallback after failed online sync ({context})"
                };
            }

            Logger.Warning(
                $"Resolved main Excel file [{preferredResolution.FileName}] was not available after sync and no usable local fallback existed in [{context}]");

            return preferredResolution;
        }

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

            bool isNewRoot = !Directory.Exists(_dataRoot);
            Directory.CreateDirectory(_dataRoot);

            if (isNewRoot)
            {
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

            var localFiles = Directory.EnumerateFiles(_dataRoot)
                .Select(Path.GetFileName)
                .Where(file => !string.IsNullOrWhiteSpace(file))
                .Cast<string>()
                .ToList();

            MainExcelResolution localResolution = ResolveMainExcelCandidate(
                localFiles,
                allowSyntheticLegacyFallback: true,
                sourceDescription: "local cache");

            ApplyMainExcelResolution(localResolution, "startup local scan");

            // Skipping the online sync during development is the "Check for new or updated data at
            // application launch" setting - it short-circuits everything below, so there is nothing
            // build-specific here.
            if (UserSettings.CheckDataOnLaunch)
            {
                RaiseStatus("Fetching online file manifest...");
                _syncManifest = await OnlineServices.FetchManifestAsync(RaiseStatus);
                _lastFetchedManifest = _syncManifest?.ToList();

                if (_syncManifest != null)
                {
                    MainExcelResolution onlineResolution = ResolveMainExcelCandidate(
                        _syncManifest.Select(entry => entry.File),
                        allowSyntheticLegacyFallback: false,
                        sourceDescription: "online manifest");

                    Logger.Info(
                        $"Main Excel candidates at startup: local=[{localResolution.FileName}] online=[{onlineResolution.FileName}]");

                    MainExcelResolution selectedResolution = !string.IsNullOrWhiteSpace(onlineResolution.FileName)
                        ? new MainExcelResolution
                        {
                            FileName = onlineResolution.FileName,
                            RequiresAppUpdate = onlineResolution.RequiresAppUpdate || localResolution.RequiresAppUpdate,
                            SourceDescription = "online manifest preferred"
                        }
                        : localResolution;

                    ApplyMainExcelResolution(selectedResolution, "startup selected candidate");

                    if (!string.IsNullOrWhiteSpace(ResolvedMainExcelFileName))
                    {
                        RaiseStatus("Checking main data file...");
                        await OnlineServices.SyncFilesAsync(
                            _syncManifest,
                            _dataRoot,
                            file => string.Equals(
                                thisNormalizeRelativePath(file),
                                thisNormalizeRelativePath(ResolvedMainExcelFileName),
                                StringComparison.OrdinalIgnoreCase),
                            RaiseStatus,
                            RaiseFileDownload,
                            label: "Main Excel data file");
                    }

                    selectedResolution = EnsureUsableMainExcelAfterSync(
                        selectedResolution,
                        localResolution,
                        "startup",
                        RaiseStatus);

                    ApplyMainExcelResolution(selectedResolution, "startup final");
                }
            }
            else
            {
                Logger.Info("Online data sync skipped - disabled in settings");
            }

            if (DataUpdateRequiresAppUpdate)
            {
                Logger.Warning("Newer main Excel data file versions exist, but an application update is required to utilize them");
                RaiseStatus("Data components outdated - application update required!");
            }

            Logger.Info($"Resolved main Excel data file to be used: [{ResolvedMainExcelFileName}]");

            RaiseStatus("Loading hardware definitions...");
            await Task.Run(LoadMainExcel);

            // No separate guard for a skipped sync: _syncManifest is only ever assigned above, so it
            // is still null whenever the manifest fetch did not run.
            if (_syncManifest != null && HardwareBoards.Count > 0)
            {
                var boardExcelFiles = HardwareBoards
                    .Select(entry => thisNormalizeRelativePath(entry.ExcelDataFile))
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                RaiseStatus("Checking board Excel data files...");
                await OnlineServices.SyncFilesAsync(
                    _syncManifest,
                    _dataRoot,
                    file =>
                    {
                        string normalizedFile = thisNormalizeRelativePath(file);
                        return boardExcelFiles.Contains(normalizedFile) &&
                               !thisIsProtectedContributionFile(normalizedFile);
                    },
                    RaiseStatus,
                    RaiseFileDownload,
                    label: "board Excel data files");
            }
        }

        // Returns true when a background sync manifest is queued and ready to process.
        public static bool HasPendingSync => _syncManifest != null;

        // ###########################################################################################
        // Syncs all non-Excel files (images etc.) using the manifest already fetched at startup.
        // Intended to run silently in the background after the UI has opened.
        // onStatus: optional callback for general progress messages, fired on the caller's thread.
        // onFile:   optional callback fired with the relative file path currently being downloaded.
        // Returns the number of files that were successfully new or updated.
        // ###########################################################################################
        public static async Task<int> SyncRemainingAsync(Action<string>? onStatus = null, Action<string>? onFile = null)
        {
            if (_syncManifest == null)
                return 0;

            var manifest = _syncManifest;
            _syncManifest = null;

            return await OnlineServices.SyncFilesAsync(
                manifest,
                _dataRoot,
                file =>
                {
                    string normalizedFile = thisNormalizeRelativePath(file);
                    return !file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                           !thisIsProtectedContributionFile(normalizedFile);
                },
                onStatus,
                onFile,
                label: "remaining data files");
        }

        // ###########################################################################################
        // Performs an immediate manual data sync while the application is already running.
        // Startup behavior remains unchanged elsewhere; this manual path keeps a single UI sync flow,
        // refreshes the compatible main Excel first so the latest board Excel list is known, then syncs
        // board Excel files and non-Excel files, and finally clears cached board data if any Excel file changed.
        // Returns sync details so the UI can decide whether hardware/board selectors must be refreshed.
        // ###########################################################################################
        public static async Task<ManualDataSyncResult> CheckForDataUpdatesNowAsync(Action<string>? onStatus = null, Action<string>? onFile = null)
        {
            void ReportStatus(string message)
            {
                RaiseStatus(message);
                onStatus?.Invoke(message);
            }

            void ReportFile(string filePath)
            {
                RaiseFileDownload(filePath);
                onFile?.Invoke(filePath);
            }

            if (string.IsNullOrWhiteSpace(_dataRoot))
            {
                Logger.Warning("Immediate data update check skipped - data root is not initialized");
                ReportStatus("Sync failed - data folder is unavailable");

                return new ManualDataSyncResult
                {
                    ChangedCount = -1,
                    MainExcelChanged = false,
                    AnyExcelChanged = false,
                    ProtectedFilesCount = _protectedContributionFiles.Count
                };
            }

            var localFiles = Directory.EnumerateFiles(_dataRoot)
                .Select(Path.GetFileName)
                .Where(file => !string.IsNullOrWhiteSpace(file))
                .Cast<string>()
                .ToList();

            MainExcelResolution localResolution = ResolveMainExcelCandidate(
                localFiles,
                allowSyntheticLegacyFallback: true,
                sourceDescription: "local cache");

            ReportStatus("Fetching online file manifest...");
            _syncManifest = await OnlineServices.FetchManifestAsync(ReportStatus);
            _lastFetchedManifest = _syncManifest?.ToList();

            if (_syncManifest == null)
            {
                return new ManualDataSyncResult
                {
                    ChangedCount = -1,
                    MainExcelChanged = false,
                    AnyExcelChanged = false,
                    ProtectedFilesCount = _protectedContributionFiles.Count
                };
            }

            var manifest = _syncManifest;
            _syncManifest = null;

            MainExcelResolution onlineResolution = ResolveMainExcelCandidate(
                manifest.Select(entry => entry.File),
                allowSyntheticLegacyFallback: false,
                sourceDescription: "online manifest");

            Logger.Info(
                $"Main Excel candidates during manual sync: local=[{localResolution.FileName}] online=[{onlineResolution.FileName}]");

            MainExcelResolution selectedResolution = !string.IsNullOrWhiteSpace(onlineResolution.FileName)
                ? new MainExcelResolution
                {
                    FileName = onlineResolution.FileName,
                    RequiresAppUpdate = onlineResolution.RequiresAppUpdate || localResolution.RequiresAppUpdate,
                    SourceDescription = "online manifest preferred"
                }
                : localResolution;

            ApplyMainExcelResolution(selectedResolution, "manual sync selected candidate");

            int changedCount = 0;
            bool anyExcelChanged = false;

            if (!string.IsNullOrWhiteSpace(ResolvedMainExcelFileName))
            {
                ReportStatus("Checking main data file...");
                int mainExcelChangedCount = await OnlineServices.SyncFilesAsync(
                    manifest,
                    _dataRoot,
                    file => string.Equals(
                        thisNormalizeRelativePath(file),
                        thisNormalizeRelativePath(ResolvedMainExcelFileName),
                        StringComparison.OrdinalIgnoreCase),
                    ReportStatus,
                    ReportFile,
                    label: "Main Excel data file");

                changedCount += mainExcelChangedCount;
                bool mainExcelChanged = mainExcelChangedCount > 0;
                anyExcelChanged |= mainExcelChanged;
            }

            selectedResolution = EnsureUsableMainExcelAfterSync(
                selectedResolution,
                localResolution,
                "manual sync",
                ReportStatus);

            ApplyMainExcelResolution(selectedResolution, "manual sync final");

            ReportStatus("Loading hardware definitions...");
            await Task.Run(LoadMainExcel);

            var boardExcelFiles = HardwareBoards
                .Select(entry => thisNormalizeRelativePath(entry.ExcelDataFile))
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            ReportStatus($"Checking data from {AppConfig.GetOnlineSourceLabel()} - please wait...");
            int remainingChangedCount = await OnlineServices.SyncFilesAsync(
                manifest,
                _dataRoot,
                file =>
                {
                    string normalizedFile = thisNormalizeRelativePath(file);

                    return !string.Equals(
                               normalizedFile,
                               thisNormalizeRelativePath(ResolvedMainExcelFileName),
                               StringComparison.OrdinalIgnoreCase) &&
                           !thisIsProtectedContributionFile(normalizedFile) &&
                           (
                               boardExcelFiles.Contains(normalizedFile) ||
                               !file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                           );
                },
                ReportStatus,
                filePath =>
                {
                    if (!string.IsNullOrWhiteSpace(filePath) &&
                        filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        anyExcelChanged = true;
                    }

                    ReportFile(filePath);
                },
                label: "remaining data files");

            changedCount += remainingChangedCount;

            if (anyExcelChanged)
            {
                BoardDataReader.ClearAllCache();
            }

            bool finalMainExcelChanged = false;
            if (!string.IsNullOrWhiteSpace(ResolvedMainExcelFileName))
            {
                finalMainExcelChanged = changedCount > 0 &&
                    DoesRelativeDataFileExist(ResolvedMainExcelFileName);
            }

            return new ManualDataSyncResult
            {
                ChangedCount = changedCount,
                MainExcelChanged = finalMainExcelChanged,
                AnyExcelChanged = anyExcelChanged,
                ProtectedFilesCount = _protectedContributionFiles.Count
            };
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
        // Rebuilds the protected user-contribution file set for the currently resolved data files
        // without changing the already loaded runtime lists used by the UI.
        // ###########################################################################################
        public static void LoadProtectedContributionStateForCurrentData()
        {
            if (string.IsNullOrWhiteSpace(_dataRoot) || string.IsNullOrWhiteSpace(ResolvedMainExcelFileName))
            {
                _protectedContributionFiles.Clear();
                return;
            }

            string mainExcelPath = Path.Combine(_dataRoot, ResolvedMainExcelFileName);
            if (!File.Exists(mainExcelPath))
            {
                _protectedContributionFiles.Clear();
                return;
            }

            string userContributionRelativePath = thisGetUserContributionMainExcelRelativePath();
            if (string.IsNullOrWhiteSpace(userContributionRelativePath))
            {
                _protectedContributionFiles.Clear();
                return;
            }

            string userContributionFullPath = Path.Combine(_dataRoot, userContributionRelativePath.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(userContributionFullPath))
            {
                _protectedContributionFiles.Clear();
                return;
            }

            var protectedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            protectedFiles.Add(thisNormalizeRelativePath(userContributionRelativePath));

            // Protection building stays tolerant: an unreadable sidecar simply protects fewer
            // files here. The destructive path (orphan cleanup) checks the result and aborts.
            thisTryReadHardwareBoardEntriesFromWorkbook(userContributionFullPath, out var contributionEntries);

            foreach (var contributionEntry in contributionEntries)
            {
                string normalizedBoardExcelFile = thisNormalizeRelativePath(contributionEntry.ExcelDataFile);
                if (string.IsNullOrWhiteSpace(normalizedBoardExcelFile))
                {
                    continue;
                }

                protectedFiles.Add(normalizedBoardExcelFile);

                string boardExcelFullPath = Path.Combine(_dataRoot, normalizedBoardExcelFile.Replace('/', Path.DirectorySeparatorChar));
                string boardJsonRelativePath = thisNormalizeRelativePath(Path.ChangeExtension(normalizedBoardExcelFile, ".json") ?? string.Empty);

                if (!string.IsNullOrWhiteSpace(boardJsonRelativePath))
                {
                    protectedFiles.Add(boardJsonRelativePath);
                }

                if (!File.Exists(boardExcelFullPath))
                {
                    continue;
                }

                foreach (string referencedFile in BoardDataReader.CollectReferencedLocalFiles(boardExcelFullPath))
                {
                    string normalizedReferencedFile = thisNormalizeRelativePath(referencedFile);
                    if (!string.IsNullOrWhiteSpace(normalizedReferencedFile))
                    {
                        protectedFiles.Add(normalizedReferencedFile);
                    }
                }

                string boardDirectory = Path.GetDirectoryName(boardExcelFullPath) ?? string.Empty;
                string kiCadDirectory = Path.Combine(boardDirectory, "KiCad data");

                if (Directory.Exists(kiCadDirectory))
                {
                    foreach (string kiCadFile in Directory.EnumerateFiles(kiCadDirectory, "*", SearchOption.AllDirectories))
                    {
                        string relativeKiCadFile = Path.GetRelativePath(_dataRoot, kiCadFile);
                        string normalizedKiCadFile = thisNormalizeRelativePath(relativeKiCadFile);

                        if (!string.IsNullOrWhiteSpace(normalizedKiCadFile))
                        {
                            protectedFiles.Add(normalizedKiCadFile);
                        }
                    }
                }
            }

            _protectedContributionFiles.Clear();

            foreach (string protectedFile in protectedFiles)
            {
                _protectedContributionFiles.Add(protectedFile);
            }
        }

        // ###########################################################################################
        // Parses --data-root from args, or falls back to a persistent AppData folder that survives
        // Velopack updates (which replace the install directory but leave AppData untouched).
        // ###########################################################################################
        internal static string ResolveDataRoot(string[] args)
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
        // Points the data layer at an explicit data root and main workbook and loads the hardware,
        // board and oscilloscope definitions from it. This is the local half of InitializeAsync with
        // no online sync and no seeding - the test suite uses it to load a temporary data root
        // instead of the user's real one.
        // ###########################################################################################
        internal static void LoadFrom(string dataRoot, string mainExcelFileName)
        {
            _dataRoot = dataRoot;
            ResolvedMainExcelFileName = mainExcelFileName;
            LoadMainExcel();
        }

        // ###########################################################################################
        // Reads hardware, board, and oscilloscope definitions from the main Excel file.
        // Column positions are resolved by header name so reordering columns is handled gracefully.
        // Hardware name is carried forward across rows where the cell is empty (merged cell pattern).
        // Also loads an optional "_UserContribution" sidecar workbook and protects its files from sync.
        // ###########################################################################################
        private static void LoadMainExcel()
        {
            HardwareBoards = new List<HardwareBoardEntry>();
            Oscilloscopes = new List<OscilloscopeEntry>();
            _protectedContributionFiles.Clear();

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

                    if (!string.IsNullOrWhiteSpace(hardwareName))
                    {
                        lastHwName = hardwareName;
                    }
                    else
                    {
                        hardwareName = lastHwName;
                    }

                    if (string.IsNullOrWhiteSpace(boardName) && string.IsNullOrWhiteSpace(excelFile))
                    {
                        continue;
                    }

                    entries.Add(new HardwareBoardEntry
                    {
                        HardwareName = hardwareName,
                        BoardName = boardName,
                        ExcelDataFile = excelFile,
                        HardwareNotes = notes
                    });
                }

                string userContributionRelativePath = thisGetUserContributionMainExcelRelativePath();
                if (!string.IsNullOrWhiteSpace(userContributionRelativePath))
                {
                    string userContributionFullPath = Path.Combine(_dataRoot, userContributionRelativePath.Replace('/', Path.DirectorySeparatorChar));

                    if (File.Exists(userContributionFullPath))
                    {
                        thisTryReadHardwareBoardEntriesFromWorkbook(userContributionFullPath, out var contributionEntries);

                        if (contributionEntries.Count > 0)
                        {
                            var existingKeys = new HashSet<string>(
                                entries.Select(entry => $"{entry.HardwareName}|{entry.BoardName}"),
                                StringComparer.OrdinalIgnoreCase);

                            foreach (var contributionEntry in contributionEntries)
                            {
                                string identity = $"{contributionEntry.HardwareName}|{contributionEntry.BoardName}";
                                if (existingKeys.Add(identity))
                                {
                                    entries.Add(contributionEntry);
                                }
                                else
                                {
                                    Logger.Warning($"Duplicate hardware/board entry skipped from user contribution file: [{contributionEntry.HardwareName}] / [{contributionEntry.BoardName}]");
                                }
                            }

                            _protectedContributionFiles.Add(thisNormalizeRelativePath(userContributionRelativePath));

                            foreach (var contributionEntry in contributionEntries)
                            {
                                string normalizedBoardExcelFile = thisNormalizeRelativePath(contributionEntry.ExcelDataFile);
                                if (string.IsNullOrWhiteSpace(normalizedBoardExcelFile))
                                {
                                    continue;
                                }

                                _protectedContributionFiles.Add(normalizedBoardExcelFile);

                                string boardExcelFullPath = Path.Combine(_dataRoot, normalizedBoardExcelFile.Replace('/', Path.DirectorySeparatorChar));
                                string boardJsonRelativePath = thisNormalizeRelativePath(Path.ChangeExtension(normalizedBoardExcelFile, ".json") ?? string.Empty);
                                if (!string.IsNullOrWhiteSpace(boardJsonRelativePath))
                                {
                                    _protectedContributionFiles.Add(boardJsonRelativePath);
                                }

                                if (!File.Exists(boardExcelFullPath))
                                {
                                    Logger.Warning($"User contribution board Excel file not found while building protection list: [{boardExcelFullPath}]");
                                    continue;
                                }

                                foreach (string referencedFile in BoardDataReader.CollectReferencedLocalFiles(boardExcelFullPath))
                                {
                                    string normalizedReferencedFile = thisNormalizeRelativePath(referencedFile);
                                    if (!string.IsNullOrWhiteSpace(normalizedReferencedFile))
                                    {
                                        _protectedContributionFiles.Add(normalizedReferencedFile);
                                    }
                                }

                                string boardDirectory = Path.GetDirectoryName(boardExcelFullPath) ?? string.Empty;
                                string kiCadDirectory = Path.Combine(boardDirectory, "KiCad data");

                                if (Directory.Exists(kiCadDirectory))
                                {
                                    foreach (string kiCadFile in Directory.EnumerateFiles(kiCadDirectory, "*", SearchOption.AllDirectories))
                                    {
                                        string relativeKiCadFile = Path.GetRelativePath(_dataRoot, kiCadFile);
                                        string normalizedKiCadFile = thisNormalizeRelativePath(relativeKiCadFile);
                                        if (!string.IsNullOrWhiteSpace(normalizedKiCadFile))
                                        {
                                            _protectedContributionFiles.Add(normalizedKiCadFile);
                                        }
                                    }
                                }
                            }

                            Logger.Info($"Loaded [{contributionEntries.Count}] hardware/board entries from user contribution file [{userContributionRelativePath}]");
                            Logger.Info($"Protected contribution files=[{_protectedContributionFiles.Count}]");
                        }
                    }
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
                            {
                                lastBrandName = brand;
                            }
                            else
                            {
                                brand = lastBrandName;
                            }

                            if (string.IsNullOrWhiteSpace(series))
                            {
                                continue;
                            }

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
                Logger.Warning($"Failed to load main Excel data file - [{ex.Message}]");
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
            var parts = text.Split(['\r', '\n'], StringSplitOptions.RemoveEmptyEntries);
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

        // ###########################################################################################
        // Resolves the optional user contribution sidecar main Excel file for the active main workbook.
        // ###########################################################################################
        private static string thisGetUserContributionMainExcelRelativePath()
        {
            if (string.IsNullOrWhiteSpace(ResolvedMainExcelFileName))
            {
                return string.Empty;
            }

            string directory = Path.GetDirectoryName(ResolvedMainExcelFileName) ?? string.Empty;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(ResolvedMainExcelFileName);
            string extension = Path.GetExtension(ResolvedMainExcelFileName);

            string contributionFileName = $"{fileNameWithoutExtension}_UserContribution{extension}";
            string relativePath = string.IsNullOrWhiteSpace(directory)
                ? contributionFileName
                : Path.Combine(directory, contributionFileName);

            return thisNormalizeRelativePath(relativePath);
        }

        // ###########################################################################################
        // Reads hardware and board entries from a main-format Excel workbook. Returns false when the
        // workbook cannot be read or does not contain the expected sheet and headers - callers that
        // delete based on the result (orphan cleanup) must treat that as "unknown", not as "empty".
        // ###########################################################################################
        private static bool thisTryReadHardwareBoardEntriesFromWorkbook(string excelPath, out List<HardwareBoardEntry> entries)
        {
            entries = new List<HardwareBoardEntry>();

            ExcelPackage.License.SetNonCommercialPersonal("Dennis Helligsø");

            try
            {
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var package = new ExcelPackage(stream);
                var sheet = package.Workbook.Worksheets[SheetHardwareBoard];

                if (sheet == null)
                {
                    Logger.Warning($"Sheet [{SheetHardwareBoard}] not found in main-format Excel file [{excelPath}]");
                    return false;
                }

                var hardwareRequiredCols = new[] { ColHardwareName, ColBoardName, ColExcelDataFile, ColHardwareNotes };
                var colMap = FindHeaderRow(sheet, hardwareRequiredCols, out int headerRow);

                if (colMap == null)
                {
                    Logger.Warning($"Header row not found in [{SheetHardwareBoard}] sheet of [{excelPath}]");
                    return false;
                }

                int maxRow = sheet.Dimension?.End.Row ?? 0;
                string lastHwName = string.Empty;

                for (int row = headerRow + 1; row <= maxRow; row++)
                {
                    string hardwareName = GetCellText(sheet, row, colMap[ColHardwareName]);
                    string boardName = GetCellText(sheet, row, colMap[ColBoardName]);
                    string excelFile = GetCellText(sheet, row, colMap[ColExcelDataFile]);
                    string notes = GetCellText(sheet, row, colMap[ColHardwareNotes]);

                    if (!string.IsNullOrWhiteSpace(hardwareName))
                    {
                        lastHwName = hardwareName;
                    }
                    else
                    {
                        hardwareName = lastHwName;
                    }

                    if (string.IsNullOrWhiteSpace(boardName) && string.IsNullOrWhiteSpace(excelFile))
                    {
                        continue;
                    }

                    entries.Add(new HardwareBoardEntry
                    {
                        HardwareName = hardwareName,
                        BoardName = boardName,
                        ExcelDataFile = excelFile,
                        HardwareNotes = notes
                    });
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to read main-format Excel file [{excelPath}] - [{ex.Message}]");
                entries.Clear();
                return false;
            }
        }

        // ###########################################################################################
        // Normalizes a relative manifest or workbook path for comparison.
        // ###########################################################################################
        private static string thisNormalizeRelativePath(string path)
        {
            return string.IsNullOrWhiteSpace(path)
                ? string.Empty
                : path.Trim().Replace('\\', '/').TrimStart('/');
        }

        // ###########################################################################################
        // Returns whether a relative file path belongs to protected user contribution data.
        // ###########################################################################################
        private static bool thisIsProtectedContributionFile(string relativePath)
        {
            return !string.IsNullOrWhiteSpace(relativePath) &&
                   _protectedContributionFiles.Contains(thisNormalizeRelativePath(relativePath));
        }

        // ###########################################################################################
        // Deletes orphan and non-used files from the data root after building a complete referenced-file
        // map from the online manifest, all main Excel files, optional user-contribution main Excel
        // files, board Excel files, board JSON sidecars, and KiCad files. Deletions are strictly
        // contained inside the data root and only run while launch-time data sync is enabled.
        // Empty directories left behind are removed afterwards, still strictly within the data root.
        // FAIL CLOSED: this is the only code path that deletes user files, so it refuses to run when
        // the referenced-file map could be incomplete - no manifest snapshot (offline launch or a
        // failed fetch), a missing resolved main workbook, or any workbook that exists but cannot
        // be read. Every gap in the map would otherwise become a deletion.
        // ###########################################################################################
        public static async Task<int> DeleteOrphanAndUnusedFilesAsync()
        {
            if (!UserSettings.CheckDataOnLaunch)
            {
                Logger.Info("Orphan/non-used file cleanup skipped - launch-time data sync is disabled");
                return 0;
            }

            if (!UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
            {
                Logger.Info("Orphan/non-used file cleanup skipped - setting is disabled");
                return 0;
            }

            return await DeleteOrphanAndUnusedFilesAsync(_lastFetchedManifest);
        }

        // ###########################################################################################
        // The cleanup itself, with the manifest snapshot passed in explicitly. The public overload
        // above applies the user-setting gates and hands over the last fetched manifest; the test
        // suite calls this one directly so it can control the snapshot without touching UserSettings.
        // ###########################################################################################
        internal static async Task<int> DeleteOrphanAndUnusedFilesAsync(List<DataFileEntry>? manifestSnapshot)
        {
            if (string.IsNullOrWhiteSpace(_dataRoot) || !Directory.Exists(_dataRoot))
            {
                Logger.Warning("Orphan/non-used file cleanup skipped - data root is not available");
                return 0;
            }

            // Fail closed: with no manifest snapshot (offline launch, failed fetch) there is no
            // authority on which files the server provides, so "orphan" cannot be determined and
            // every unreferenced-but-legitimate file would be deleted.
            if (manifestSnapshot == null || manifestSnapshot.Count == 0)
            {
                Logger.Warning("Orphan/non-used file cleanup skipped - no online manifest snapshot is available (offline or failed sync), so orphans cannot be identified safely");
                return 0;
            }

            await OrphanAndUnusedFileCleanupSemaphore.WaitAsync();

            try
            {
                return await Task.Run(() =>
                {
                    string dataRootFullPath = thisEnsureTrailingDirectorySeparator(Path.GetFullPath(_dataRoot));
                    var mappedFiles = CollectAllReferencedDataFiles(dataRootFullPath, manifestSnapshot);

                    // Fail closed: an incomplete map means an unknown set of files would be
                    // wrongly classified as orphans, so nothing may be deleted.
                    if (mappedFiles == null)
                    {
                        Logger.Warning("Orphan/non-used file cleanup aborted - the referenced-file map could not be built completely, nothing was deleted");
                        return 0;
                    }

                    Logger.Info(
                        $"Starting orphan/non-used file cleanup inside data root [{dataRootFullPath}] with [{mappedFiles.Count}] mapped files");

                    int deletedFileCount = 0;

                    foreach (string filePath in Directory.EnumerateFiles(dataRootFullPath, "*", SearchOption.AllDirectories))
                    {
                        string fullFilePath = Path.GetFullPath(filePath);

                        if (!thisIsPathWithinDataRoot(dataRootFullPath, fullFilePath))
                        {
                            Logger.Warning($"Skipped file outside data root during orphan cleanup: [{fullFilePath}]");
                            continue;
                        }

                        string relativePath = thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, fullFilePath));

                        if (string.IsNullOrWhiteSpace(relativePath) || mappedFiles.Contains(relativePath))
                        {
                            continue;
                        }

                        try
                        {
                            File.Delete(fullFilePath);
                            deletedFileCount++;
                            Logger.Info($"Deleted orphan/non-used data file: [{relativePath}]");
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"Failed to delete orphan/non-used data file [{relativePath}] - [{ex.Message}]");
                        }
                    }

                    int deletedDirectoryCount = thisDeleteEmptyDirectoriesInsideDataRoot(dataRootFullPath);

                    Logger.Info(
                        $"Background orphan/non-used file cleanup complete - deleted [{deletedFileCount}] files and [{deletedDirectoryCount}] empty directories");

                    return deletedFileCount;
                });
            }
            finally
            {
                OrphanAndUnusedFileCleanupSemaphore.Release();
            }
        }

        // ###########################################################################################
        // Builds the full referenced-file map used by orphan cleanup across the online manifest,
        // all main workbooks, and all referenced board workbooks. Returns null when the map could
        // not be built completely (a workbook that exists but fails to read, or a missing resolved
        // main workbook) - the caller must then abort the cleanup rather than delete on a gap.
        // ###########################################################################################
        private static HashSet<string>? CollectAllReferencedDataFiles(string dataRootFullPath, List<DataFileEntry>? manifestSnapshot)
        {
            var mappedFiles = new HashSet<string>(
                OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                    ? StringComparer.OrdinalIgnoreCase
                    : StringComparer.Ordinal);

            thisAddMappedFilesFromManifestSnapshot(mappedFiles, dataRootFullPath, manifestSnapshot);

            var mainExcelRelativePaths = thisGetAllMainExcelRelativePaths(dataRootFullPath);

            Logger.Info(
                $"Orphan cleanup discovered [{mainExcelRelativePaths.Count}] main Excel data files: [{string.Join(" | ", mainExcelRelativePaths)}]");

            string resolvedMainExcelRelativePath = thisNormalizeRelativePath(ResolvedMainExcelFileName);

            if (!string.IsNullOrWhiteSpace(resolvedMainExcelRelativePath))
            {
                Logger.Info(
                    $"Orphan cleanup will map only resolved main Excel data file: [{resolvedMainExcelRelativePath}]");

                if (!thisTryAddMappedFilesFromMainExcel(mappedFiles, dataRootFullPath, resolvedMainExcelRelativePath))
                {
                    return null;
                }
            }
            else
            {
                Logger.Warning(
                    "Orphan cleanup did not have a resolved main Excel data file name; mapping all discovered main Excel files");

                foreach (string mainExcelRelativePath in mainExcelRelativePaths)
                {
                    if (!thisTryAddMappedFilesFromMainExcel(mappedFiles, dataRootFullPath, mainExcelRelativePath))
                    {
                        return null;
                    }
                }
            }

            Logger.Info(
                $"Mapped [{mappedFiles.Count}] referenced files from manifest and [{mainExcelRelativePaths.Count}] main Excel data files for orphan cleanup");

            return mappedFiles;
        }

        // ###########################################################################################
        // Adds all files from the latest fetched online manifest to the orphan-cleanup mapped-file set
        // so auxiliary reference files such as README files are never deleted just because they are not
        // directly referenced by Excel data.
        // ###########################################################################################
        private static void thisAddMappedFilesFromManifestSnapshot(
            HashSet<string> mappedFiles,
            string dataRootFullPath,
            List<DataFileEntry>? manifestSnapshot)
        {
            if (manifestSnapshot == null || manifestSnapshot.Count == 0)
            {
                Logger.Warning("No online manifest snapshot was available for orphan cleanup mapping");
                return;
            }

            int originalCount = mappedFiles.Count;

            foreach (DataFileEntry entry in manifestSnapshot)
            {
                thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, entry.File, "online manifest");
            }

            Logger.Info(
                $"Mapped [{mappedFiles.Count - originalCount}] files from the online manifest snapshot for orphan cleanup");
        }

        // ###########################################################################################
        // Returns all main Excel data files currently present inside the data root.
        // ###########################################################################################
        private static List<string> thisGetAllMainExcelRelativePaths(string dataRootFullPath)
        {
            return Directory.EnumerateFiles(dataRootFullPath, "*", SearchOption.AllDirectories)
                .Select(path => thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, path)))
                .Where(path =>
                {
                    string fileName = Path.GetFileName(path);

                    return string.Equals(fileName, AppConfig.MainExcelFileName, StringComparison.OrdinalIgnoreCase) ||
                           (
                               fileName.StartsWith(AppConfig.MainExcelFileNamePrefix, StringComparison.OrdinalIgnoreCase) &&
                               fileName.EndsWith(AppConfig.MainExcelFileSuffix, StringComparison.OrdinalIgnoreCase)
                           );
                })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ###########################################################################################
        // Adds all referenced files from one main Excel workbook and its optional user-contribution
        // sidecar workbook into the orphan-cleanup mapped-file set. Returns false when the map would
        // be incomplete: the main workbook is missing, or any workbook exists but cannot be read.
        // ###########################################################################################
        private static bool thisTryAddMappedFilesFromMainExcel(
            HashSet<string> mappedFiles,
            string dataRootFullPath,
            string mainExcelRelativePath)
        {
            thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, mainExcelRelativePath, "main Excel data file");

            string mainExcelFullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                mainExcelRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!File.Exists(mainExcelFullPath))
            {
                Logger.Warning($"Main Excel data file missing during orphan cleanup mapping: [{mainExcelRelativePath}]");
                return false;
            }

            if (!thisTryReadHardwareBoardEntriesFromWorkbook(mainExcelFullPath, out List<HardwareBoardEntry> entries))
            {
                return false;
            }

            foreach (HardwareBoardEntry entry in entries)
            {
                if (!thisTryAddMappedFilesFromBoardExcel(mappedFiles, dataRootFullPath, entry.ExcelDataFile))
                {
                    return false;
                }
            }

            string userContributionRelativePath = thisGetUserContributionMainExcelRelativePathFor(mainExcelRelativePath);
            string userContributionFullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                userContributionRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!File.Exists(userContributionFullPath))
            {
                return true;
            }

            thisTryAddMappedRelativePath(
                mappedFiles,
                dataRootFullPath,
                userContributionRelativePath,
                "user contribution main Excel data file");

            // The sidecar names the user's own local-only boards, so an unreadable sidecar means
            // exactly those unrestorable files would be misclassified as orphans - fail closed.
            if (!thisTryReadHardwareBoardEntriesFromWorkbook(userContributionFullPath, out List<HardwareBoardEntry> contributionEntries))
            {
                return false;
            }

            foreach (HardwareBoardEntry entry in contributionEntries)
            {
                if (!thisTryAddMappedFilesFromBoardExcel(mappedFiles, dataRootFullPath, entry.ExcelDataFile))
                {
                    return false;
                }
            }

            return true;
        }

        // ###########################################################################################
        // Adds one board Excel file and all files it uses to the orphan-cleanup mapped-file set.
        // Returns false when the board workbook exists but cannot be read - its reference set is
        // then unknown, which is not the same as empty. A missing board workbook is fine: it
        // references nothing, and its synced assets stay protected through the manifest snapshot.
        // ###########################################################################################
        private static bool thisTryAddMappedFilesFromBoardExcel(
            HashSet<string> mappedFiles,
            string dataRootFullPath,
            string boardExcelRelativePath)
        {
            string normalizedBoardExcelRelativePath = thisNormalizeRelativePath(boardExcelRelativePath);
            if (string.IsNullOrWhiteSpace(normalizedBoardExcelRelativePath))
            {
                return true;
            }

            thisTryAddMappedRelativePath(
                mappedFiles,
                dataRootFullPath,
                normalizedBoardExcelRelativePath,
                "board Excel data file");

            string boardJsonRelativePath =
                thisNormalizeRelativePath(Path.ChangeExtension(normalizedBoardExcelRelativePath, ".json") ?? string.Empty);

            if (!string.IsNullOrWhiteSpace(boardJsonRelativePath))
            {
                thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, boardJsonRelativePath, "board JSON sidecar");
            }

            string boardExcelFullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                normalizedBoardExcelRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!File.Exists(boardExcelFullPath))
            {
                return true;
            }

            if (!BoardDataReader.TryCollectReferencedLocalFiles(boardExcelFullPath, out HashSet<string> referencedFiles))
            {
                Logger.Warning($"Board Excel data file could not be read during orphan cleanup mapping: [{normalizedBoardExcelRelativePath}]");
                return false;
            }

            foreach (string referencedFile in referencedFiles)
            {
                thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, referencedFile, "board Excel referenced file");
            }

            string boardDirectory = Path.GetDirectoryName(boardExcelFullPath) ?? string.Empty;
            string kiCadDirectory = Path.Combine(boardDirectory, "KiCad data");

            if (!Directory.Exists(kiCadDirectory))
            {
                return true;
            }

            foreach (string kiCadFile in Directory.EnumerateFiles(kiCadDirectory, "*", SearchOption.AllDirectories))
            {
                string relativeKiCadFile = thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, kiCadFile));
                thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, relativeKiCadFile, "KiCad data file");
            }

            return true;
        }

        // ###########################################################################################
        // Adds a relative path to the orphan-cleanup mapped-file set only when it resolves inside the
        // current data root after full path normalization.
        // ###########################################################################################
        private static void thisTryAddMappedRelativePath(
            HashSet<string> mappedFiles,
            string dataRootFullPath,
            string relativePath,
            string sourceDescription)
        {
            string normalizedRelativePath = thisNormalizeRelativePath(relativePath);
            if (string.IsNullOrWhiteSpace(normalizedRelativePath))
            {
                return;
            }

            string fullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                normalizedRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!thisIsPathWithinDataRoot(dataRootFullPath, fullPath))
            {
                Logger.Warning(
                    $"Skipped mapped file outside data root from [{sourceDescription}]: [{relativePath}]");
                return;
            }

            string safeRelativePath = thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, fullPath));
            mappedFiles.Add(safeRelativePath);
        }

        // ###########################################################################################
        // Resolves the optional user-contribution sidecar workbook path for an arbitrary main workbook.
        // ###########################################################################################
        private static string thisGetUserContributionMainExcelRelativePathFor(string mainExcelRelativePath)
        {
            if (string.IsNullOrWhiteSpace(mainExcelRelativePath))
            {
                return string.Empty;
            }

            string directory = Path.GetDirectoryName(mainExcelRelativePath) ?? string.Empty;
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(mainExcelRelativePath);
            string extension = Path.GetExtension(mainExcelRelativePath);

            string contributionFileName = $"{fileNameWithoutExtension}_UserContribution{extension}";
            string relativePath = string.IsNullOrWhiteSpace(directory)
                ? contributionFileName
                : Path.Combine(directory, contributionFileName);

            return thisNormalizeRelativePath(relativePath);
        }

        // ###########################################################################################
        // Returns whether a fully normalized path is located inside the current data root.
        // ###########################################################################################
        private static bool thisIsPathWithinDataRoot(string dataRootFullPath, string candidateFullPath)
        {
            StringComparison comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;

            return candidateFullPath.StartsWith(dataRootFullPath, comparison);
        }

        // ###########################################################################################
        // Ensures a directory path ends with a single trailing directory separator for safe prefix tests.
        // ###########################################################################################
        private static string thisEnsureTrailingDirectorySeparator(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            return path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? path
                : path + Path.DirectorySeparatorChar;
        }

        // ###########################################################################################
        // Deletes empty directories inside the data root after orphan file cleanup.
        // Traversal is deepest-first so child folders are removed before their parents.
        // ###########################################################################################
        private static int thisDeleteEmptyDirectoriesInsideDataRoot(string dataRootFullPath)
        {
            int deletedDirectoryCount = 0;

            var directories = Directory.EnumerateDirectories(dataRootFullPath, "*", SearchOption.AllDirectories)
                .Select(Path.GetFullPath)
                .Where(path => thisIsPathWithinDataRoot(dataRootFullPath, path))
                .OrderByDescending(path => path.Length)
                .ToList();

            foreach (string directoryPath in directories)
            {
                try
                {
                    if (!Directory.Exists(directoryPath))
                    {
                        continue;
                    }

                    if (Directory.EnumerateFileSystemEntries(directoryPath).Any())
                    {
                        continue;
                    }

                    string relativePath = thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, directoryPath));
                    Directory.Delete(directoryPath, recursive: false);
                    deletedDirectoryCount++;
                    Logger.Info($"Deleted empty data directory: [{relativePath}]");
                }
                catch (Exception ex)
                {
                    string relativePath = thisNormalizeRelativePath(Path.GetRelativePath(dataRootFullPath, directoryPath));
                    Logger.Warning($"Failed to delete empty data directory [{relativePath}] - [{ex.Message}]");
                }
            }

            return deletedDirectoryCount;
        }





    }
}