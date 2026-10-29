using CRT;
using Handlers.OnlineHandling;
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
        private static string _dataRoot = string.Empty;

        // CRT 3.x's Excel data folder at a real launch (FindLegacyDataRoot), never in a test (LoadFrom).
        private static string _legacyDataRoot = string.Empty;

        // The user's own oscilloscope models (OwnOscilloscopeFile) at a real launch, never in a test.
        private static string _ownOscilloscopesPath = string.Empty;

        private static List<DataFileEntry>? _syncManifest;
        private static List<DataFileEntry>? _lastFetchedManifest;
        private static readonly System.Threading.SemaphoreSlim OrphanAndUnusedFileCleanupSemaphore = new(1, 1);

        public static string DataRoot => _dataRoot;
        public static List<HardwareBoardEntry> HardwareBoards { get; private set; } = [];
        public static List<OscilloscopeEntry> Oscilloscopes { get; private set; } = new();

        // The models Oscilloscopes is made of, as rows of the main file: the published ones, then the
        // user's own (OwnOscilloscopeFile) - what the Oscilloscope tab's model editor starts from and
        // changes. Each list is replaced whole, never changed in place.
        public static IReadOnlyList<MainDocument.OscilloscopeRow> PublishedScopeModels { get; private set; } = [];

        public static IReadOnlyList<MainDocument.OscilloscopeRow> OwnScopeModels { get; private set; } = [];

        // One read or write of the own models' file at a time, the latest list last.
        private static readonly System.Threading.SemaphoreSlim OwnScopeModelsWriteGate = new(1, 1);

        // The two lists and the three facts below change together under this - the load on the pool,
        // the editor on the UI thread.
        private static readonly object OwnScopeModelsLock = new();

        // The own models' file the list held was read from - null until a read went well. A reload of
        // the same file keeps the list held (LoadOwnScopeModels).
        private static string? _ownScopeModelsReadFrom;

        // Why the user's own models must not be changed this run - their file could not be read at
        // launch (OwnOscilloscopeFile.LoadOrFind) - or null. The model editor says it and saves nothing.
        public static string? OwnScopeModelsProblem { get; private set; }

        public static string ResolvedMainExcelFileName { get; private set; } = string.Empty;
        public static bool DataUpdateRequiresAppUpdate { get; private set; }

        // Which rows the most recent LoadBoardDataAsync call overlaid from a local draft - the
        // cheap "is this row drafted" answer the UI checks to draw a chip or a tint (see
        // BoardDraftSummary's own header comment for why this is not derived from the merged
        // BoardData itself). Set alongside the BoardData it describes; a board load with no draft
        // resets it to BoardDraftSummary.Empty, the same "reset before you know the answer" shape
        // ApplyDraft itself uses for a null/empty draft.
        public static BoardDraftSummary LastLoadedDraftSummary { get; private set; } = BoardDraftSummary.Empty;

        // The base revision the most recent LoadBoardDataAsync call's draft recorded, and whether
        // that draft registers a board of its own - the two things the drift check needs about a
        // draft beyond its rows (NewContributeStrategy.md Phase 2, session 2d).
        //
        // Siblings of LastLoadedDraftSummary and set from the SAME draft object for the same
        // reason its own comment gives: re-resolving them would let the two describe different
        // drafts if draft.json changed on disk in between.
        //
        // UNLIKE LastLoadedDraftSummary, these are NOT suppressed by ViewOfficialPublishedOnly.
        // That toggle hides the contributor's edits from the rendered board; it does not make drift
        // untrue, and the check must still know the answer. Whether a WARNING is shown while the
        // toggle is on is the UI's decision, taken at the banner rather than by blinding the data.
        public static string LastLoadedDraftBaseRevision { get; private set; } = string.Empty;

        public static bool LastLoadedDraftIsNewBoard { get; private set; }

        // Raised with a general status message (e.g. "Checking files...", "Sync complete")
        public static event Action<string>? StatusChanged;

        // Raised with the relative file path of whichever file is currently being processed
        public static event Action<string>? FileDownloadChanged;

        // ###########################################################################################
        // Result details returned from an immediate manual data sync.
        // ###########################################################################################
        public sealed class ManualDataSyncResult
        {
            public int ChangedCount { get; init; }
            public bool MainExcelChanged { get; init; }
            public bool AnyExcelChanged { get; init; }
        }

        private sealed class MainExcelResolution
        {
            public string FileName { get; init; } = string.Empty;
            public bool RequiresAppUpdate { get; init; }
            public string SourceDescription { get; init; } = string.Empty;
        }

        // ###########################################################################################
        // Resolves the best main file from a file list, preferring the newest compatible
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
        // Applies the selected main file resolution to runtime state and emits a detailed log line.
        // ###########################################################################################
        private static void ApplyMainExcelResolution(MainExcelResolution resolution, string context)
        {
            ResolvedMainExcelFileName = resolution.FileName;
            DataUpdateRequiresAppUpdate = resolution.RequiresAppUpdate;

            Logger.Info(
                $"Main file resolution [{context}] source=[{resolution.SourceDescription}] file=[{(string.IsNullOrWhiteSpace(resolution.FileName) ? "none" : resolution.FileName)}] requires_app_update=[{resolution.RequiresAppUpdate}]");
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
        // Ensures the selected main file exists locally after sync and falls back to the local
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
                    $"Resolved main file [{preferredResolution.FileName}] was not available after sync - falling back to local cached [{localFallbackResolution.FileName}] in [{context}]");

                onStatus?.Invoke("Latest main data file could not be downloaded - using local cached data");

                return new MainExcelResolution
                {
                    FileName = localFallbackResolution.FileName,
                    RequiresAppUpdate = preferredResolution.RequiresAppUpdate || localFallbackResolution.RequiresAppUpdate,
                    SourceDescription = $"local fallback after failed online sync ({context})"
                };
            }

            Logger.Warning(
                $"Resolved main file [{preferredResolution.FileName}] was not available after sync and no usable local fallback existed in [{context}]");

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
            _legacyDataRoot = DataManager.FindLegacyDataRoot();
            DataManager.UseOwnOscilloscopesPath(Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                AppConfig.AppFolderName,
                OwnOscilloscopeFile.FileName));
            Logger.Info($"Data root is [{_dataRoot}]");

            bool isNewRoot = !Directory.Exists(_dataRoot);
            Directory.CreateDirectory(_dataRoot);

            if (isNewRoot)
            {
                var bundledData = BundledDataLocation.Find(AppContext.BaseDirectory, Directory.Exists);
                if (bundledData != null)
                {
                    Logger.Info("Data root folder created — seeding from install package");
                    RaiseStatus("Seeding data from install package...");
                    await Task.Run(() => CopyDirectory(bundledData, _dataRoot));
                    Logger.Info("Data seeded from install package");
                }
                else
                {
                    Logger.Info($"Data root folder created - all files will be downloaded from the {AppConfig.GetOnlineSourceLabel()}");
                }
            }
            else if (UserSettings.CheckDataOnLaunch)
            {
                Logger.Info($"Checking the {AppConfig.GetOnlineSourceLabel()} for new or updated files");
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
                    // Before any cleanup can run (OscilloscopeCaptureRelocation).
                    DataManager.RelocateOscilloscopeCapturesAtLaunch(_syncManifest);

                    MainExcelResolution onlineResolution = ResolveMainExcelCandidate(
                        _syncManifest.Select(entry => entry.File),
                        allowSyntheticLegacyFallback: false,
                        sourceDescription: "online manifest");

                    Logger.Info(
                        $"Main file candidates at startup: local=[{localResolution.FileName}] online=[{onlineResolution.FileName}]");

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
                            label: "Main data file");
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
                Logger.Warning("Newer main data file versions exist, but an application update is required to utilize them");
                RaiseStatus("Data components outdated - application update required!");
            }

            Logger.Info($"Resolved main data file to be used: [{ResolvedMainExcelFileName}]");

            RaiseStatus("Loading hardware definitions...");
            await Task.Run(LoadMainExcel);

            // No separate guard for a skipped sync: _syncManifest is only ever assigned above, so it
            // is still null whenever the manifest fetch did not run.
            if (_syncManifest != null && HardwareBoards.Count > 0)
            {
                // Draft-only boards are excluded: their ExcelDataFile names a file that exists only
                // as an identity key and is never on disk or in the manifest (see NewBoardIdentity),
                // so asking sync to fetch it would be a guaranteed miss on every launch.
                var boardExcelFiles = HardwareBoards
                    .Where(entry => !entry.IsDraftOnly)
                    .Select(entry => thisNormalizeRelativePath(entry.ExcelDataFile))
                    .Where(path => !string.IsNullOrWhiteSpace(path))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                RaiseStatus("Checking board data files...");
                await OnlineServices.SyncFilesAsync(
                    _syncManifest,
                    _dataRoot,
                    file => boardExcelFiles.Contains(thisNormalizeRelativePath(file)),
                    RaiseStatus,
                    RaiseFileDownload,
                    label: "board data files");
            }
        }

        // Returns true when a background sync manifest is queued and ready to process.
        public static bool HasPendingSync => _syncManifest != null;

        // ###########################################################################################
        // Syncs every file that is not a list - neither a main file nor a board file - using the
        // manifest already fetched at startup: in a 4.0 tree that is the stored files and the MiniPro
        // tests. Intended to run silently in the background after the UI has opened.
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
                file => !DataManager.IsListFile(thisNormalizeRelativePath(file)),
                onStatus,
                onFile,
                label: "remaining data files");
        }

        // ###########################################################################################
        // Whether a data-root-relative path is one of the tree's LISTS - a main file
        // ("Classic-Repair-Toolbox.v<version>.json" at the top) or a board file ("<id>/board.json") -
        // which the launch syncs before the window opens, the main file the newest CRT may read
        // and only the board files it lists. Everything else is synced in the background.
        // ###########################################################################################
        internal static bool IsListFile(string relativePath)
        {
            string path = thisNormalizeRelativePath(relativePath);

            if (DataTreeLayout.IsBoardFile(path))
            {
                return true;
            }

            return !path.Contains('/') &&
                   path.StartsWith(AppConfig.MainExcelFileNamePrefix, StringComparison.OrdinalIgnoreCase) &&
                   path.EndsWith(AppConfig.MainExcelFileSuffix, StringComparison.OrdinalIgnoreCase);
        }

        // ###########################################################################################
        // Performs an immediate manual data sync while the application is already running.
        // Startup behavior remains unchanged elsewhere; this manual path keeps a single UI sync flow,
        // refreshes the compatible main file first so the latest board list is known, then syncs
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
                    AnyExcelChanged = false
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
                    AnyExcelChanged = false
                };
            }

            var manifest = _syncManifest;
            _syncManifest = null;

            MainExcelResolution onlineResolution = ResolveMainExcelCandidate(
                manifest.Select(entry => entry.File),
                allowSyntheticLegacyFallback: false,
                sourceDescription: "online manifest");

            Logger.Info(
                $"Main file candidates during manual sync: local=[{localResolution.FileName}] online=[{onlineResolution.FileName}]");

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
                    label: "Main data file");

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

            // Draft-only boards excluded for the same reason as the startup sync above: their
            // ExcelDataFile is an identity key, not a file the server has or will ever have.
            var boardExcelFiles = HardwareBoards
                .Where(entry => !entry.IsDraftOnly)
                .Select(entry => thisNormalizeRelativePath(entry.ExcelDataFile))
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Not "Checking data from the ... source - please wait...": that is the data-check card's
            // heading already, and the card said it twice (DataSyncStatusTip.Checking).
            ReportStatus("Checking board data files...");
            int remainingChangedCount = await OnlineServices.SyncFilesAsync(
                manifest,
                _dataRoot,
                file =>
                {
                    string normalizedFile = thisNormalizeRelativePath(file);

                    return boardExcelFiles.Contains(normalizedFile) || !DataManager.IsListFile(normalizedFile);
                },
                ReportStatus,
                filePath =>
                {
                    if (!string.IsNullOrWhiteSpace(filePath) && DataTreeLayout.IsBoardFile(filePath))
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
                AnyExcelChanged = anyExcelChanged
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
            return Path.Combine(appData, AppConfig.AppFolderName, AppConfig.DataFolderName);
        }

        // ###########################################################################################
        // CRT 3.x's Excel data folder, when it is there - "" otherwise. Only read: the shared files
        // an Excel-era draft cites are found in it while that draft is converted (DraftConversion).
        // ###########################################################################################
        internal static string FindLegacyDataRoot()
        {
            string folder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                AppConfig.AppFolderName,
                AppConfig.LegacyDataFolderName);

            return Directory.Exists(folder) ? folder : string.Empty;
        }

        // ###########################################################################################
        // Points the data layer at an explicit data root and main workbook and loads the hardware,
        // board and oscilloscope definitions from it. This is the local half of InitializeAsync with
        // no online sync and no seeding - the test suite uses it to load a temporary data root
        // instead of the user's real one.
        // ###########################################################################################
        internal static void LoadFrom(string dataRoot, string mainExcelFileName, string ownOscilloscopesPath = "")
        {
            _dataRoot = dataRoot;
            _legacyDataRoot = string.Empty;
            DataManager.UseOwnOscilloscopesPath(ownOscilloscopesPath);
            ResolvedMainExcelFileName = mainExcelFileName;
            LoadMainExcel();
        }

        // ###########################################################################################
        // The own-models file CRT works with from now on. Another file (a test's load) - or none -
        // holds nothing of the last one: the list, its problem and the read are forgotten HERE, where
        // the file changes, never at a reload's start (code review, 2026-10-29: cleared there, a file
        // that had failed to read looked like another file, and a model made while the manual data
        // check ran was written over it).
        // ###########################################################################################
        private static void UseOwnOscilloscopesPath(string path)
        {
            lock (OwnScopeModelsLock)
            {
                if (string.IsNullOrWhiteSpace(path) || !string.Equals(_ownOscilloscopesPath, path, StringComparison.Ordinal))
                {
                    OwnScopeModels = [];
                    OwnScopeModelsProblem = null;
                    _ownScopeModelsReadFrom = null;
                    _ownScopeModelRenames.Clear();
                }

                _ownOscilloscopesPath = path;
            }
        }

        // ###########################################################################################
        // Reads the hardware, board and oscilloscope definitions from the main file, then merges in
        // the boards that exist only as drafts.
        //
        // *** THE MODEL LISTS STAY AS THEY ARE UNTIL THE NEW MAIN FILE IS READ (code review,
        // 2026-10-29). *** Emptied at the start, a manual data check left the model editor with no
        // published model for as long as it ran: a name was checked against the own models alone, and
        // the drop-downs it rebuilt had none of the published ones. Only a main file that cannot be read
        // takes the published models out.
        // ###########################################################################################
        private static void LoadMainExcel()
        {
            HardwareBoards = new List<HardwareBoardEntry>();

            if (string.IsNullOrEmpty(ResolvedMainExcelFileName))
            {
                Logger.Warning("No compatible main data file matched the current application version");
                DataManager.NoPublishedScopeModels();
                return;
            }

            var excelPath = Path.Combine(_dataRoot, ResolvedMainExcelFileName);

            if (!File.Exists(excelPath))
            {
                Logger.Warning($"Main data file not found - [{excelPath}]");
                DataManager.NoPublishedScopeModels();
                return;
            }

            // ###########################################################################################
            // *** THE MAIN FILE IS THE CRT 4.0 TREE'S "Classic-Repair-Toolbox.v4.0.0.json" (2026-10-24). ***
            // The board list in the order of the drop-downs, each board by its id - its board file is
            // "<id>/board.json", which is also the identity CRT carries for it (ExcelDataFile) - and the
            // oscilloscopes with their brand on every entry. Boards made the old way, listed by a
            // "_UserContribution" workbook, are not read here: a 4.0 first launch takes them in as
            // drafts (LocalDataMigration).
            // ###########################################################################################
            MainDocument main;

            try
            {
                main = DataTreeJson.ReadMain(File.ReadAllText(excelPath));
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Failed to load main data file - [{ex.Message}]");
                DataManager.NoPublishedScopeModels();
                return;
            }

            // Read by the server's rule (MainFileListing.Rows - code review, 2026-10-26): a blank
            // hardware name is the row above's, so CRT lists a board where the server's checks see it.
            var entries = MainFileListing.Rows(main)
                .Where(row => row.ExcelDataFile.Length > 0)
                .Select(row => new HardwareBoardEntry
                {
                    HardwareName = row.HardwareName,
                    BoardName = row.BoardName,
                    ExcelDataFile = row.ExcelDataFile,
                    HardwareNotes = row.Notes,
                })
                .ToList();

            // Boards that exist only as a local draft ("Add a new board"). Merged last, so a listed
            // board wins a name collision over a draft. A board folder put into Drafts/ by hand
            // becomes a draft FIRST, so a new board among them is merged in on this same pass
            // (owner request, 2026-09-27).
            DraftManager.ImportHandPlacedFolders(
                entries.Select(entry => new KnownDraftBoard(entry.ExcelDataFile, entry.IsPublished)),
                _legacyDataRoot,
                _dataRoot);

            MergeDraftOnlyBoards(entries);

            HardwareBoards = entries;

            Logger.Info($"Listed [{entries.Count}] boards from main data file:");
            foreach (var boardEntry in entries)
            {
                Logger.Info($"    [{boardEntry.ExcelDataFile}]");
            }

            DataManager.BeforeOwnScopeModelsForTests?.Invoke();
            DataManager.LoadOwnScopeModels(main.Oscilloscopes);

            Logger.Info($"Loaded [{Oscilloscopes.Count}] oscilloscope definitions from main data file - [{OwnScopeModels.Count}] of them your own");
            DataManager.ScopeModelsListed?.Invoke();
        }

        // ###########################################################################################
        // THE SCOPE MODELS LISTED AGAIN - raised by every load of the main file, ON THE THREAD IT RAN
        // ON (a manual data check's is the pool's), once the lists hold the new models. Main follows it
        // on the Oscilloscope tab (FollowScopeModelsReloaded) - at once, not after the whole data check
        // (code review, 2026-10-29: in between, a change typed into a model the reload had renamed was
        // dropped). Subscribed only by Main.StartAsync, so a test's load reaches nobody.
        // ###########################################################################################
        public static event Action? ScopeModelsListed;

        // Run just before the own models are taken up - a test's way to look at the lists while a
        // reload is under way (DataManagerTests). Never set by CRT.
        internal static Action? BeforeOwnScopeModelsForTests { get; set; }

        // What the loads of this run renamed own models to ("<model> (my version)"): the name a model
        // goes by now, by its vendor and the name it had - null when it was not renamed. The
        // Oscilloscope tab follows a renamed model by this, never by guessing from a name's start.
        public static string? OwnScopeModelRenamedTo(string vendor, string seriesOrModel)
        {
            lock (OwnScopeModelsLock)
            {
                return DataManager._ownScopeModelRenames.TryGetValue($"{vendor.Trim()}\n{seriesOrModel.Trim()}", out string? renamed) ? renamed : null;
            }
        }

        // Under OwnScopeModelsLock; capitals aside, as model names are.
        private static readonly Dictionary<string, string> _ownScopeModelRenames = new(StringComparer.OrdinalIgnoreCase);

        // ###########################################################################################
        // THE USER'S OWN MODELS, beside the `published` ones just read - found once in 3.x's main
        // workbook (OwnOscilloscopeFile); a test's load has no file of the user's. Three steps, each
        // simple because of ONE RULE: the editor changes the list only once its file was read
        // (SaveOwnScopeModelsAsync refuses before), and CRT is the file's only writer.
        //
        //   1. THE FILE, READ ONCE (code review, 2026-10-29). A reload (the manual data check, on the
        //      pool) reads it again only while no read went well - so nothing can have been changed
        //      or written since, and what it reads simply becomes the list. Read once, the list held
        //      is the truth: reading it again could put back an older list than one just saved.
        //   2. RECONCILED, IN ONE STEP WITH THE LIST HELD NOW: a model published since under an own
        //      model's name takes it over, or names it apart (OscilloscopeModelRules.Reconciled). An
        //      edit made meanwhile is reconciled with the rest - never undone, never waited for.
        //   3. WRITTEN BACK when that changed it, by the save's own write: the latest list, in turn.
        //
        // *** WHAT COULD NOT BE READ IS NEVER WRITTEN (code review, 2026-10-29). *** A file that could
        // not be read reads as "no models"; OwnScopeModelsProblem then says why, and nothing writes the
        // file this run - neither the reconcile here nor the editor.
        // ###########################################################################################
        private static void LoadOwnScopeModels(IReadOnlyList<MainDocument.OscilloscopeRow> published)
        {
            string path = _ownOscilloscopesPath;
            var notes = new List<string>();

            if (!DataManager.OwnScopeModelsFileRead(path))
            {
                OwnScopeModelsWriteGate.Wait();

                try
                {
                    IReadOnlyList<MainDocument.OscilloscopeRow> read =
                        OwnOscilloscopeFile.LoadOrFind(path, _legacyDataRoot, published, notes, out string? problem);

                    lock (OwnScopeModelsLock)
                    {
                        OwnScopeModels = read.ToList();
                        OwnScopeModelsProblem = problem;
                        _ownScopeModelsReadFrom = problem is null && !string.IsNullOrWhiteSpace(path) ? path : null;
                    }
                }
                finally
                {
                    OwnScopeModelsWriteGate.Release();
                }
            }

            bool changed;

            lock (OwnScopeModelsLock)
            {
                PublishedScopeModels = published.ToList();

                (IReadOnlyList<MainDocument.OscilloscopeRow> reconciled, changed) =
                    OscilloscopeModelRules.Reconciled(OwnScopeModels, published, out IReadOnlyList<(string Vendor, string From, string To)> renamed);
                DataManager.ListScopeModelsCore(reconciled);

                foreach ((string vendor, string from, string to) in renamed)
                {
                    DataManager._ownScopeModelRenames[$"{vendor}\n{from}"] = to;
                }
            }

            if (changed && !string.IsNullOrWhiteSpace(path))
            {
                OwnScopeModelsWriteGate.Wait();

                try
                {
                    if (DataManager.WriteHeldOwnScopeModels(path) is string failure)
                    {
                        notes.Add($"Your own oscilloscope models could not be written to [{path}] ({failure}).");
                    }
                }
                finally
                {
                    OwnScopeModelsWriteGate.Release();
                }
            }

            foreach (string note in notes)
            {
                Logger.Warning(note);
            }
        }

        // Whether the own-models file at `path` was read, and well - the one thing that lets the list
        // change and the file be written.
        private static bool OwnScopeModelsFileRead(string path)
        {
            lock (OwnScopeModelsLock)
            {
                return !string.IsNullOrWhiteSpace(path) && string.Equals(_ownScopeModelsReadFrom, path, StringComparison.Ordinal);
            }
        }

        // ###########################################################################################
        // The list held NOW written to `path` - null when written, else why not. Under the write gate,
        // so writes follow each other and the file ends with the latest list (a save waiting behind
        // another is never overtaken by an older list). A file not read well is never written.
        // ###########################################################################################
        private static string? WriteHeldOwnScopeModels(string path)
        {
            IReadOnlyList<MainDocument.OscilloscopeRow> toWrite;

            lock (OwnScopeModelsLock)
            {
                if (!string.Equals(_ownScopeModelsReadFrom, path, StringComparison.Ordinal))
                {
                    return OwnScopeModelsProblem ?? DataManager.OwnScopeModelsNotReadYet;
                }

                toWrite = OwnScopeModels;
            }

            try
            {
                OwnOscilloscopeFile.Write(path, toWrite);
                Logger.Info($"Saved [{toWrite.Count}] oscilloscope models of your own to [{path}]");
                return null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Your own oscilloscope models could not be written to [{path}] - [{ex.Message}]");
                return ex.Message;
            }
        }

        internal const string OwnScopeModelsNotReadYet =
            "Your own oscilloscope models have not been read yet, so CRT changes none of them - nothing is lost. Run the data check, or start CRT again.";

        // ###########################################################################################
        // A change of the user's own models that the file did not take, taken back out of the lists -
        // `held` is what the file still holds. Nothing is written. The model editor's delete uses it:
        // a model the file keeps must not leave the lists (it came back at the next start).
        // ###########################################################################################
        public static void PutBackOwnScopeModels(IReadOnlyList<MainDocument.OscilloscopeRow> held)
        {
            ArgumentNullException.ThrowIfNull(held);

            lock (OwnScopeModelsLock)
            {
                DataManager.ListScopeModelsCore(held);
            }
        }

        // No main file read: no published model, the own ones as held.
        private static void NoPublishedScopeModels()
        {
            lock (OwnScopeModelsLock)
            {
                PublishedScopeModels = [];
                DataManager.ListScopeModelsCore(OwnScopeModels);
            }

            DataManager.ScopeModelsListed?.Invoke();
        }

        // The published models, then `own` - each list replaced whole. Under OwnScopeModelsLock.
        private static void ListScopeModelsCore(IReadOnlyList<MainDocument.OscilloscopeRow> own)
        {
            OwnScopeModels = own.ToList();
            Oscilloscopes = PublishedScopeModels
                .Select(scope => OscilloscopeEntry.From(scope))
                .Concat(OwnScopeModels.Select(scope => OscilloscopeEntry.From(scope, isOwn: true)))
                .Where(scope => !string.IsNullOrWhiteSpace(scope.SeriesOrModel))
                .ToList();
        }

        // ###########################################################################################
        // THE USER'S OWN MODELS CHANGED by the Oscilloscope tab's model editor: listed at once (on the
        // UI thread, like every reader of the lists), then written to the user's file off it - null
        // when written, else what kept it from being written. Writes follow each other, so the file
        // ends with the latest list. A test's load has no file: the list alone changes.
        //
        // While the file could not be read (OwnScopeModelsProblem), or was not read yet, NOTHING
        // changes - the list nor the file - and the reason is answered (code review, 2026-10-29: the
        // editor's first save replaced an unreadable file with the one model just made).
        // ###########################################################################################
        public static async Task<string?> SaveOwnScopeModelsAsync(IReadOnlyList<MainDocument.OscilloscopeRow> own)
        {
            ArgumentNullException.ThrowIfNull(own);

            string path = _ownOscilloscopesPath;

            lock (OwnScopeModelsLock)
            {
                if (OwnScopeModelsProblem is string problem)
                {
                    return problem;
                }

                if (!string.IsNullOrWhiteSpace(path) && !string.Equals(_ownScopeModelsReadFrom, path, StringComparison.Ordinal))
                {
                    return DataManager.OwnScopeModelsNotReadYet;
                }

                DataManager.ListScopeModelsCore(own);
            }

            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            await OwnScopeModelsWriteGate.WaitAsync().ConfigureAwait(false);

            try
            {
                return await Task.Run(() => DataManager.WriteHeldOwnScopeModels(path)).ConfigureAwait(false);
            }
            finally
            {
                OwnScopeModelsWriteGate.Release();
            }
        }

        // ###########################################################################################
        // Appends every draft-only board to an entry list, skipping any whose hardware/board name
        // already names a board from the main file - by the "{HardwareName}|{BoardName}" identity,
        // skipped with a warning.
        //
        // A collision here is not a corrupt state - it happens naturally when a contributor drafts a
        // whole new board and it is later published officially, at which point the synced entry
        // takes over and the draft's rows apply to it as an ordinary overlay. The warning is how
        // that transition becomes visible in the log rather than a silent change of behaviour.
        //
        // *** THE SAME FOLDER IS A COLLISION TOO, whatever the names (2026-09-27). *** A board made
        // the old way is listed by a "_UserContribution" workbook under its display names
        // ("Commodore 128"), and a copy of it put into Drafts/ is registered under its FOLDER names
        // ("C128") - which a submission needs (see DraftFolderImport). The names differ, so the
        // name check let both through: the board twice in the drop-downs and twice on the Drafts
        // tab, both reading the one draft. The listed entry is kept, so the board stays where the
        // contributor knows it, and it reads the draft exactly as a published board would.
        //
        // *** ONLY WHEN THE LISTED ENTRY REALLY READS THIS DRAFT (code review, 2026-09-27). *** A
        // listed entry finds a draft's workbook by ITS OWN file name. A new board published into
        // BETA is listed as "Data C128 310378 Open128 v2.0.0.xlsx" (the tree's generation) while its
        // draft is still "Data C128 310378 Open128.xlsx" - so the listed entry could not read it, and
        // skipping the draft's own entry left the draft unreachable until production retired it:
        // not openable, not editable, not resubmittable, even after a push-back asked for a
        // correction. So the folder rule applies only when a listed entry carries the draft's own
        // workbook key; otherwise the draft keeps its entry, as it did before the rule.
        // ###########################################################################################
        private static void MergeDraftOnlyBoards(List<HardwareBoardEntry> entries)
        {
            var draftOnlyBoards = DraftManager.EnumerateDraftOnlyBoards();
            if (draftOnlyBoards.Count == 0)
            {
                return;
            }

            var existingKeys = new HashSet<string>(
                entries.Select(entry => $"{entry.HardwareName}|{entry.BoardName}"),
                StringComparer.OrdinalIgnoreCase);

            // *** BY FOLDER IN CRT 4.0 (2026-10-24). *** Every identity of a board's folder reads the
            // same board file there ("<folder>/board.json"), so a listed board in the draft's folder
            // reads the draft - whatever file name the draft's own identity carries (a 3.x "Add a
            // new board" named a workbook). A second entry for it would be a second row of one draft.
            var existingFolders = new HashSet<string>(
                entries.Select(entry => BoardDescriptorRules.BoardIdFromExcelDataFile(entry.ExcelDataFile)).Where(id => id.Length > 0),
                StringComparer.OrdinalIgnoreCase);

            int addedCount = 0;

            foreach (var draftEntry in draftOnlyBoards)
            {
                string folder = BoardDescriptorRules.BoardIdFromExcelDataFile(draftEntry.ExcelDataFile);

                if (existingFolders.Contains(folder))
                {
                    Logger.Info($"Draft-only board [{draftEntry.HardwareName}] / [{draftEntry.BoardName}] is already listed in [{folder}] - the listed entry reads the draft");
                }
                else if (existingKeys.Add($"{draftEntry.HardwareName}|{draftEntry.BoardName}"))
                {
                    entries.Add(draftEntry);
                    existingFolders.Add(folder);
                    addedCount++;
                }
                else
                {
                    Logger.Warning($"Draft-only board skipped - a board with this hardware/board already exists: [{draftEntry.HardwareName}] / [{draftEntry.BoardName}]");
                }
            }

            if (addedCount > 0)
            {
                Logger.Info($"Loaded [{addedCount}] hardware/board entries from local drafts");
            }
        }

        // ###########################################################################################
        // The workbook key of every board the MAIN workbook lists - the boards whose file in Data/
        // really is the published one. What draft retirement may compare a draft against: a
        // draft-only board has no file there, and a _UserContribution board's file is the
        // contributor's own (see HardwareBoardEntry.IsPublished).
        // ###########################################################################################
        public static List<string> PublishedExcelDataFiles() =>
            HardwareBoards
                .Where(entry => entry.IsPublished)
                .Select(entry => entry.ExcelDataFile)
                .Where(file => !string.IsNullOrWhiteSpace(file))
                .ToList();

        // ###########################################################################################
        // Rebuilds the draft-only half of HardwareBoards from disk, leaving everything that came
        // from the main file exactly as it is.
        //
        // This exists so creating or discarding a board takes effect immediately rather than at the
        // next restart. It deliberately does NOT call LoadMainExcel: that re-reads and re-parses the
        // whole main file, and would also rebuild Oscilloscopes as a side effect of what is meant to
        // be a narrow refresh.
        //
        // Call this after creating a new board AND after discarding one - a discarded board that
        // stayed in HardwareBoards would still be listed in the drop-downs, pointing at a draft
        // folder that no longer exists.
        // ###########################################################################################
        public static void RefreshDraftOnlyBoards()
        {
            var entries = HardwareBoards.Where(entry => !entry.IsDraftOnly).ToList();
            MergeDraftOnlyBoards(entries);
            HardwareBoards = entries;
        }

        // ###########################################################################################
        // Lazily loads and caches all sheets from the board-specific Excel file linked to the entry.
        // Delegates to BoardDataReader for parsing and caching. Returns null on failure.
        //
        // Also resolves this board's local draft, if any (NewContributeStrategy.md Phase 2), and
        // hands it to BoardDataReader so the returned BoardData already has the draft overlaid -
        // every caller of this method sees drafted edits without having to know drafts exist.
        // DraftManager.LoadDraftFor re-reads draft.json on every call rather than caching it, so an
        // edit just saved into a draft is picked up by the very next load with no cache to
        // invalidate - draft.json is small (kilobytes) and this is not a hot path.
        //
        // LastLoadedDraftSummary is set from the SAME draft object used for the overlay - never
        // re-resolved - so the two can never describe different drafts even if draft.json changes
        // on disk between the two reads (session 2b: the UI reads this right after awaiting this
        // method, so a race window is there in principle, but re-resolving would only replace one
        // small window with another; a settled call site never observes this).
        //
        // UserSettings.ViewOfficialPublishedOnly (session 2b, task 6 - "view the board as
        // officially published") suppresses the overlay here, at the ONE place it is applied,
        // rather than at each caller: the draft is still RESOLVED and LastLoadedDraftSummary still
        // reports it (so a "you are viewing the published version, N rows hidden" banner could be
        // added later without re-plumbing this method), but BoardDataReader.LoadAsync is handed
        // `null` for its draft argument, so the returned BoardData is the official cache entry
        // itself with nothing merged on top - never a live view of a stale draft under a
        // misleading flag.
        // ###########################################################################################
        public static async Task<BoardData?> LoadBoardDataAsync(HardwareBoardEntry entry)
        {
            // ###########################################################################################
            // *** A DRAFT IS NOW A BOARD FOLDER, SO THIS CHOOSES A FILE INSTEAD OF MERGING ONE ***
            // (NewContributeStrategy.md Phase 6, owner request 2026-09-23).
            //
            // It used to read the published workbook and merge a draft's recorded row deltas over
            // it, which meant a drafted board was a computation rather than a file - there was
            // nothing a contributor could open in Excel. Now the draft folder holds a real board
            // workbook and THAT is what gets read.
            //
            // The rule itself lives in DraftBoardSource (pure, unit tested) rather than here, and
            // "view boards as officially published" is handed to it rather than applied twice.
            // ###########################################################################################
            BoardSourceSelection source = DraftBoardSource.Resolve(
                _dataRoot,
                DraftManager.DraftsRoot,
                entry.ExcelDataFile,
                UserSettings.ViewOfficialPublishedOnly,
                entry.IsPublished);

            string excelPath = source.WorkbookPath.Length > 0
                ? source.WorkbookPath
                : DraftBoardSource.PublishedPathOf(_dataRoot, entry.ExcelDataFile);

            LastLoadedDraftBaseRevision = source.Marker?.BaseRevision ?? string.Empty;
            LastLoadedDraftIsNewBoard = source.IsNewBoard;

            // A board that exists only as a local draft (session 2c, task 9) has no official .xlsx
            // by construction. The marker is the DRAFT's own registration, never merely "the file is
            // missing" - for a board the main workbook DOES list, a missing file is a real sync
            // failure and must keep logging and returning null rather than quietly rendering empty.
            //
            // Read off the marker regardless of "view boards as officially published", so that with
            // the toggle on a draft-only board correctly renders as a blank board (officially, it
            // does not exist yet) rather than failing to open at all.
            bool allowMissingOfficialFile = source.IsNewBoard;

            // ###########################################################################################
            // No draft is passed: there is no overlay any more. The workbook chosen above IS the
            // board, drafted or not.
            //
            // *** THE CACHE KEY MUST NAME THE FILE THAT WAS ACTUALLY READ, NOT THE BOARD. ***
            //
            // BoardDataReader caches by the key it is given, and this used to pass the board's
            // ExcelDataFile - which was safe while a draft was an OVERLAY applied after the cache
            // lookup, because only one file was ever read for a board.
            //
            // It is not safe now. A drafted board has TWO workbooks - the published one and the
            // draft's - and "view boards as officially published" switches between them. With one
            // key for both, the first load caches whichever file it read and the second load gets
            // it back regardless of the toggle: turning the toggle off again kept showing the
            // published board. Caught by
            // DataManagerDraftOverlayTests.Turning_ViewOfficialPublishedOnly_off_again_reads_the_draft_once_more,
            // which fails against the single-key version.
            //
            // Keying by the resolved PATH is correct by construction: two different files can never
            // share a cache entry, and the same file re-read is still a hit.
            // ###########################################################################################
            BoardData? board = await DataManager.ReadBoardWorkbookAsync(excelPath, allowMissingOfficialFile);

            // ###########################################################################################
            // WHICH ROWS THE UI MARKS AS DRAFTED, now derived rather than recorded.
            //
            // BoardDraftSummary used to be built from the draft's own delta list. With the draft
            // stored as a workbook there is no delta list - and there must not be one, because the
            // contributor can edit that workbook in Excel with this application closed. So the
            // drafted rows are worked out by COMPARING the loaded draft against the published copy.
            //
            // Only for a board actually being read from a draft: a published load has nothing to
            // compare against and resets the summary, the same "reset before you know the answer"
            // shape this property always had.
            //
            // The published board is read through the ordinary cached reader, so a second board
            // load pays nothing for it.
            // ###########################################################################################
            LastLoadedDraftSummary = source.IsDraft && board != null
                ? BoardDraftSummary.FromChanges(
                    BoardDataDiffer.Compare(
                        await DataManager.LoadPublishedForComparisonAsync(source.PublishedWorkbookPath),
                        board))
                : BoardDraftSummary.Empty;

            return board;
        }

        // ###########################################################################################
        // Drops every cached parse of one board - BOTH its published workbook and its draft's.
        //
        // *** A BOARD NOW OCCUPIES MORE THAN ONE CACHE KEY. *** Since Phase 6 the cache is keyed
        // by the PATH that was read (see LoadBoardDataAsync for why it had to stop being the
        // board's ExcelDataFile), and a drafted board has two of those. Every caller that used to
        // pass ExcelDataFile to BoardDataReader.ClearCache was therefore clearing a key nothing
        // uses any more - silently, since clearing an absent key is not an error, and the symptom
        // would be a stale board after saving an edit.
        //
        // Clearing both unconditionally is deliberate: whether a draft exists can change between
        // the load and the clear (the contributor may have just created or discarded one), and
        // removing a key that was never there costs nothing.
        // ###########################################################################################
        public static void ClearBoardCache(string excelDataFile)
        {
            if (string.IsNullOrWhiteSpace(excelDataFile))
            {
                return;
            }

            string published = DraftBoardSource.PublishedPathOf(_dataRoot, excelDataFile);
            if (published.Length > 0)
            {
                BoardDataReader.ClearCache(published);
            }

            string drafted = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, excelDataFile);
            if (drafted.Length > 0)
            {
                BoardDataReader.ClearCache(drafted);
            }

            // The pre-Phase-6 key. Harmless if absent, and it means a build that still passes the
            // board identity somewhere unnoticed does not leave a stale entry behind forever.
            BoardDataReader.ClearCache(excelDataFile);
        }

        // ###########################################################################################
        // Loads the PUBLISHED copy of a board, purely to compare a draft against it.
        //
        // *** THE SAME CACHE KEY AS ANY OTHER LOAD OF THAT FILE - its own path. *** This used to use
        // a separate "published:" key, from when the drafted load was cached under the board's
        // ExcelDataFile and a shared key would have made the comparison diff a board against
        // itself. Since the cache is keyed by the PATH that was read (see LoadBoardDataAsync), the
        // draft and its published copy already have different keys - so the prefix only made the
        // published workbook get parsed and held TWICE whenever "view boards as officially
        // published" was toggled on a drafted board. The comparison only reads the board, so
        // sharing the entry is safe.
        //
        // Returns null when there is no published copy, which is the ordinary case for a
        // draft-only board. BoardDataDiffer treats a null published board as "every drafted row is
        // an addition", which is the literal truth.
        // ###########################################################################################
        private static async Task<BoardData?> LoadPublishedForComparisonAsync(string publishedWorkbookPath)
        {
            if (string.IsNullOrWhiteSpace(publishedWorkbookPath) || !File.Exists(publishedWorkbookPath))
            {
                return null;
            }

            return await DataManager.ReadBoardWorkbookAsync(publishedWorkbookPath, allowMissingOfficialFile: false);
        }

        // Board workbook reads that overlap share one parse (SharedLoads, code review 2026-10-13) -
        // keyed by the path read (the key BoardDataReader caches it under), whether a missing file
        // is allowed, and the files' STAMP as BoardDataReader takes it: a request made after the
        // workbook or its sidecar was written starts its own read, never joins one of the old file.
        private static readonly SharedLoads<BoardData?> BoardWorkbookReads = new(StringComparer.OrdinalIgnoreCase);

        private static Task<BoardData?> ReadBoardWorkbookAsync(string excelPath, bool allowMissingOfficialFile) =>
            DataManager.BoardWorkbookReads.RunAsync(
                $"{excelPath}\u001F{allowMissingOfficialFile}\u001F{BoardDataReader.StampOf(excelPath)}",
                () => BoardDataReader.LoadAsync(excelPath, excelPath, allowMissingOfficialFile));

        // ###########################################################################################
        // Fires the StatusChanged event with the given message.
        // ###########################################################################################
        private static void RaiseStatus(string message) => StatusChanged?.Invoke(message);

        // ###########################################################################################
        // Fires the FileDownloadChanged event with the given file path.
        // ###########################################################################################
        private static void RaiseFileDownload(string filePath) => FileDownloadChanged?.Invoke(filePath);

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
        // Deletes orphan and non-used files from the data root after building a complete referenced-file
        // map from the online manifest, the main file, the board files it lists and the stored files
        // they name. Deletions are strictly contained inside the data root and only run while
        // launch-time data sync is enabled.
        // Empty directories left behind are removed afterwards, still strictly within the data root.
        // FAIL CLOSED: this is the only code path that deletes user files, so it refuses to run when
        // the referenced-file map could be incomplete - no manifest snapshot (offline launch or a
        // failed fetch), a missing resolved main file, or any list that exists but cannot be read.
        // Every gap in the map would otherwise become a deletion.
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

            return await DeleteOrphanAndUnusedFilesAsync(
                _lastFetchedManifest,
                OscilloscopeCaptureRelocation.Plan.ForCurrentSettings(_dataRoot));
        }

        // ###########################################################################################
        // The cleanup itself, with the manifest snapshot passed in explicitly. The public overload
        // above applies the user-setting gates and hands over the last fetched manifest; the test
        // suite calls this one directly so it can control the snapshot without touching UserSettings.
        //
        // *** NEVER AN OSCILLOSCOPE CAPTURE (code review, 2026-10-13). *** `captures` names the
        // capture folder - nothing inside it is deleted, and a capture folder that holds the whole
        // data folder means nothing is deleted at all - and first moves the captures the old default
        // left in the data root's top level out of it (OscilloscopeCaptureRelocation). What that
        // move was not sure of stays where it is and is kept away from the cleanup too.
        // ###########################################################################################
        internal static async Task<int> DeleteOrphanAndUnusedFilesAsync(
            List<DataFileEntry>? manifestSnapshot,
            OscilloscopeCaptureRelocation.Plan? captures = null)
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

                    // Captures first: kept out of reach, or the cleanup does not run at all.
                    string? captureFolderInsideDataRoot = null;
                    var keptAwayFiles = new HashSet<string>(
                        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                            ? StringComparer.OrdinalIgnoreCase
                            : StringComparer.Ordinal);

                    if (captures != null && !string.IsNullOrWhiteSpace(captures.CaptureFolder))
                    {
                        if (OscilloscopeCaptureRelocation.IsSameOrInside(dataRootFullPath, captures.CaptureFolder))
                        {
                            Logger.Warning($"Orphan/non-used file cleanup skipped - the oscilloscope image folder [{captures.CaptureFolder}] holds the data folder, and nothing in it is deleted");
                            return 0;
                        }

                        if (OscilloscopeCaptureRelocation.IsSameOrInside(captures.CaptureFolder, dataRootFullPath))
                        {
                            captureFolderInsideDataRoot = Path.GetFullPath(captures.CaptureFolder);
                        }

                        if (captures.MoveTarget != null)
                        {
                            var outcome = DataManager.RelocateOscilloscopeCaptures(dataRootFullPath, captures.MoveTarget, manifestSnapshot);
                            keptAwayFiles.UnionWith(outcome.LeftInPlace.Select(Path.GetFullPath));
                        }
                    }

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

                        if (keptAwayFiles.Contains(fullFilePath) ||
                            (captureFolderInsideDataRoot != null &&
                             OscilloscopeCaptureRelocation.IsSameOrInside(fullFilePath, captureFolderInsideDataRoot)))
                        {
                            Logger.Info($"Kept during orphan cleanup - an oscilloscope capture, or an image that may be one: [{relativePath}]");
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

                    int deletedDirectoryCount = thisDeleteEmptyDirectoriesInsideDataRoot(dataRootFullPath, captureFolderInsideDataRoot);

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
        // Moves the oscilloscope captures the old default left in the data root's top level to
        // `moveTarget` (OscilloscopeCaptureRelocation). Needs the manifest: a file it lists is data,
        // whatever its name - so with no manifest nothing is moved. Never throws.
        // ###########################################################################################
        private static OscilloscopeCaptureRelocation.Outcome RelocateOscilloscopeCaptures(
            string dataRootFullPath,
            string moveTarget,
            List<DataFileEntry>? manifest)
        {
            if (manifest == null || manifest.Count == 0)
            {
                return new OscilloscopeCaptureRelocation.Outcome([], []);
            }

            var listed = manifest
                .Select(entry => thisNormalizeRelativePath(entry.File))
                .ToHashSet(OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                    ? StringComparer.OrdinalIgnoreCase
                    : StringComparer.Ordinal);

            return OscilloscopeCaptureRelocation.MoveCapturesOutOfDataRoot(
                dataRootFullPath,
                moveTarget,
                name => listed.Contains(thisNormalizeRelativePath(name)));
        }

        // ###########################################################################################
        // At launch, once the manifest is in: the captures the old default saved in the data root
        // move to the capture folder - for everybody, not only those who tick the cleanup - so all
        // of a user's captures are where the Oscilloscope tab's folder button opens.
        // ###########################################################################################
        private static void RelocateOscilloscopeCapturesAtLaunch(List<DataFileEntry>? manifest)
        {
            try
            {
                var plan = OscilloscopeCaptureRelocation.Plan.ForCurrentSettings(_dataRoot);

                if (plan.MoveTarget != null && !string.IsNullOrWhiteSpace(_dataRoot) && Directory.Exists(_dataRoot))
                {
                    var outcome = DataManager.RelocateOscilloscopeCaptures(Path.GetFullPath(_dataRoot), plan.MoveTarget, manifest);

                    if (outcome.Moved.Count > 0)
                    {
                        Logger.Info($"Moved [{outcome.Moved.Count}] oscilloscope captures out of the data folder into [{plan.MoveTarget}]");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Oscilloscope captures in the data folder could not be moved - [{ex.Message}]");
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
                $"Orphan cleanup discovered [{mainExcelRelativePaths.Count}] main data files: [{string.Join(" | ", mainExcelRelativePaths)}]");

            string resolvedMainExcelRelativePath = thisNormalizeRelativePath(ResolvedMainExcelFileName);

            if (!string.IsNullOrWhiteSpace(resolvedMainExcelRelativePath))
            {
                Logger.Info(
                    $"Orphan cleanup will map only resolved main data file: [{resolvedMainExcelRelativePath}]");

                if (!thisTryAddMappedFilesFromMainExcel(mappedFiles, dataRootFullPath, resolvedMainExcelRelativePath))
                {
                    return null;
                }
            }
            else
            {
                Logger.Warning(
                    "Orphan cleanup did not have a resolved main data file name; mapping all discovered main Excel files");

                foreach (string mainExcelRelativePath in mainExcelRelativePaths)
                {
                    if (!thisTryAddMappedFilesFromMainExcel(mappedFiles, dataRootFullPath, mainExcelRelativePath))
                    {
                        return null;
                    }
                }
            }

            Logger.Info(
                $"Mapped [{mappedFiles.Count}] referenced files from manifest and [{mainExcelRelativePaths.Count}] main data files for orphan cleanup");

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
        // Returns all main data files currently present inside the data root.
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
        // Adds the main file, every board file it lists and every stored file those boards name to
        // the orphan-cleanup mapped-file set. Returns false when the map would be incomplete: the
        // main file is missing or unreadable, or a board file exists but cannot be read.
        // ###########################################################################################
        private static bool thisTryAddMappedFilesFromMainExcel(
            HashSet<string> mappedFiles,
            string dataRootFullPath,
            string mainExcelRelativePath)
        {
            thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, mainExcelRelativePath, "main data file");

            string mainExcelFullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                mainExcelRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!File.Exists(mainExcelFullPath))
            {
                Logger.Warning($"Main data file missing during orphan cleanup mapping: [{mainExcelRelativePath}]");
                return false;
            }

            MainDocument main;

            try
            {
                main = DataTreeJson.ReadMain(File.ReadAllText(mainExcelFullPath));
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Main data file could not be read during orphan cleanup mapping: [{mainExcelRelativePath}] - [{ex.Message}]");
                return false;
            }

            foreach (MainDocument.BoardListingRow board in main.Boards.Where(board => !string.IsNullOrWhiteSpace(board.BoardId)))
            {
                if (!thisTryAddMappedFilesFromBoardExcel(mappedFiles, dataRootFullPath, DataTreeLayout.BoardFilePath(board.BoardId)))
                {
                    return false;
                }
            }

            return true;
        }

        // ###########################################################################################
        // Adds one board file and every stored file it names to the orphan-cleanup mapped-file set.
        // Returns false when the board file exists but cannot be read - its reference set is then
        // unknown, which is not the same as empty. A missing board file is fine: it references
        // nothing, and its synced files stay protected through the manifest snapshot.
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
                "board data file");

            string boardExcelFullPath = Path.GetFullPath(Path.Combine(
                dataRootFullPath,
                normalizedBoardExcelRelativePath.Replace('/', Path.DirectorySeparatorChar)));

            if (!File.Exists(boardExcelFullPath))
            {
                return true;
            }

            if (!BoardDataReader.TryCollectReferencedLocalFiles(boardExcelFullPath, out HashSet<string> referencedFiles))
            {
                Logger.Warning($"Board data file could not be read during orphan cleanup mapping: [{normalizedBoardExcelRelativePath}]");
                return false;
            }

            foreach (string referencedFile in referencedFiles)
            {
                thisTryAddMappedRelativePath(mappedFiles, dataRootFullPath, referencedFile, "board data referenced file");
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
        // Returns whether a fully normalized path is located inside the current data root.
        //
        // The comparison is against the root plus a TRAILING SEPARATOR, not against the bare root.
        // A plain StartsWith treats a SIBLING whose name merely begins with the root's name as
        // being inside it - with a root of [...\CRT\Data], the folder [...\CRT\DataBackup] passes -
        // and this predicate guards orphan-file deletion and empty-directory deletion, so that
        // mistake is the difference between cleaning up the data root and walking into a folder
        // next to it. ExternalTargetLauncher.IsContainedWithinRoot draws the same boundary the same
        // way for the same reason; the two are deliberately identical in behaviour.
        // ###########################################################################################
        private static bool thisIsPathWithinDataRoot(string dataRootFullPath, string candidateFullPath)
        {
            StringComparison comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;

            if (string.Equals(candidateFullPath, dataRootFullPath, comparison))
            {
                return false;
            }

            return candidateFullPath.StartsWith(
                thisEnsureTrailingDirectorySeparator(dataRootFullPath), comparison);
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
        private static int thisDeleteEmptyDirectoriesInsideDataRoot(string dataRootFullPath, string? captureFolderInsideDataRoot)
        {
            int deletedDirectoryCount = 0;

            var directories = Directory.EnumerateDirectories(dataRootFullPath, "*", SearchOption.AllDirectories)
                .Select(Path.GetFullPath)
                .Where(path => thisIsPathWithinDataRoot(dataRootFullPath, path))
                .Where(path => captureFolderInsideDataRoot == null ||
                               !OscilloscopeCaptureRelocation.IsSameOrInside(path, captureFolderInsideDataRoot))
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