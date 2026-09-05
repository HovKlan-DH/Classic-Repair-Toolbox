using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.OnlineHandling;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Central configuration — all tunable application values are defined here.
    // Referenced by: OnlineServices, DataManager, UpdateService, Main, Logger.
    // ###########################################################################################
    public static class AppConfig
    {
        // ===== Simulation ==========================================================================
        // The build configuration does NOT change what the application does. Simulations are runtime
        // switches parsed from the command line — see SimulationOptions — so a DEBUG build and a
        // RELEASE build given the same arguments behave identically.

        // Version the simulated update claims to be, unless "--simulate-update=<version>" overrides it.
        // Used by: SimulationOptions.Parse, UpdateService.PendingVersion
        public const string SimulatedUpdateVersion = "99.0.0";

        // ===== App Identity ========================================================================

        // Short application code name used for User-Agent headers and API control payloads.
        // Used by: OnlineServices
        public const string AppShortName = "CRT";

        // Name of the local AppData subfolder used for data and log storage.
        // Used by: DataManager.ResolveDataRoot, Logger.Initialize
        public const string AppFolderName = "Classic-Repair-Toolbox";

        // Name of the log file written inside the AppFolderName directory.
        // Used by: Logger.Initialize
        public const string LogFileName = "Classic-Repair-Toolbox.log";

        // Name of the crash file written inside the AppFolderName directory, alongside the log.
        //
        // Deliberately a SEPARATE file from LogFileName, because the two have opposite lifetimes:
        // the log is truncated on every launch, and a user reporting a crash has almost always
        // relaunched the application before they get around to sending anything - which erases the
        // very thing they were asked for. The crash file is only ever appended to, so it survives
        // any number of restarts and holds every crash the installation has ever had.
        // Used by: CrashLogger.ResolveCrashFilePath
        public const string CrashFileName = "Classic-Repair-Toolbox.crash.log";
        
        // Name of the JSON file storing user preferences. Stored alongside the log file.
        // Used by: UserSettings.Load
        public const string SettingsFileName = "Classic-Repair-Toolbox.settings.json";

        // Name of the JSON file storing custom drawn polyline traces. Stored alongside the log file.
        // Used by: TraceStorage.LoadFromFile
        public const string TracesFileName = "Classic-Repair-Toolbox.traces.json";

        // Name of the local folder holding worklog workbooks (repair jobs) and their entries. Stored
        // alongside the log file, but deliberately its own subfolder: worklog data is purely local
        // and must never be synced like "Data" nor mixed in with settings/log files. Overridden
        // entirely by "--workbooks-root=", the same idea as "--data-root=" for the synced Data folder.
        // Used by: WorklogManager.Load
        public const string WorklogFolderName = "Workbooks";

        // Name of the JSON file holding one workbook's own record, stored inside that workbook's
        // own subfolder of WorklogFolderName (e.g. "Workbook/1/index.json"). There is deliberately
        // no separate file indexing all workbooks - deleting a workbook's subfolder is how a
        // workbook is removed, with nothing else left to keep in sync.
        // Used by: WorklogManager
        public const string WorklogIndexFileName = "index.json";

        // Name of the JSON file holding one workbook's list of worklog entries, stored inside that
        // same workbook subfolder alongside WorklogIndexFileName (e.g. "Workbook/1/entries.json").
        // Used by: WorklogManager
        public const string WorklogEntriesFileName = "entries.json";

        // Prefix and suffix for the versioned main Excel file containing hardware definitions.
        // Used by: DataManager.InitializeAsync, DataManager.LoadMainExcel
        public const string MainExcelFileNamePrefix = "Classic-Repair-Toolbox.v";
        public const string MainExcelFileSuffix = ".xlsx";

        // Name of the main Excel file containing all hardware and board definitions.
        // Used by: DataManager.InitializeAsync, DataManager.LoadMainExcel
        public const string MainExcelFileName = "Classic-Repair-Toolbox.xlsx";

        // ===== Online Services =====================================================================

        // URL to the JSON manifest listing all data files and their SHA-256 checksums.
        // Used by: OnlineServices.FetchManifestAsync
        public const string ChecksumsUrl = "https://classic-repair-toolbox.dk/app-data/dataChecksums.json";
        public const string ChecksumsUrl_test = "https://classic-repair-toolbox.dk/app-data-BETA/dataChecksums.json";

        // URL for the phone-home version check endpoint.
        // Used by: OnlineServices.CheckInVersionAsync
        public const string CheckVersionUrl = "https://classic-repair-toolbox.dk/app-checkin/";

        // URL receiving component contribution uploads (Assets/Webserver/app-contribution/api/index.php).
        // Used by: ComponentContributionWindow.ProcessAndSendContributionAsync
        public const string ContributionUploadUrl = "https://classic-repair-toolbox.dk/app-contribution/api/";

        // Timeout for genuinely small API calls - a short form POST and its short reply.
        // Used by: OnlineServices.CheckInVersionAsync
        public static readonly TimeSpan ApiTimeout = TimeSpan.FromSeconds(5);

        // Timeout for the checksum manifest fetch. This is NOT a lightweight call and must not go
        // back to sharing ApiTimeout: the manifest lists every data file on the server (~11,000
        // entries, ~3.3 MB uncompressed) and is the largest single transfer at startup.
        // HttpClient.Timeout is a total budget covering DNS, connect, TLS and reading the whole
        // body, so the old 5 seconds failed consistently - not intermittently - for every user on
        // less than roughly 5.5 Mbit/s.
        // Used by: OnlineServices.FetchManifestAsync
        public static readonly TimeSpan ManifestTimeout = TimeSpan.FromSeconds(30);

        // Timeout per individual file download — files can be large on slow connections.
        // Used by: OnlineServices.SyncFilesAsync
        public static readonly TimeSpan DownloadTimeout = TimeSpan.FromSeconds(30);

        // Timeout for feedback upload requests, which may include large attachments.
        // Used by: TabFeedback.ProcessAndSendFeedbackAsync
        public static readonly TimeSpan UploadTimeout = TimeSpan.FromMinutes(5);

        // ===== GitHub Updates ======================================================================

        // GitHub repository owner used to check for application updates via Velopack.
        // Used by: UpdateService.CheckForUpdateAsync
        public const string GitHubOwner = "HovKlan-DH";

        // GitHub repository name used to check for application updates via Velopack.
        // Used by: UpdateService.CheckForUpdateAsync
        public const string GitHubRepo = "Classic-Repair-Toolbox";

        // ===== Schematics Viewer ==================================================================

        // Zoom multiplier applied per mouse wheel step.
        // Used by: Main.OnSchematicsZoom
        public const double SchematicsZoomFactor = 1.5;

        // Maximum allowed zoom level.
        // Used by: Main.OnSchematicsZoom
        public const double SchematicsMaxZoom = 20.0;

        // Maximum pixel width used when pre-scaling schematic thumbnail images.
        // Used by: Main.OnBoardSelectionChanged, Main.CreateScaledThumbnail, Main.CreateScaledThumbnailWithHighlights
        public const int ThumbnailMaxWidth = 800;

        // Logical pixel size of the splash screen window, matching Splash.axaml Width/Height.
        // Used by: App.OnFrameworkInitializationCompleted to center the splash on the saved screen.
        public const int SplashWidth = 600;
        public const int SplashHeight = 350;

        // ###########################################################################################
        // Numeric application version used for version comparisons against versioned data files.
        // ###########################################################################################
        public static readonly string AppNumericVersionString = GetNumericAppVersion();

        // ###########################################################################################
        // Human-readable application version shown in UI and logs.
        // Prefers InformationalVersion and strips SemVer build metadata.
        // ###########################################################################################
        public static readonly string AppDisplayVersionString = GetDisplayAppVersion();

        // ###########################################################################################
        // Builds the numeric application version from the executing assembly version metadata.
        // Includes the revision component only when it is explicitly greater than zero.
        // ###########################################################################################
        private static string GetNumericAppVersion()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;

            if (version == null)
            {
                return "0.0.0";
            }

            return version.Revision > 0
                ? $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}"
                : $"{version.Major}.{version.Minor}.{version.Build}";
        }

        // ###########################################################################################
        // Builds the display version from the executing assembly informational metadata.
        // Falls back to the numeric version when no informational version is available.
        // ###########################################################################################
        private static string GetDisplayAppVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();

            var informationalVersion = assembly
                .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                ?.InformationalVersion;

            if (!string.IsNullOrWhiteSpace(informationalVersion))
            {
                var plusIndex = informationalVersion.IndexOf('+');
                return plusIndex >= 0
                    ? informationalVersion.Substring(0, plusIndex)
                    : informationalVersion;
            }

            return GetNumericAppVersion();
        }

        // ###########################################################################################
        // Returns true when the application was compiled as a DEBUG build.
        // RELEASE builds always report false here.
        //
        // This is reported in the log and nothing else. No application behaviour may depend on it —
        // anything that should differ between a normal run and a development run belongs in
        // SimulationOptions as a command-line switch, so both builds can reach it.
        // ###########################################################################################
        public static bool IsDebugBuild
        {
            get
            {
#if DEBUG
                return true;
#else
                return false;
#endif
            }
        }

        // ###########################################################################################
        // Returns the effective checksum manifest URL for the current build and user setting.
        // ###########################################################################################
        public static string GetChecksumsUrl()
        {
            return UserSettings.DownloadDataFromTestSource
                ? ChecksumsUrl_test
                : ChecksumsUrl;
        }

        // ###########################################################################################
        // Returns the "online source" / "online BETA source" phrase used in data-sync status text,
        // reflecting the current UserSettings.DownloadDataFromTestSource setting.
        // ###########################################################################################
        public static string GetOnlineSourceLabel()
        {
            return UserSettings.DownloadDataFromTestSource
                ? "online BETA source"
                : "online source";
        }

    }

    public partial class App : Application
    {
        // ###########################################################################################
        // Loads the Avalonia XAML resources for the application instance.
        // ###########################################################################################
        public override void Initialize()
        {
            AvaloniaXamlLoader.Load(this);
        }

        // ###########################################################################################
        // Applies the persisted theme mode, including JSON-defined user preference colors.
        // ###########################################################################################
        public void ApplyConfiguredTheme()
        {
            if (UserSettings.ThemeVariant == "UserPreference")
            {
                UserSettings.ReloadUserThemeColors();
            }

            this.ClearUserPreferenceThemeResources();

            switch (UserSettings.ThemeVariant)
            {
                case "Dark":
                    this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Dark;
                    break;

                case "UserPreference":
                    this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Light;
                    this.ApplyUserPreferenceThemeResources();
                    break;

                default:
                    this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Light;
                    break;
            }
        }

        // ###########################################################################################
        // Removes any previously applied JSON-backed theme overrides from application resources.
        // ###########################################################################################
        private void ClearUserPreferenceThemeResources()
        {
            foreach (var resourceKey in UserSettings.GetUserThemeColors().Keys)
            {
                this.Resources.Remove(resourceKey);
            }
        }

        // ###########################################################################################
        // Applies JSON-backed user theme colors by overriding the existing application resources.
        // ###########################################################################################
        private void ApplyUserPreferenceThemeResources()
        {
            foreach (var entry in UserSettings.GetUserThemeColors())
            {
                if (!Color.TryParse(entry.Value, out var color))
                {
                    Logger.Warning($"Skipped invalid user theme color: [{entry.Key}] [{entry.Value}]");
                    continue;
                }

                if (UserSettings.IsUserThemeColorResourceKey(entry.Key))
                {
                    this.Resources[entry.Key] = color;
                }
                else
                {
                    this.Resources[entry.Key] = new SolidColorBrush(color);
                }
            }

            Logger.Info($"Applied user preference theme colors: [{UserSettings.GetUserThemeColors().Count} entries]");
        }

        // ###########################################################################################
        // Registers global exception handlers so that ANY unexpected crash reaches the log and the
        // crash file, rather than the application simply vanishing from the user's screen.
        //
        // There are four distinct ways this application can die, and they need four handlers - a
        // crash caught by one is invisible to the others. Two of them were previously unhandled,
        // which is why crashes were being seen with nothing written anywhere:
        //
        //   Dispatcher          - an exception thrown inside a UI event handler, a binding, or a
        //                         layout/render pass on the UI thread. This is BY FAR the most
        //                         likely source in this application (every tab's logic runs here)
        //                         and it was NOT handled before, so those crashes were silent.
        //   Task scheduler      - a faulted Task that nobody awaited. Note this only fires when the
        //                         Task is FINALISED, so it is inherently late and may never arrive
        //                         at all; it is a backstop, not a primary net.
        //   AppDomain           - anything else that reaches the top of a thread, including the
        //                         render thread. The process is already dying by this point.
        //   Startup             - see OnFrameworkInitializationCompleted, which is "async void" and
        //                         so has no caller able to observe a throw.
        //
        // WHY THE DISPATCHER HANDLER DOES NOT SWALLOW THE EXCEPTION: setting "Handled = true" would
        // keep the window alive after an arbitrary failure, leaving the user with a UI in an
        // unknown, half-updated state that quietly corrupts their worklog data on the next save.
        // The report is written and the crash is then allowed to proceed exactly as it did before,
        // so this changes what is RECORDED, never what the application does.
        // ###########################################################################################
        private static void SetupGlobalExceptionLogging()
        {
            CrashLogger.Initialize(AppConfig.AppDisplayVersionString);

            // The UI thread. This is the handler that was missing, and it is the one that matters
            // most: an exception in any click handler, template, binding or layout pass lands here.
            //
            // Written as "Dispatcher.UIThread" EXPLICITLY. Both of these events are INSTANCE members
            // of a Dispatcher, and "Application" happens to inherit a "Dispatcher" property - so a
            // bare "Dispatcher.UnhandledException" inside this class silently binds to
            // "this.Dispatcher" rather than to the type, which reads like a static subscription
            // while being anything but. Naming the UI dispatcher outright says which one is meant.
            //
            // The FILTER runs first and while the stack is still intact - before the dispatcher has
            // unwound it - so it is the place a full stack trace is most reliably available.
            // "RequestCatch" is deliberately left untouched: changing it would alter whether the
            // exception is caught, and this handler only observes.
            Dispatcher.UIThread.UnhandledExceptionFilter += (s, e) =>
            {
                CrashLogger.Log("Dispatcher (UI thread)", e.Exception, isFatal: true);
            };

            // Also subscribed, because the filter does not run in every path that ends here - a
            // miss is far more expensive than a duplicate. The duplicate itself is no longer
            // written out in full: CrashLogger recognises the same exception instance arriving from
            // a second handler and records it as one "[Also seen by: ...]" line under the report it
            // already wrote, so one fault reads as one report while every handler keeps its cover.
            Dispatcher.UIThread.UnhandledException += (s, e) =>
            {
                CrashLogger.Log("Dispatcher (UI thread, unhandled)", e.Exception, isFatal: true);

                // Deliberately NOT setting "e.Handled = true" - see the note above.
            };

            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                var exception = e.ExceptionObject as Exception;

                // "IsTerminating" is false only in rare hosted cases; it is reported rather than
                // assumed, because whether the process survived is the first thing to know when
                // reading a report against a user's description of what they saw.
                CrashLogger.Log(
                    $"AppDomain (terminating: {e.IsTerminating})",
                    exception,
                    isFatal: e.IsTerminating);

                if (exception == null)
                {
                    // A non-Exception throw. Previously this branch logged nothing whatsoever.
                    Logger.Critical($"Unhandled non-Exception object thrown: [{e.ExceptionObject}]");
                }
            };

            TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                CrashLogger.Log("Unobserved Task", e.Exception, isFatal: false);

                // Observing it keeps the previous behaviour: an unawaited background failure has
                // never been allowed to take the application down, and that is still right - the
                // difference is that it is now recorded in full rather than as one log line.
                e.SetObserved();
            };
        }

        // ###########################################################################################
        // Shows the splash screen, initializes data (syncing with online source), then opens the main window.
        // ###########################################################################################
        public override async void OnFrameworkInitializationCompleted()
        {
            // Logging and the crash handlers come FIRST, before anything that could throw, so that
            // a failure in startup itself is still recorded.
            Logger.Initialize();
            SetupGlobalExceptionLogging();

            // This method is "async void", which Avalonia requires here but which has a sharp edge:
            // there is no caller able to observe a throw, so an exception after the first "await"
            // is lost entirely - the splash would simply sit on screen, or the window would never
            // appear, with nothing written down. A crash during startup is also the WORST one to
            // lose, since the user cannot reach the Configuration tab to send their log.
            //
            // TaskScheduler.UnobservedTaskException does not cover this: there is no Task for it to
            // finalise. So the whole body is wrapped explicitly.
            try
            {
                await this.StartApplicationAsync();
            }
            catch (Exception ex)
            {
                CrashLogger.Log("Application startup", ex, isFatal: true);

                // Rethrowing here would only reach the same "async void" void, so instead the
                // failure is surfaced the one way that still works this early: shut down with a
                // non-zero exit code, having already written the report to disk.
                try
                {
                    if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime failedLifetime)
                    {
                        failedLifetime.Shutdown(1);
                    }
                }
                catch (Exception shutdownFailure)
                {
                    Logger.Critical($"Shutdown after startup failure also failed: {shutdownFailure}");
                }
            }
        }

        // ###########################################################################################
        // The real startup sequence: splash, data initialisation, then the main window.
        //
        // Split out of OnFrameworkInitializationCompleted purely so that method can wrap it in a
        // try/catch - see the note there about "async void" losing exceptions.
        // ###########################################################################################
        private async Task StartApplicationAsync()
        {

            // QuestPDF (the workbook PDF export) REQUIRES a licence type to be declared before it
            // generates anything, and throws on the first export if it is not - so it is set here,
            // at startup, rather than at the export call site where a missing line would only be
            // discovered by a user trying to export.
            //
            // Community is the correct type for this project: QuestPDF grants it free of charge to
            // individuals and to organisations under $1M USD annual revenue, which Classic Repair
            // Toolbox - a hobbyist tool given away - is. An organisation above that threshold
            // shipping a fork of this app would need its own commercial licence from QuestPDF.
            QuestPDF.Settings.License = QuestPDF.Infrastructure.LicenseType.Community;

            // The build configuration goes on this line because startup timings are meaningless
            // without it: DEBUG builds are JIT-only, and ReadyToRun applies only to a RID-targeted
            // publish - so a timing from the wrong build is worse than none. It says nothing about
            // application behaviour, which is identical in both builds.
            Logger.Info($"Classic Repair Toolbox version [{AppConfig.AppDisplayVersionString}] [{(AppConfig.IsDebugBuild ? "DEBUG" : "RELEASE")} build] launched");


            // Startup timing. The first milestone covers everything that happened before the log
            // file even existed - Velopack, the Avalonia AppBuilder, platform detection and parsing
            // App.axaml - which is exactly the stretch where nothing is on screen yet.
            var processStartTime = StartupTimeline.TryResolveProcessStartTime();
            var startupTimeline = new StartupTimeline(processStartTime ?? DateTime.Now);

            if (processStartTime == null)
            {
                Logger.Warning("Process start time unavailable - startup milestones are measured from the first log line, not from process start");
            }

            Logger.Info(startupTimeline.Record("Runtime and UI framework ready", DateTime.Now));

            UserSettings.Load();

            // Loaded unconditionally here, NOT inside the desktop-lifetime branch below. Every
            // WorklogManager read returns empty until this has run, and CreateWorkbook refuses with
            // "no usable workbook root folder", so a lifetime other than the classic desktop one
            // would leave the whole worklog feature silently inert.
            //
            // It reads "--workbooks-root=" the same way DataManager reads "--data-root=", and takes
            // the arguments from the process rather than from desktop.Args precisely so it does not
            // have to wait for that branch. Element 0 is the executable path, which is why it is
            // skipped - it is not an argument, and passing it in would have the resolver test a
            // file path against the switch prefix.
            WorklogManager.Load(Environment.GetCommandLineArgs().Skip(1).ToArray());

            // Loud on purpose. A simulated update looks exactly like a real one in a screenshot, so
            // the log has to be the place where that is unambiguous when a bug report arrives.
            //
            // This sits after UserSettings.Load() rather than up beside the version line because it
            // reports whether the launch version check is enabled, and before the load that setting
            // would read as its default instead of the user's actual value.
            if (SimulationOptions.Current.IsAnyActive)
            {
                Logger.Warning("Simulation mode active:");

                foreach (var line in SimulationOptions.Current.DescribeForLog(UserSettings.CheckVersionOnLaunch))
                {
                    Logger.Warning($"    {line}");
                }
            }

            // Apply selected theme early
            this.ApplyConfiguredTheme();

            if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                var splash = new Splash();

                // Place the splash on the same screen the main window will open on
                if (UserSettings.HasWindowPlacement)
                {
                    double scaling = UserSettings.WindowScreenScaling;
                    var centerX = UserSettings.WindowScreenX + UserSettings.WindowScreenWidth / 2;
                    var centerY = UserSettings.WindowScreenY + UserSettings.WindowScreenHeight / 2;
                    splash.WindowStartupLocation = WindowStartupLocation.Manual;
                    splash.Position = new PixelPoint(
                        centerX - (int)(AppConfig.SplashWidth * scaling / 2),
                        centerY - (int)(AppConfig.SplashHeight * scaling / 2));
                }

                desktop.MainWindow = splash;

                // Create a TaskCompletionSource to bridge the event into an awaitable task
                var splashOpened = new TaskCompletionSource();
                splash.Opened += (s, e) => splashOpened.TrySetResult();

                splash.Show();

                // Wait until Avalonia explicitly fires the "opened" event, guaranteeing the UI is visibly drawn
                await splashOpened.Task;

                Logger.Info(startupTimeline.Record("Splash visible", DateTime.Now));

                // Either use local data or sync it from online source
                await DataManager.InitializeAsync(desktop.Args ?? []);

                Logger.Info(startupTimeline.Record("Data initialised", DateTime.Now));

                var main = new Main();
                desktop.MainWindow = main;

                // The window's outward-facing startup - the hardware/board data it reads from
                // DataManager, the update check and the background data sync. These used to run from
                // Main's constructor; they live in StartAsync so the window can be constructed
                // without touching the network or DataManager's statics (see StartAsync's comment).
                //
                // BEFORE Show(), because that is where the constructor did this work: the combos and
                // the first board load have to be in place before the window is first painted, or
                // the user sees an empty dropdown and an empty component list until the load
                // finishes, and Main's own Opened handler runs ahead of the load instead of after
                // it. StartAsync's synchronous head (PopulateHardwareDropDown, which cascades into
                // the whole first board load) therefore still completes before Show(), exactly as it
                // did from the constructor.
                //
                // Deliberately not awaited: everything after that head is long-running background
                // work, and the constructor never blocked on it either. Awaiting here would hold the
                // splash open for the whole data sync and change what "Main window shown" means in
                // the timeline logged below. StartAsync catches and logs its own exceptions, since
                // discarding the Task means nothing else observes them.
                _ = main.StartAsync();

                main.Show();
                splash.Close();

                Logger.Info(startupTimeline.Record("Main window shown", DateTime.Now));
                Logger.Info("Application UI opened");

                // UI has finished loading, so we can do a check-in
                _ = OnlineServices.CheckInVersionAsync();
            }

            base.OnFrameworkInitializationCompleted();
        }


    }


}