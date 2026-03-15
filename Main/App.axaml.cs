using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Handlers.DataHandling;
using Handlers.OnlineHandling;
using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Central configuration — all tunable application values are defined here.
    // Referenced by: OnlineServices, DataManager, UpdateService, Main, Logger.
    // ###########################################################################################
    public static class AppConfig
    {
        // ===== Debug ===============================================================================
        // Only referenced inside #if DEBUG blocks — ignored entirely in Release builds.

        // Enables online sync in DEBUG builds (normally skipped for faster development iteration).
        // Used by: DataManager.InitializeAsync
        public static readonly bool DebugSimulateSync = true;

        // Simulates an available app update in DEBUG builds for UI testing.
        // Used by: UpdateService.CheckForUpdateAsync, UpdateService.PendingVersion
        public static readonly bool DebugSimulateUpdate = true;

        // Fake version string shown in the update banner during debug update simulations.
        // Used by: UpdateService.PendingVersion
        public const string DebugSimulatedVersion = "99.0.0";

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
        
        // Name of the JSON file storing user preferences. Stored alongside the log file.
        // Used by: UserSettings.Load
        public const string SettingsFileName = "Classic-Repair-Toolbox.settings.json";

        // Name of the JSON file storing custom drawn polyline traces. Stored alongside the log file.
        // Used by: TraceStorage.LoadFromFile
        public const string TracesFileName = "Classic-Repair-Toolbox.traces.json";

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

        // URL for the phone-home version check endpoint.
        // Used by: OnlineServices.CheckInVersionAsync
        public const string CheckVersionUrl = "https://classic-repair-toolbox.dk/app-checkin/";

        // Timeout for lightweight API calls (manifest fetch, version check).
        // Used by: OnlineServices.FetchManifestAsync, OnlineServices.CheckInVersionAsync
        public static readonly TimeSpan ApiTimeout = TimeSpan.FromSeconds(5);

        // Timeout per individual file download — files can be large on slow connections.
        // Used by: OnlineServices.SyncFilesAsync
        public static readonly TimeSpan DownloadTimeout = TimeSpan.FromSeconds(30);

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

        // Minimum allowed zoom level (1.0 = 100%).
        // Used by: Main.OnSchematicsZoom
        public const double SchematicsMinZoom = 0.9;

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
        // Builds a display-safe semantic version string.
        // ###########################################################################################
        public static readonly string AppVersionString = GetAppVersion();

        private static string GetAppVersion()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;

            if (version == null)
            {
                return "0.0.0";
            }

            // Include the 4th digit (Revision) if it's explicitly set above 0,
            // otherwise falling back to standard Major.Minor.Build format.
            return version.Revision > 0
                ? $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}"
                : $"{version.Major}.{version.Minor}.{version.Build}";
        }

    }

    public partial class App : Application
    {
        public override void Initialize()
        {
            AvaloniaXamlLoader.Load(this);
        }

        // ###########################################################################################
        // Registers global exception handlers to capture unexpected crashes into the log.
        // ###########################################################################################
        private void SetupGlobalExceptionLogging()
        {
            AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            {
                if (e.ExceptionObject is Exception ex)
                {
                    Logger.Critical($"Unhandled AppDomain Exception: {ex}");
                }
            };

            TaskScheduler.UnobservedTaskException += (s, e) =>
            {
                Logger.Critical($"Unobserved Task Exception: {e.Exception}");
                e.SetObserved(); // Prevents the application from crashing
            };
        }

        // ###########################################################################################
        // Shows the splash screen, initializes data (syncing with online source), then opens the main window.
        // ###########################################################################################
        public override async void OnFrameworkInitializationCompleted()
        {
            Logger.Initialize();
            this.SetupGlobalExceptionLogging();

            Logger.Info($"Classic Repair Toolbox version [{AppConfig.AppVersionString}] launched");

            UserSettings.Load();

            // Apply selected theme early
            if (UserSettings.ThemeVariant == "Dark")
                this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Dark;
            else if (UserSettings.ThemeVariant == "Light")
                this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Light;
            else
                this.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Default;

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

                // Either use local data or sync it from online source
//                await DataManager.InitializeAsync(desktop.Args ?? []);
                await DataManager.InitializeAsync(desktop.Args ?? Array.Empty<string>()); // supporting .NET6

                var main = new Main();
                desktop.MainWindow = main;
                main.Show();
                splash.Close();

                Logger.Info("Application UI opened");

                // UI has finished loading, so we can do a check-in
                _ = OnlineServices.CheckInVersionAsync();
            }

            base.OnFrameworkInitializationCompleted();
        }


    }


}