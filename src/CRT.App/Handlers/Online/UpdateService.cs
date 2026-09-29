using CRT;
using Handlers.DataHandling;
using System;
using System.Threading;
using System.Threading.Tasks;
using Velopack;
using Velopack.Sources;

namespace Handlers.OnlineHandling
{
    // ###########################################################################################
    // Handles checking for, downloading, and applying application updates via Velopack.
    // Uses GitHub Releases as the update source.
    // ###########################################################################################
    public static class UpdateService
    {
        private static UpdateManager? _manager = null;
        private static UpdateInfo? _pendingUpdate = null;
        private static string? _lastCheckError;

        // ###########################################################################################
        // Returns the error message from the last failed update check, or null if no error occurred.
        // ###########################################################################################
        public static string? LastCheckError => _lastCheckError;

        // ###########################################################################################
        // Checks GitHub Releases for a newer version.
        // Returns true if an update is available, false if up to date, null if the check failed.
        // ###########################################################################################
        public static async Task<bool?> CheckForUpdateAsync()
        {
            _lastCheckError = null;

            if (SimulationOptions.Current.SimulateUpdate)
            {
                Logger.Warning($"Simulated update - reporting version [{SimulationOptions.Current.SimulatedUpdateVersion}] as available");
                return true;
            }

            try
            {
                bool allowAlpha = UserSettings.AllowAlphaVersionNotification;
                bool allowBeta = UserSettings.ShowDevelopmentVersionNotification;

                // GitHub's "prerelease" flag has no concept of ALPHA vs BETA - it either considers
                // every pre-release or none - so the stage is read back off the version string and
                // the feed is filtered BEFORE Velopack ranks it. Filtering afterwards, on the single
                // candidate it returns, meant a newer alpha permanently hid an available beta from a
                // beta-only user; see StageFilteredUpdateSource's header.
                var source = new StageFilteredUpdateSource(
                    new GithubSource(
                        $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}",
                        null,
                        allowAlpha || allowBeta),
                    allowAlpha,
                    allowBeta);

                _manager = new UpdateManager(source);

                _pendingUpdate = await _manager.CheckForUpdatesAsync();

                return _pendingUpdate != null;
            }
            catch (Velopack.Exceptions.NotInstalledException)
            {
                // The normal outcome when running from "dotnet run" or a plain build output rather
                // than a Velopack install - not an error worth alarming anyone about.
                _lastCheckError = "Not running as an installed application";
                Logger.Warning("Update check skipped - not running as a Velopack-installed application");
                return null;
            }
            catch (Exception ex)
            {
                _lastCheckError = ex.Message;
                Logger.Warning($"Update check failed - [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Downloads the pending update, then applies it and restarts the app.
        // onProgress: optional callback receiving download progress (0-100).
        // Returns true if successful, false if the download/install failed.
        // ###########################################################################################
        // `cancellationToken` stops the DOWNLOAD (2026-09-28: the "please wait" overlay gives up
        // after two minutes with no progress) - so a stalled download that recovered later can never
        // restart the application under somebody who has carried on working.
        public static async Task<bool> DownloadAndInstallAsync(Action<int>? onProgress = null, CancellationToken cancellationToken = default)
        {
            if (SimulationOptions.Current.SimulateUpdate)
            {
                Logger.Warning("Simulated update - faking the download");
                for (int i = 0; i <= 100; i += 5)
                {
                    onProgress?.Invoke(i);
                    await Task.Delay(50, cancellationToken);
                }
                Logger.Warning("Simulated update - download complete (the restart is deliberately skipped)");
                return true;
            }

            if (_manager == null || _pendingUpdate == null)
            {
                Logger.Warning("No pending update - call CheckForUpdateAsync first");
                return false;
            }

            try
            {
                await _manager.DownloadUpdatesAsync(_pendingUpdate, onProgress, cancellationToken);
                Logger.Info("Update downloaded - restarting into new version");
                _manager.ApplyUpdatesAndRestart(_pendingUpdate);
                return true; // Execution technically halts on the line above if restart succeeds
            }
            catch (Exception ex)
            {
                // The caller stopped it - the "please wait" overlay's two-minute limit with no
                // progress. An expected, user-facing path ("stopped moving, try again"), not a failed
                // install, so it is not logged as one (code review, 2026-09-29).
                if (UpdateService.WasStoppedByCaller(ex, cancellationToken))
                {
                    Logger.Warning("Update download stopped - it made no progress within the wait limit");
                    return false;
                }

                Logger.Critical($"Update install failed - [{ex.Message}]");
                return false; // Safely return false instead of crashing the app
            }
        }

        // ###########################################################################################
        // Whether a failed download was the CALLER cancelling it, rather than a fault. Only a
        // cancellation the caller's own token asked for: HttpClient reports its own timeout as a
        // TaskCanceledException too, with the caller's token untouched, and that one IS a failure.
        // ###########################################################################################
        internal static bool WasStoppedByCaller(Exception exception, CancellationToken cancellationToken) =>
            exception is OperationCanceledException && cancellationToken.IsCancellationRequested;

        // ###########################################################################################
        // Returns the version string of the available update, or null if none was found.
        // ###########################################################################################
        public static string? PendingVersion =>
            SimulationOptions.Current.SimulateUpdate
                ? SimulationOptions.Current.SimulatedUpdateVersion
                : _pendingUpdate?.TargetFullRelease.Version.ToString();
    }
}