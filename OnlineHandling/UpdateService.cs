using CRT;
using DataHandling;
using System;
using System.Threading.Tasks;
using Velopack;
using Velopack.Sources;

namespace OnlineHandling
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

#if DEBUG
            if (AppConfig.DebugSimulateUpdate)
            {
                return true;
            }

            _lastCheckError = "Update check disabled in debug builds";
            Logger.Debug("Update check skipped - debug build");
            return null;
#else
            try
            {
                _manager = new UpdateManager(new GithubSource(
                    $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}",
                    null,
                    UserSettings.ShowDevelopmentVersionNotification));

                _pendingUpdate = await _manager.CheckForUpdatesAsync();
                return _pendingUpdate != null;
            }
            catch (Velopack.Exceptions.NotInstalledException)
            {
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
#endif
        }

        // ###########################################################################################
        // Downloads the pending update, then applies it and restarts the app.
        // onProgress: optional callback receiving download progress (0-100).
        // Returns true if successful, false if the download/install failed.
        // ###########################################################################################
        public static async Task<bool> DownloadAndInstallAsync(Action<int>? onProgress = null)
        {
#if DEBUG
            if (AppConfig.DebugSimulateUpdate)
            {
                Logger.Info("Debug simulation - faking update download");
                for (int i = 0; i <= 100; i += 5)
                {
                    onProgress?.Invoke(i);
                    await Task.Delay(50);
                }
                Logger.Info("Debug simulation - download complete (restart skipped in debug)");
                return true;
            }
#endif

            if (_manager == null || _pendingUpdate == null)
            {
                Logger.Warning("No pending update - call CheckForUpdateAsync first");
                return false;
            }

            try
            {
                await _manager.DownloadUpdatesAsync(_pendingUpdate, onProgress);
                Logger.Info("Update downloaded - restarting into new version");
                _manager.ApplyUpdatesAndRestart(_pendingUpdate);
                return true; // Execution technically halts on the line above if restart succeeds
            }
            catch (Exception ex)
            {
                Logger.Critical($"Update install failed - [{ex.Message}]");
                return false; // Safely return false instead of crashing the app
            }
        }

        // ###########################################################################################
        // Returns the version string of the available update, or null if none was found.
        // ###########################################################################################
        public static string? PendingVersion =>
#if DEBUG
            AppConfig.DebugSimulateUpdate ? AppConfig.DebugSimulatedVersion :
#endif
            _pendingUpdate?.TargetFullRelease.Version.ToString();
    }
}