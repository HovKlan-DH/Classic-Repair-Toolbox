using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.OnlineHandling;
using System;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Application and data updates: the GitHub release-update banner (check, install, dismiss),
    // the online data-sync flow (manual and background) and its own banner, and the release-notes
    // link. See Main.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Checks for an available update and refreshes the application update banner immediately.
        // This does not interfere with the separate main-Excel compatibility warning banner.
        // ###########################################################################################
        internal async Task CheckForAppUpdateNowAsync()
        {
            bool? available = await UpdateService.CheckForUpdateAsync();

            if (available == true)
            {
                this.ShowApplicationUpdateAvailableBanner();
                return;
            }

            this.HideApplicationUpdateBanner();
        }

        // ###########################################################################################
        // Shows the dedicated dismissable warning banner explaining that newer main Excel data
        // exists but requires a newer application version before it can be used.
        // ###########################################################################################
        private void ShowMainExcelRequiresAppUpdateBanner()
        {
            this.MainExcelRequiresAppUpdateBannerText.Text =
                "Newer main Excel data file is available, but requires a newer application version, due to breaking changes - please update the application. No more data updates will be given for this application version, and worst-case is that future data update will break UI or functionality. Consider yourself informed 😁";
            this.MainExcelRequiresAppUpdateBannerDismissButton.IsEnabled = true;
            this.MainExcelRequiresAppUpdateBanner.IsVisible = true;
        }

        // ###########################################################################################
        // Hides the dedicated main-Excel compatibility warning banner and clears its persisted
        // visible state for the current session.
        // ###########################################################################################
        private void HideMainExcelRequiresAppUpdateBanner()
        {
            this.MainExcelRequiresAppUpdateBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Shows the normal application update banner when GitHub reports that a newer installed
        // application package is available for download.
        // ###########################################################################################
        private void ShowApplicationUpdateAvailableBanner()
        {
            // The "(simulated)" suffix is the only thing distinguishing a faked update from a real
            // one on screen, and screenshots are what arrive in bug reports.
            string simulatedSuffix = SimulationOptions.Current.SimulateUpdate ? " (simulated)" : string.Empty;
            this.UpdateBannerText.Text = $"Version [{UpdateService.PendingVersion}] is available{simulatedSuffix}";
            this.UpdateBannerInstallButton.IsVisible = true;
            this.UpdateBannerViewNotesButton.IsVisible = true;
            this.UpdateBannerInstallButton.IsEnabled = true;
            this.UpdateBannerViewNotesButton.IsEnabled = true;
            this.UpdateBannerDismissButton.IsEnabled = true;
            this.UpdateBanner.IsVisible = true;
        }

        // ###########################################################################################
        // Hides the normal application update banner without affecting the separate main-Excel
        // compatibility warning banner.
        // ###########################################################################################
        private void HideApplicationUpdateBanner()
        {
            this.UpdateBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the dedicated main-Excel compatibility warning banner.
        // ###########################################################################################
        private void OnMainExcelRequiresAppUpdateBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.HideMainExcelRequiresAppUpdateBanner();
        }

        // ###########################################################################################
        // Dismisses the normal application update banner without cancelling the update.
        // ###########################################################################################
        private void OnUpdateBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this.HideApplicationUpdateBanner();
        }

        // ###########################################################################################
        // Performs the manual UI data update check as a single visible sync flow.
        // Startup sync behavior remains unchanged and is handled separately during application launch.
        // keepBannerTextStatic: when true, the banner keeps its initial text during the full refresh.
        // ###########################################################################################
        internal async Task CheckForDataUpdatesNowAsync(bool keepBannerTextStatic = false)
        {
            this._currentSyncFileRelativePath = string.Empty;

            this.SetSyncBannerText($"Checking data from {AppConfig.GetOnlineSourceLabel()} - please wait...");
            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBanner.IsVisible = true;

            this.StartDataSyncStatusIconSpin();

            try
            {
                var syncResult = await DataManager.CheckForDataUpdatesNowAsync(
                    status => Dispatcher.UIThread.Post(() =>
                    {
                        if (keepBannerTextStatic)
                        {
                            return;
                        }

                        if (status.Contains("up to date", StringComparison.OrdinalIgnoreCase) ||
                            status.StartsWith("Update complete", StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }

                        // IMPORTANT: status only on line 1 (file path is line 2, managed separately)
                        this.SetSyncBannerText(status);
                    }),
                    filePath => Dispatcher.UIThread.Post(() =>
                    {
                        // IMPORTANT: file path only on line 2
                        this._currentSyncFileRelativePath = filePath ?? string.Empty;

                        if (keepBannerTextStatic)
                        {
                            return;
                        }

                        // Keep whatever line 1 currently says; only refresh line 2.
                        string currentLine1 = this.SyncBannerText.Text?
                            .Split('\n')[0]
                            .Trim() ?? string.Empty;

                        this.SetSyncBannerText(currentLine1);
                    }));

                if (syncResult.ChangedCount < 0)
                {
                    return;
                }

                if (syncResult.MainExcelChanged)
                {
                    this.RefreshHardwareAndBoardSelectionsAfterMainExcelSync();
                }

                // Clear file line once we switch to the final summary.
                this._currentSyncFileRelativePath = string.Empty;

                if (syncResult.ChangedCount > 0)
                {
                    string bannerText = syncResult.ChangedCount == 1
                        ? "[1] file updated - please refresh board"
                        : $"[{syncResult.ChangedCount}] files updated - please refresh board";

                    this.SetSyncBannerText(ComponentListBuilder.BuildSyncBannerText(bannerText, syncResult.ProtectedFilesCount));
                    this.SyncBannerRefreshButton.IsVisible = true;
                    this.SyncBanner.IsVisible = true;
                }
                else if (syncResult.ProtectedFilesCount > 0)
                {
                    this.SetSyncBannerText(ComponentListBuilder.BuildSyncBannerText("All data files are up to date", syncResult.ProtectedFilesCount));
                    this.SyncBannerRefreshButton.IsVisible = false;
                    this.SyncBanner.IsVisible = true;
                }
                else
                {
                    this.SyncBanner.IsVisible = false;
                }

                if (DataManager.DataUpdateRequiresAppUpdate)
                {
                    this.ShowMainExcelRequiresAppUpdateBanner();
                }

                if (UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    _ = DataManager.DeleteOrphanAndUnusedFilesAsync();
                }
            }
            finally
            {
                this._currentSyncFileRelativePath = string.Empty;
                this.StopDataSyncStatusIconSpin();
            }
        }

        // ###########################################################################################
        // Syncs any remaining non-Excel files and returns the number of files that changed.
        // keepBannerTextStatic: when true, intermediate status text is suppressed during the run.
        // ###########################################################################################
        private async Task<int> SyncRemainingFilesAsync(bool keepBannerTextStatic = false)
        {
            if (!DataManager.HasPendingSync)
            {
                return 0;
            }

            return await DataManager.SyncRemainingAsync(
                status => Dispatcher.UIThread.Post(() =>
                {
                    if (keepBannerTextStatic)
                    {
                        return;
                    }

                    if (status.Contains("up to date", StringComparison.OrdinalIgnoreCase) ||
                        status.StartsWith("Sync complete", StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    // status only on line 1
                    this.SetSyncBannerText(status);
                }),
                filePath => Dispatcher.UIThread.Post(() =>
                {
                    // file only on line 2
                    this._currentSyncFileRelativePath = filePath ?? string.Empty;

                    if (keepBannerTextStatic)
                    {
                        return;
                    }

                    string currentLine1 = this.SyncBannerText.Text?
                        .Split('\n')[0]
                        .Trim() ?? string.Empty;

                    this.SetSyncBannerText(currentLine1);
                }));
        }

        // ###########################################################################################
        // Starts the remaining background sync without blocking the caller.
        //
        // Returns a Task rather than being async void so StartAsync can compose it. It still swallows
        // its own exceptions in the catch below, so nothing faults out of here either way - the
        // signature change is about being awaitable, not about changing failure behaviour.
        // ###########################################################################################
        private async Task StartBackgroundSyncAsync(bool keepBannerTextStatic = false)
        {
            try
            {
                if (!UserSettings.CheckDataOnLaunch || !DataManager.HasPendingSync)
                {
                    if (!this._isShowingDataSyncDisabledBanner)
                    {
                        this.SyncBanner.IsVisible = false;
                        this.SyncBannerRefreshButton.IsVisible = false;
                    }

                    return;
                }

                this.SyncBannerText.Text = $"Checking data from {AppConfig.GetOnlineSourceLabel()} - please wait...";
                this.SyncBannerRefreshButton.IsVisible = false;
                this.SyncBanner.IsVisible = true;

                this.StartDataSyncStatusIconSpin();

                int changed = await this.SyncRemainingFilesAsync(keepBannerTextStatic);

                DataManager.LoadProtectedContributionStateForCurrentData();
                int protectedFilesCount = DataManager.ProtectedContributionFileCount;

                if (changed > 0)
                {
                    string bannerText = changed == 1
                        ? "[1] file updated in background - please refresh board"
                        : $"[{changed}] files updated in background - please refresh board";

                    this.SyncBannerText.Text = ComponentListBuilder.BuildSyncBannerText(bannerText, protectedFilesCount);
                    this.SyncBannerRefreshButton.IsVisible = true;
                    this.SyncBanner.IsVisible = true;
                }
                else if (protectedFilesCount > 0)
                {
                    this.SyncBannerText.Text = ComponentListBuilder.BuildSyncBannerText("All data files are up to date", protectedFilesCount);
                    this.SyncBannerRefreshButton.IsVisible = false;
                    this.SyncBanner.IsVisible = true;
                }
                else
                {
                    this.SyncBanner.IsVisible = false;
                }

                if (UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    _ = DataManager.DeleteOrphanAndUnusedFilesAsync();
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Background data sync failed - [{ex.Message}]");
            }
            finally
            {
                this.thisBackgroundStartupSyncCompletionSource.TrySetResult(true);
                this.StopDataSyncStatusIconSpin();
            }
        }

        // ###########################################################################################
        // Manually reloads the current board configuration.
        // ###########################################################################################
        private void OnRefreshBoardClick(object? sender, RoutedEventArgs e)
        {
            this.SyncBanner.IsVisible = false;
            this.SyncBannerRefreshButton.IsVisible = false;
            this.ReloadCurrentBoardFromDisk(string.Empty);
        }

        // ###########################################################################################
        // Dismisses the sync banner.
        // ###########################################################################################
        private void OnSyncBannerDismiss(object? sender, RoutedEventArgs e)
        {
            this._isShowingDataSyncDisabledBanner = false;
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Dismisses the sync banner when clicking anywhere on it.
        // ###########################################################################################
        private void OnSyncBannerPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            this._isShowingDataSyncDisabledBanner = false;
            this.SyncBanner.IsVisible = false;
        }

        // ###########################################################################################
        // Opens the GitHub release notes page for the pending update version.
        // ###########################################################################################
        private void OnViewReleaseNotesClick(object? sender, RoutedEventArgs e)
        {
            string version = UpdateService.PendingVersion ?? string.Empty;
            string url = string.IsNullOrWhiteSpace(version)
                ? $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}/releases"
                : $"https://github.com/{AppConfig.GitHubOwner}/{AppConfig.GitHubRepo}/releases/tag/{version}";
            OpenUrl(url);
        }

        // ###########################################################################################
        // Downloads and installs the pending update, showing progress in the banner text.
        //
        // A SUCCESSFUL install never returns: ApplyUpdatesAndRestart replaces the process. So every
        // path that reaches the code after the await is a FAILURE, and the banner has to be handed
        // back to the user - it was left reading "Downloading update..." with all three buttons
        // disabled, which is a dead end: no way to retry, no way to dismiss it, and no indication
        // that anything went wrong. DownloadAndInstallAsync catches its own exceptions and reports
        // failure by returning false (the detail is in the log), so the false branch is not an
        // unusual case to be ignored - it is the only way a failure can present at all.
        //
        // The try is there for the same reason TabWorkbooks.OpenEntryEditor carries one: this is an
        // "async void" handler, so anything thrown here - including by the button property writes
        // before the await - is rethrown on the sync context with no caller to catch it and reaches
        // App's global handler as a process-fatal crash. Losing the whole application over a failed
        // update check would be a far worse outcome than the update not installing.
        // ###########################################################################################
        private async void OnInstallUpdateClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.UpdateBannerInstallButton.IsEnabled = false;
                this.UpdateBannerViewNotesButton.IsEnabled = false;
                this.UpdateBannerDismissButton.IsEnabled = false;
                this.UpdateBannerText.Text = "Downloading update...";

                bool installed = await UpdateService.DownloadAndInstallAsync(progress =>
                {
                    Dispatcher.UIThread.Post(() => this.UpdateBannerText.Text = $"Downloading update: {progress}%");
                });

                if (!installed)
                {
                    this.RestoreUpdateBannerAfterFailedInstall();
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Installing the update failed unexpectedly: [{ex.Message}]");
                this.RestoreUpdateBannerAfterFailedInstall();
            }
        }

        // ###########################################################################################
        // Puts the update banner back into an actionable state after a failed install, so the user
        // can retry or dismiss it. Says the install failed and points at the log, since the reason
        // itself was written there by whatever gave up (UpdateService logs the exception).
        // ###########################################################################################
        private void RestoreUpdateBannerAfterFailedInstall()
        {
            this.UpdateBannerText.Text = "The update could not be installed - see the log for details.";

            this.UpdateBannerInstallButton.IsEnabled = true;
            this.UpdateBannerViewNotesButton.IsEnabled = true;
            this.UpdateBannerDismissButton.IsEnabled = true;
        }

        // Lets a headless test drive the failed-install recovery without a real update to fail.
        // The click handler itself cannot be exercised: UpdateService.DownloadAndInstallAsync reaches
        // GitHub over the network, which no test may do.
        internal void RestoreUpdateBannerAfterFailedInstallForTests() =>
            this.RestoreUpdateBannerAfterFailedInstall();
    }
}
