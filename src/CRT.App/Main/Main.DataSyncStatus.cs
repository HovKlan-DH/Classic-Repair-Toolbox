using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Data-sync settings reactions (check-on-launch, workbooks scope, worklog currency), the
    // "local edit disables launch sync" banner and its confirmation dialog, the sync status
    // icon (spin, hover, click-to-sync), and the startup background data validation / orphan
    // file cleanup passes. See Main.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // ###########################################################################################
        // Reacts to "Check for new or updated data at application launch" changes.
        // Re-enabling hides the local-edit warning banner, but disabling from Configuration must
        // not show it. That banner is only shown explicitly after label-editor apply.
        // ###########################################################################################
        private void OnCheckDataOnLaunchSettingChanged(bool isEnabled)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.UpdateDataSyncStatusIcon();

                if (isEnabled)
                {
                    this.HideDataSyncDisabledBanner();
                }
            });
        }

        // ###########################################################################################
        // Reacts to the Configuration tab's "Show all workbooks" / "Show this board's workbooks"
        // radio group being flipped, by rebuilding the Workbooks tab through the one funnel every
        // other worklog change already uses.
        //
        // Needed because the scope decides WHICH workbooks RefreshWorklogBar passes the tab, so the
        // tab cannot re-derive it on its own without a refresh. Before this the event was raised and
        // subscribed by nothing: the setting only appeared to work because switching to the
        // Workbooks tab happens to refresh it on attach - incidental, and no help at all to a
        // Workbooks tab that is already on screen when the setting changes.
        //
        // Posted rather than run inline, matching OnCheckDataOnLaunchSettingChanged: the event is
        // raised from a UserSettings setter, so this keeps the rebuild off the setter's own stack.
        // ###########################################################################################
        private void OnWorkbooksScopeSettingChanged()
        {
            Dispatcher.UIThread.Post(() => this.RefreshWorklogBar());
        }

        // ###########################################################################################
        // Reprints every cost on screen when the Configuration tab's currency drop-down changes,
        // through the same funnel and for the same reason as the scope handler above.
        //
        // The Workbooks tab's summary strip, its entry cards and the worklog editor all build their
        // figures once per refresh, so a currency changed while that tab is on screen would leave
        // the old code beside every number until something else happened to rebuild them - and the
        // most likely moment to change this setting is right after noticing the wrong currency on
        // exactly that screen.
        //
        // Posted rather than run inline, so the rebuild stays off the UserSettings setter's stack.
        // ###########################################################################################
        private void OnWorklogCurrencySettingChanged()
        {
            Dispatcher.UIThread.Post(() => this.RefreshWorklogBar());
        }

        // ###########################################################################################
        // Shows a dismissable main-window banner explaining that launch-time data synchronization
        // is disabled and must be re-enabled manually in Configuration when appropriate.
        // ###########################################################################################
        private void ShowDataSyncDisabledBanner()
        {
            this.SyncBannerText.Text =
                "Data synchronization has been disabled because the board Excel data was edited locally. Re-enable it in the \"Configuration\" tab when safe.";
            this.SyncBannerRefreshButton.IsVisible = false;
            this.SyncBanner.IsVisible = true;
            this._isShowingDataSyncDisabledBanner = true;
        }

        // ###########################################################################################
        // Hides the synchronization-disabled banner without affecting other sync banner flows.
        // ###########################################################################################
        private void HideDataSyncDisabledBanner()
        {
            if (!this._isShowingDataSyncDisabledBanner)
            {
                return;
            }

            this.SyncBanner.IsVisible = false;
            this.SyncBannerRefreshButton.IsVisible = false;
            this._isShowingDataSyncDisabledBanner = false;
        }

        // ###########################################################################################
        // Disables launch-time data synchronization after a local board Excel edit, updates the
        // Configuration tab checkbox, and shows a warning dialog to explain the safety change.
        // ###########################################################################################
        internal async Task DisableLaunchDataSyncAfterLocalBoardEditAsync()
        {
            if (!UserSettings.CheckDataOnLaunch)
            {
                this.ShowDataSyncDisabledBanner();
                return;
            }

            this.TabConfiguration.SetCheckDataOnLaunchCheckBoxValue(false);
            UserSettings.CheckDataOnLaunch = false;

            await this.ShowDataSyncDisabledAfterLocalBoardEditDialogAsync();
        }

        // ###########################################################################################
        // Shows a modal warning dialog explaining why launch-time data synchronization was turned
        // off after the component label editor saved changes to the board Excel file.
        // ###########################################################################################
        private async Task ShowDataSyncDisabledAfterLocalBoardEditDialogAsync()
        {
            var closeButton = new Button
            {
                Content = "OK",
                MinWidth = 110,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center
            };

            var dialog = new Window
            {
                Title = "Data synchronization disabled",
                Width = 520,
                MinWidth = 420,
                CanResize = false,
                ShowInTaskbar = false,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            closeButton.Click += (_, _) => dialog.Close();

            dialog.Content = new Border
            {
                Padding = new Thickness(18),
                Child = new StackPanel
                {
                    Spacing = 14,
                    Children =
                    {
                        new TextBlock
                        {
                            Text =
                                "To avoid potential data loss, \"Check for new or updated data at application launch\" has been disabled because the board Excel data was changed by the component label editor.",
                            TextWrapping = TextWrapping.Wrap
                        },
                        new TextBlock
                        {
                            Text =
                                "A banner is now shown in the main window to indicate that synchronization is disabled.",
                            TextWrapping = TextWrapping.Wrap
                        },
                        closeButton
                    }
                }
            };

            await dialog.ShowDialog(this);
        }

        // ###########################################################################################
        // Enables or disables hardware/board navigation while the schematics label editor owns the board state.
        // ###########################################################################################
        internal void SetSchematicsEditorNavigationEnabled(bool isEnabled)
        {
            this.HardwareComboBox.IsEnabled = isEnabled;
            this.BoardComboBox.IsEnabled = isEnabled;
        }

        // ###########################################################################################
        // Shows or hides the main-window banner indicating that the schematics label editor owns
        // the current board state and navigation is temporarily locked.
        // ###########################################################################################
        internal void SetSchematicsLabelEditorModeBannerVisible(bool isVisible)
        {
            this.SchematicsLabelEditorModeBanner.IsVisible = isVisible;
        }

        // ###########################################################################################
        // Updates the global launch-time data-sync status icon, clickability and tooltip in the main window.
        // When launch-time sync is disabled, hovering the icon temporarily shows the manual refresh icon.
        // ###########################################################################################
        private void UpdateDataSyncStatusIcon()
        {
            bool isEnabled = UserSettings.CheckDataOnLaunch;
            bool isCheckingOnline = this._dataSyncStatusIconSpinRequestCount > 0;
            bool allowManualRefreshWhileDisabled = !isEnabled && this._isHoveringDataSyncStatusIcon;
            bool isClickable = !isCheckingOnline;

            this.DataSyncStatusIconTextBlock.IsVisible = !isCheckingOnline;
            this.DataSyncStatusSpinnerCanvas.IsVisible = isCheckingOnline;

            this.DataSyncStatusIconTextBlock.Text = isEnabled || allowManualRefreshWhileDisabled
                ? "\uf021"
                : "\uf05e";

            if (this.TryFindResource(isEnabled || allowManualRefreshWhileDisabled ? "Text_Success_Fg" : "Text_Fail_Fg", out var brushResource) &&
                brushResource is IBrush brush)
            {
                this.DataSyncStatusIconTextBlock.Foreground = brush;
                this.DataSyncStatusSpinnerEllipse.Stroke = brush;
            }
            else
            {
                IBrush fallbackBrush = isEnabled || allowManualRefreshWhileDisabled
                    ? Brushes.ForestGreen
                    : Brushes.IndianRed;

                this.DataSyncStatusIconTextBlock.Foreground = fallbackBrush;
                this.DataSyncStatusSpinnerEllipse.Stroke = fallbackBrush;
            }

            this.DataSyncStatusIconBorder.Cursor = new Cursor(
                isClickable
                    ? StandardCursorType.Hand
                    : StandardCursorType.Arrow);

            ToolTip.SetTip(
                this.DataSyncStatusIconBorder,
                isCheckingOnline
                    ? $"Checking data from {AppConfig.GetOnlineSourceLabel()}..."
                    : isEnabled
                        ? "Data update is enabled. Click to refresh data now"
                        : this._isHoveringDataSyncStatusIcon
                            ? "Data update at launch is disabled. Click to run a manual refresh now"
                            : "Data update at launch is disabled. Hover to show manual refresh");
        }

        // ###########################################################################################
        // Handles clicks on the top-right data-sync status icon.
        // Clicking always allows a manual refresh unless a sync is already in progress.
        // ###########################################################################################
        private async void OnDataSyncStatusIconPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (!e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
            {
                return;
            }

            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                return;
            }

            e.Handled = true;

            // "async void" - see OnBoardSelectionChanged. CheckForDataUpdatesNowAsync guards its own
            // network work, so this covers the rest of the path rather than the sync itself.
            try
            {
                // Allow the banner to update with progress + current file (2-line banner).
                await this.CheckForDataUpdatesNowAsync(keepBannerTextStatic: false);
            }
            catch (Exception ex)
            {
                Logger.Critical($"Checking for data updates failed: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Shows the manual refresh icon when hovering the disabled data-sync status icon.
        // ###########################################################################################
        private void OnDataSyncStatusIconPointerEntered(object? sender, PointerEventArgs e)
        {
            this._isHoveringDataSyncStatusIcon = true;
            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Restores the normal disabled icon when the pointer leaves the data-sync status icon.
        // ###########################################################################################
        private void OnDataSyncStatusIconPointerExited(object? sender, PointerEventArgs e)
        {
            this._isHoveringDataSyncStatusIcon = false;
            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Starts animating the top-right data-sync status spinner while online data checks are running.
        // Nested calls are reference-counted so overlapping sync flows do not stop the spinner early.
        // ###########################################################################################
        private void StartDataSyncStatusIconSpin()
        {
            if (!UserSettings.CheckDataOnLaunch)
            {
                return;
            }

            this._dataSyncStatusIconSpinRequestCount++;

            if (this._dataSyncStatusIconSpinTimer == null)
            {
                this._dataSyncStatusIconSpinTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(25)
                };

                this._dataSyncStatusIconSpinTimer.Tick += (_, _) =>
                {
                    this._dataSyncStatusIconSpinAngle = (this._dataSyncStatusIconSpinAngle + 1.5) % 52.0;
                    this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = -this._dataSyncStatusIconSpinAngle;
                };
            }

            if (!this._dataSyncStatusIconSpinTimer.IsEnabled)
            {
                this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = -this._dataSyncStatusIconSpinAngle;
                this._dataSyncStatusIconSpinTimer.Start();
            }

            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Stops animating the top-right data-sync status spinner after online data checks finish.
        // ###########################################################################################
        private void StopDataSyncStatusIconSpin()
        {
            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                this._dataSyncStatusIconSpinRequestCount--;
            }

            if (this._dataSyncStatusIconSpinRequestCount > 0)
            {
                return;
            }

            this._dataSyncStatusIconSpinAngle = 0.0;
            this._dataSyncStatusIconSpinTimer?.Stop();
            this.DataSyncStatusSpinnerEllipse.StrokeDashOffset = 0.0;

            this.UpdateDataSyncStatusIcon();
        }

        // ###########################################################################################
        // Starts the background data validation task and converts failures into log entries only.
        // ###########################################################################################
        private static Task StartBackgroundDataValidationAsync()
        {
            return Task.Run(async () =>
            {
                try
                {
                    await DataValidator.ValidateAllDataAsync();
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Background data validation failed - [{ex.Message}]");
                }
            });
        }

        // ###########################################################################################
        // Schedules orphan/non-used file cleanup once per application session when both launch-time
        // sync and orphan cleanup are enabled.
        // ###########################################################################################
        internal void ScheduleOrphanAndUnusedFileCleanupIfEnabled()
        {
            if (!UserSettings.CheckDataOnLaunch ||
                !UserSettings.AllowDeletionOfOrphanAndNonUsedFiles ||
                this.thisHasScheduledOrphanAndUnusedFileCleanup)
            {
                return;
            }

            this.thisHasScheduledOrphanAndUnusedFileCleanup = true;
            _ = this.RunOrphanAndUnusedFileCleanupAfterStartupAsync();
        }

        // ###########################################################################################
        // Waits until startup UI work, background sync, and background validation are all settled before
        // running the orphan/non-used file cleanup as a low-priority background task.
        // ###########################################################################################
        private async Task RunOrphanAndUnusedFileCleanupAfterStartupAsync()
        {
            try
            {
                await this.thisWindowOpenedCompletionSource.Task;
                await this.thisBackgroundStartupSyncCompletionSource.Task;

                if (this.thisBackgroundDataValidationTask != null)
                {
                    await this.thisBackgroundDataValidationTask;
                }

                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);

                await Task.Delay(TimeSpan.FromSeconds(15));

                if (!UserSettings.AllowDeletionOfOrphanAndNonUsedFiles)
                {
                    return;
                }

                await DataManager.DeleteOrphanAndUnusedFilesAsync();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Background orphan/non-used file cleanup failed - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Rebuilds the sync banner text as two lines:
        // Line 1: status/progress only
        // Line 2: current relative path being transferred (single-line, optional)
        // ###########################################################################################
        private void SetSyncBannerText(string statusLine)
        {
            string cleanStatus = (statusLine ?? string.Empty)
                .Replace("\r", string.Empty, StringComparison.Ordinal)
                .Replace("\n", " ", StringComparison.Ordinal)
                .Trim();

            string cleanFile = (this._currentSyncFileRelativePath ?? string.Empty)
                .Replace("\r", string.Empty, StringComparison.Ordinal)
                .Replace("\n", string.Empty, StringComparison.Ordinal)
                .Trim();

            this.SyncBannerText.Text = string.IsNullOrWhiteSpace(cleanFile)
                ? cleanStatus
                : $"{cleanStatus}\n{cleanFile}";
        }
    }
}
