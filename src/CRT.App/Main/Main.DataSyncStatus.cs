using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Data-sync settings reactions (check-on-launch, workbooks scope, worklog currency), the
    // "local edit disables launch sync" notice and its confirmation dialog, the sync status
    // icon (spin, hover, click-to-sync, and a running check's progress in its tooltip), and the
    // startup background data validation / orphan file cleanup passes. See Main.axaml.cs for the
    // file map of the whole partial class.
    // ###########################################################################################
    public partial class Main
    {
        // A running check's progress, shown in the sync icon's tooltip (DataSyncStatusTip): the
        // status line, and the file being transferred.
        private string thisDataSyncProgress = string.Empty;
        private string _currentSyncFileRelativePath = string.Empty;

        // What the last finished check found, under the tooltip's idle text - null before one ran.
        private string? thisLastDataSyncCheck;

        // The data-check card shows a running check (rather than what one found), and whether it
        // was closed while it did.
        private bool thisIsDataCheckOnCard;
        private bool thisDataCheckClosedWhileRunning;

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
            NoticeEmphasis.Show(
                this.SyncBannerText,
                "Data synchronization has been disabled because the board Excel data was edited locally. Re-enable it in the \"Configuration\" tab when safe.");
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

            this.RefreshDataSyncStatusTip();
        }

        // ###########################################################################################
        // Rewrites the sync icon's tooltip alone - cheap enough for every file a check transfers. An
        // open tooltip follows the change, so whoever points at the icon watches the check go.
        // ###########################################################################################
        private void RefreshDataSyncStatusTip()
        {
            ToolTip.SetTip(
                this.DataSyncStatusIconBorder,
                DataSyncStatusTip.Build(
                    isChecking: this._dataSyncStatusIconSpinRequestCount > 0,
                    isEnabled: UserSettings.CheckDataOnLaunch,
                    isHovering: this._isHoveringDataSyncStatusIcon,
                    sourceLabel: AppConfig.GetOnlineSourceLabel(),
                    progress: this.thisDataSyncProgress,
                    file: this._currentSyncFileRelativePath,
                    lastCheck: this.thisLastDataSyncCheck));
        }

        // ###########################################################################################
        // THE DATA-CHECK NOTICE (owner report, 2026-10-11: "I am checking for new data updates at
        // app launch, but I do not see any information about that anywhere"). The check has its own
        // card over the tabs (DataCheckBanner): what it is doing while it runs, then what it found
        // (DataCheckNotice) - "all up to date" goes by itself, "please refresh board" stays. The
        // card floats, so coming and going moves nothing. The sync icon's tooltip says the same,
        // for whoever closed the card.
        //
        // Internal for MainWindowTests (and the demo) - the checks that call these need a network.
        // ###########################################################################################

        // A check is starting: no progress from the one before, and its card says it has started.
        internal void BeginDataSyncProgress()
        {
            this.thisDataSyncProgress = string.Empty;
            this._currentSyncFileRelativePath = string.Empty;
            this.thisIsDataCheckOnCard = true;
            this.thisDataCheckClosedWhileRunning = false;

            this.CancelNoticeHide(this.DataCheckBanner);
            this.DataCheckBannerRefreshButton.IsVisible = false;
            this.DataCheckBanner.IsVisible = true;
            this.RefreshDataSyncProgress();
        }

        // A status the sync reported. Its closing summaries are left out - the outcome says that.
        internal void ReportDataSyncStatus(string status)
        {
            if (!DataSyncStatusTip.IsProgress(status))
            {
                return;
            }

            this.thisDataSyncProgress = status;
            this.RefreshDataSyncProgress();
        }

        // The file the sync is transferring now.
        internal void ReportDataSyncFile(string? filePath)
        {
            this._currentSyncFileRelativePath = filePath ?? string.Empty;
            this.RefreshDataSyncProgress();
        }

        // ###########################################################################################
        // A check has finished: the card says what it found and the tooltip keeps it. `atLaunch` is
        // the check CRT runs by itself at launch - see DataCheckNotice for what stays and what goes.
        //
        // Closed while it ran, the card is not brought back to say something that needs nothing
        // done - only for files to refresh, or a failure somebody asked about.
        // ###########################################################################################
        internal void FinishDataCheck(int changedCount, bool atLaunch)
        {
            this.thisLastDataSyncCheck = DataSyncStatusTip.LastCheck(changedCount, DateTime.Now);
            this._currentSyncFileRelativePath = string.Empty;
            this.thisIsDataCheckOnCard = false;
            this.RefreshDataSyncStatusTip();

            var notice = DataCheckNotice.After(changedCount, atLaunch, AppConfig.GetOnlineSourceLabel());

            if (notice.GoesByItself && this.thisDataCheckClosedWhileRunning)
            {
                return;
            }

            this.CancelNoticeHide(this.DataCheckBanner);
            NoticeEmphasis.Show(this.DataCheckBannerText, notice.Text);
            this.DataCheckBannerRefreshButton.IsVisible = notice.OffersRefresh;
            this.DataCheckBanner.IsVisible = true;

            if (notice.GoesByItself)
            {
                this.HideNoticeLater(this.DataCheckBanner);
            }
        }

        // The tooltip, and the card while it shows a running check.
        private void RefreshDataSyncProgress()
        {
            this.RefreshDataSyncStatusTip();

            if (this.thisIsDataCheckOnCard)
            {
                NoticeEmphasis.Show(this.DataCheckBannerText, DataSyncStatusTip.Checking(
                    AppConfig.GetOnlineSourceLabel(),
                    this.thisDataSyncProgress,
                    this._currentSyncFileRelativePath));
            }
        }

        // Its close button, or a click anywhere on it.
        private void OnDataCheckBannerDismiss(object? sender, RoutedEventArgs e) => this.CloseDataCheckNotice();

        private void OnDataCheckBannerPointerPressed(object? sender, PointerPressedEventArgs e) => this.CloseDataCheckNotice();

        private void CloseDataCheckNotice()
        {
            this.thisDataCheckClosedWhileRunning = this.thisIsDataCheckOnCard;
            this.DataCheckBanner.IsVisible = false;
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

            if (this.IsDataSyncStatusIconSpinning)
            {
                return;
            }

            e.Handled = true;

            // "async void" - see OnBoardSelectionChanged. CheckForDataUpdatesNowAsync guards its own
            // network work, so this covers the rest of the path rather than the sync itself.
            try
            {
                // The pointer is on the icon, so its tooltip shows the check's progress as it goes.
                await this.CheckForDataUpdatesNowAsync();
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
        //
        // It spins with "Check for new or updated data at application launch" OFF too (2026-10-11).
        // A manual check runs then as well, and the spinning icon - its tooltip saying how far the
        // check got - is what is left of it on screen once its card has been closed. Refusing here
        // also unbalanced the count: the matching Stop still took one off, ending the spinner under
        // another check still running. So every Stop is paired with a Start that ran (see
        // StartBackgroundSyncAsync).
        // ###########################################################################################
        internal void StartDataSyncStatusIconSpin()
        {
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
        internal void StopDataSyncStatusIconSpin()
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

        // Whether a check is running by the spinner's count - for tests (and what the icon's click asks).
        internal bool IsDataSyncStatusIconSpinning => this._dataSyncStatusIconSpinRequestCount > 0;

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

    }
}
