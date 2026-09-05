using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Handlers.DataHandling;
using System;
using System.Diagnostics;
using System.IO;

namespace CRT
{
    public partial class TabConfiguration : UserControl
    {
        private bool thisSuppressCheckDataOnLaunchChanged;

        public TabConfiguration()
        {
            this.InitializeComponent();

            this.ThemeVariantComboBox.SelectedIndex = UserSettings.ThemeVariant switch
            {
                "Dark" => 1,
                "UserPreference" => 2,
                _ => 0
            };

            // Initialize configuration checkboxes — subscribe after setting initial values
            // to avoid triggering redundant saves during startup
            this.CheckVersionOnLaunchCheckBox.IsChecked = UserSettings.CheckVersionOnLaunch;
            this.CheckDataOnLaunchCheckBox.IsChecked = UserSettings.CheckDataOnLaunch;
            this.AllowDeletionOfOrphanAndNonUsedFilesCheckBox.IsChecked =
                UserSettings.AllowDeletionOfOrphanAndNonUsedFiles;
            this.ShowDevelopmentVersionNotificationCheckBox.IsChecked = UserSettings.ShowDevelopmentVersionNotification;
            this.MultipleInstancesForComponentPopupCheckBox.IsChecked = UserSettings.MultipleInstancesForComponentPopup;
            this.EnableNetworkConnectedOscilloscopeTabCheckBox.IsChecked = UserSettings.EnableNetworkConnectedOscilloscopeTab;
            this.EnableMiniproExperimentalModeCheckBox.IsChecked = UserSettings.EnableMiniproExperimentalMode;
            this.EnableMiniproExperimentalDemoModeCheckBox.IsChecked = UserSettings.EnableMiniproExperimentalDemoMode;
            this.UpdateEnableMiniproExperimentalDemoModeCheckBoxState();
            this.EnableWorklogCheckBox.IsChecked = UserSettings.EnableWorklog;

            bool isAllBoardsScope = string.Equals(UserSettings.WorkbooksScope, "AllBoards", StringComparison.Ordinal);
            this.WorkbooksScopeAllBoardsRadioButton.IsChecked = isAllBoardsScope;
            this.WorkbooksScopeCurrentBoardRadioButton.IsChecked = !isAllBoardsScope;

            this.EnableMiniproExperimentalDemoModeCheckBox.IsCheckedChanged += this.OnEnableMiniproExperimentalDemoModeChanged;

            this.UpdateAllowDeletionOfOrphanAndNonUsedFilesCheckBoxState();

            this.DownloadDataFromTestSourceCheckBox.IsVisible = true;
            this.DownloadDataFromTestSourceCheckBox.IsChecked = UserSettings.DownloadDataFromTestSource;
            this.UpdateDownloadDataFromTestSourceCheckBoxState();

            this.ThemeVariantComboBox.SelectionChanged += this.OnThemeVariantSelectionChanged;
            this.CheckVersionOnLaunchCheckBox.IsCheckedChanged += this.OnCheckVersionOnLaunchChanged;
            this.CheckDataOnLaunchCheckBox.IsCheckedChanged += this.OnCheckDataOnLaunchChanged;
            this.AllowDeletionOfOrphanAndNonUsedFilesCheckBox.IsCheckedChanged += this.OnAllowDeletionOfOrphanAndNonUsedFilesChanged;
            this.DownloadDataFromTestSourceCheckBox.IsCheckedChanged += this.OnDownloadDataFromTestSourceChanged;
            this.ShowDevelopmentVersionNotificationCheckBox.IsCheckedChanged += this.OnShowDevelopmentVersionNotificationChanged;
            this.MultipleInstancesForComponentPopupCheckBox.IsCheckedChanged += this.OnMultipleInstancesForComponentPopupChanged;
            this.EnableNetworkConnectedOscilloscopeTabCheckBox.IsCheckedChanged += this.OnEnableNetworkConnectedOscilloscopeTabChanged;
            this.EnableMiniproExperimentalModeCheckBox.IsCheckedChanged += this.OnEnableMiniproExperimentalModeChanged;
            this.EnableWorklogCheckBox.IsCheckedChanged += this.OnEnableWorklogChanged;
            this.WorkbooksScopeAllBoardsRadioButton.IsCheckedChanged += this.OnWorkbooksScopeChanged;
            this.WorkbooksScopeCurrentBoardRadioButton.IsCheckedChanged += this.OnWorkbooksScopeChanged;
        }

        // ###########################################################################################
        // Keeps the "Download data from BETA source" checkbox enabled only while launch-time data sync
        // is enabled, without changing the stored source preference.
        // ###########################################################################################
        private void UpdateDownloadDataFromTestSourceCheckBoxState()
        {
            bool isCheckDataOnLaunchEnabled = this.CheckDataOnLaunchCheckBox.IsChecked == true;
            this.DownloadDataFromTestSourceCheckBox.IsEnabled = isCheckDataOnLaunchEnabled;
        }

        // ###########################################################################################
        // Persists whether orphan and non-used files may be deleted from the data root.
        // The cleanup can only run when launch-time data synchronization is enabled.
        // ###########################################################################################
        private void OnAllowDeletionOfOrphanAndNonUsedFilesChanged(object? sender, RoutedEventArgs e)
        {
            bool isEnabled = this.AllowDeletionOfOrphanAndNonUsedFilesCheckBox.IsChecked == true;
            UserSettings.AllowDeletionOfOrphanAndNonUsedFiles = isEnabled;

            if (!isEnabled || !UserSettings.CheckDataOnLaunch)
            {
                return;
            }

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                mainWindow.ScheduleOrphanAndUnusedFileCleanupIfEnabled();
            }
        }

        // ###########################################################################################
        // Keeps the orphan/non-used file deletion checkbox enabled only while launch-time data sync
        // is enabled, without changing the stored deletion preference.
        // ###########################################################################################
        private void UpdateAllowDeletionOfOrphanAndNonUsedFilesCheckBoxState()
        {
            bool isCheckDataOnLaunchEnabled = this.CheckDataOnLaunchCheckBox.IsChecked == true;
            this.AllowDeletionOfOrphanAndNonUsedFilesCheckBox.IsEnabled = isCheckDataOnLaunchEnabled;
        }

        // ###########################################################################################
        // Applies and persists the selected application theme from the drop-down list.
        // ###########################################################################################
        private void OnThemeVariantSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            var themeVariant = this.ThemeVariantComboBox.SelectedIndex switch
            {
                1 => "Dark",
                2 => "UserPreference",
                _ => "Light"
            };

            UserSettings.ThemeVariant = themeVariant;

            if (Application.Current is App app)
            {
                app.ApplyConfiguredTheme();
            }
        }

        // ###########################################################################################
        // Persists the "Check for new version at launch" preference when the checkbox is toggled.
        // Enabling it also performs the check immediately.
        // ###########################################################################################
        private async void OnCheckVersionOnLaunchChanged(object? sender, RoutedEventArgs e)
        {
            bool isEnabled = this.CheckVersionOnLaunchCheckBox.IsChecked == true;
            UserSettings.CheckVersionOnLaunch = isEnabled;

            if (!isEnabled)
            {
                return;
            }

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                await mainWindow.CheckForAppUpdateNowAsync();
            }
        }

        // ###########################################################################################
        // Persists the "Check for new or updated data at launch" preference when the checkbox is
        // toggled. Enabling it also performs the check immediately.
        // ###########################################################################################
        private async void OnCheckDataOnLaunchChanged(object? sender, RoutedEventArgs e)
        {
            if (this.thisSuppressCheckDataOnLaunchChanged)
            {
                return;
            }

            bool isEnabled = this.CheckDataOnLaunchCheckBox.IsChecked == true;
            UserSettings.CheckDataOnLaunch = isEnabled;
            this.UpdateAllowDeletionOfOrphanAndNonUsedFilesCheckBoxState();
            this.UpdateDownloadDataFromTestSourceCheckBoxState();

            if (!isEnabled)
            {
                return;
            }

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                await mainWindow.CheckForDataUpdatesNowAsync();
                mainWindow.ScheduleOrphanAndUnusedFileCleanupIfEnabled();
            }
        }

        // ###########################################################################################
        // Updates the "Check for new or updated data at launch" checkbox from outside this tab
        // without triggering its persistence handler a second time.
        // ###########################################################################################
        public void SetCheckDataOnLaunchCheckBoxValue(bool isChecked)
        {
            this.thisSuppressCheckDataOnLaunchChanged = true;
            this.CheckDataOnLaunchCheckBox.IsChecked = isChecked;
            this.UpdateAllowDeletionOfOrphanAndNonUsedFilesCheckBoxState();
            this.UpdateDownloadDataFromTestSourceCheckBoxState();
            this.thisSuppressCheckDataOnLaunchChanged = false;
        }

        // ###########################################################################################
        // Persists the "Show notification for DEVELOPMENT versions" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnShowDevelopmentVersionNotificationChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.ShowDevelopmentVersionNotification = this.ShowDevelopmentVersionNotificationCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Enable network connected oscilloscope tab" preference and shows or hides the
        // "Oscilloscope" tab in the main window to match it.
        // ###########################################################################################
        private void OnEnableNetworkConnectedOscilloscopeTabChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.EnableNetworkConnectedOscilloscopeTab =
                this.EnableNetworkConnectedOscilloscopeTabCheckBox.IsChecked == true;

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                mainWindow.ApplyOscilloscopeTabVisibility();
            }
        }

        // ###########################################################################################
        // Enables the Minipro demo mode checkbox only while Minipro experimental mode is enabled.
        // ###########################################################################################
        private void UpdateEnableMiniproExperimentalDemoModeCheckBoxState()
        {
            bool isExperimentalModeEnabled = this.EnableMiniproExperimentalModeCheckBox.IsChecked == true;
            this.EnableMiniproExperimentalDemoModeCheckBox.IsEnabled = isExperimentalModeEnabled;
        }

        // ###########################################################################################
        // Opens the help page describing the Workbooks tab, through the shared launcher that every
        // external target in this app goes through - same shape as the MiniPro help below it.
        // ###########################################################################################
        private void OnEnableWorklogHelpClick(object? sender, RoutedEventArgs e)
        {
            const string helpUrl = "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/blob/main/docs/workbooks.md";

            if (!ExternalTargetLauncher.TryOpen(helpUrl))
            {
                Logger.Warning($"Rejected external target from Configuration tab: [{helpUrl}]");
            }
        }

        // ###########################################################################################
        // Opens the help page describing the MiniPro programmer, through the shared launcher that
        // every external target in this app goes through.
        // ###########################################################################################
        private void OnEnableMiniproExperimentalModeHelpClick(object? sender, RoutedEventArgs e)
        {
            const string helpUrl = "https://github.com/HovKlan-DH/Classic-Repair-Toolbox/blob/main/docs/minipro-programmer.md";

            if (!ExternalTargetLauncher.TryOpen(helpUrl))
            {
                Logger.Warning($"Rejected external target from Configuration tab: [{helpUrl}]");
            }
        }

        // ###########################################################################################
        // Persists the "Enable MiniPro programmer functionality" preference and updates dependent UI.
        // ###########################################################################################
        private void OnEnableMiniproExperimentalModeChanged(object? sender, RoutedEventArgs e)
        {
            bool isEnabled = this.EnableMiniproExperimentalModeCheckBox.IsChecked == true;
            UserSettings.EnableMiniproExperimentalMode = isEnabled;

            if (!isEnabled)
            {
                this.EnableMiniproExperimentalDemoModeCheckBox.IsChecked = false;
                UserSettings.EnableMiniproExperimentalDemoMode = false;
            }

            this.UpdateEnableMiniproExperimentalDemoModeCheckBoxState();
        }

        // ###########################################################################################
        // Persists the "Enable Worklog" preference and shows or hides the worklog bar above the
        // tabs in the main window to match it.
        // ###########################################################################################
        private void OnEnableWorklogChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.EnableWorklog = this.EnableWorklogCheckBox.IsChecked == true;

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                mainWindow.ApplyWorklogBarVisibility();
            }
        }

        // ###########################################################################################
        // Persists which scope the Workbooks tab's workbook list uses - every board's workbooks, or
        // only the currently selected board's - when either radio button is toggled.
        //
        // One click raises IsCheckedChanged TWICE, once for the button being unchecked and once for
        // the one being checked, and BOTH see the same post-transition state. So the SENDER is what
        // decides here: only the button that just became checked writes, and the uncheck is ignored
        // outright. Reading the group's state instead (the previous form) happened to write the same
        // value twice and relied entirely on UserSettings.WorkbooksScope's unchanged-value guard to
        // absorb it - which now matters, because that setter raises WorkbooksScopeChanged and Main
        // rebuilds the whole Workbooks tab off it: a doubled write is a doubled full disk rescan and
        // schematic re-decode per click.
        // ###########################################################################################
        private void OnWorkbooksScopeChanged(object? sender, RoutedEventArgs e)
        {
            if (sender is not RadioButton radioButton || radioButton.IsChecked != true)
            {
                return;
            }

            if (ReferenceEquals(radioButton, this.WorkbooksScopeAllBoardsRadioButton))
            {
                UserSettings.WorkbooksScope = "AllBoards";
            }
            else if (ReferenceEquals(radioButton, this.WorkbooksScopeCurrentBoardRadioButton))
            {
                UserSettings.WorkbooksScope = "CurrentBoard";
            }
        }

        // ###########################################################################################
        // Persists the "Enable experimental demo mode for Minipro" preference when toggled.
        // ###########################################################################################
        private void OnEnableMiniproExperimentalDemoModeChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.EnableMiniproExperimentalDemoMode =
                this.EnableMiniproExperimentalDemoModeCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Opens the persistent AppData folder that contains the log and settings files.
        // ###########################################################################################
        private async void OnOpenAppDataFolderClick(object? sender, RoutedEventArgs e)
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = Path.Combine(appData, AppConfig.AppFolderName);

            try
            {
                Directory.CreateDirectory(directory);

                if (TryOpenDirectory(directory, out string failureDetails))
                {
                    return;
                }

                Logger.Warning(
                    $"Failed to open AppData folder: [{directory}] - launcher details: [{failureDetails}]");

                await this.ShowOpenAppDataFolderFailedDialogAsync(directory);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open AppData folder: [{directory}] - [{ex.Message}]");
                await this.ShowOpenAppDataFolderFailedDialogAsync(directory);
            }
        }

        // ###########################################################################################
        // Shows a dialog with the AppData folder path when automatic opening fails.
        // ###########################################################################################
        private async System.Threading.Tasks.Task ShowOpenAppDataFolderFailedDialogAsync(string directory)
        {
            if (TopLevel.GetTopLevel(this) is not Window owner)
            {
                return;
            }

            var closeButton = new Button
            {
                Content = "OK",
                MinWidth = 110,
                HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Center,
                HorizontalContentAlignment = Avalonia.Layout.HorizontalAlignment.Center
            };

            var dialog = new Window
            {
                Title = "Unable to open folder",
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
                    Text = "The application could not open the data/log/settings folder automatically.",
                    TextWrapping = TextWrapping.Wrap
                },
                new TextBlock
                {
                    Text = "You can open it manually using this path:",
                    TextWrapping = TextWrapping.Wrap
                },
                new SelectableTextBlock
                {
                    Text = directory,
                    TextWrapping = TextWrapping.Wrap
                },
                closeButton
            }
                }
            };

            await dialog.ShowDialog(owner);
        }

        // ###########################################################################################
        // Reloads user-preference theme colors from settings and reapplies the configured theme.
        // ###########################################################################################
        private void OnReloadUserPreferenceThemeClick(object? sender, RoutedEventArgs e)
        {
            UserSettings.ReloadUserThemeColors();

            if (Application.Current is App app)
            {
                app.ApplyConfiguredTheme();
            }
        }

        // ###########################################################################################
        // Persists the "Open multiple component info windows" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnMultipleInstancesForComponentPopupChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.MultipleInstancesForComponentPopup =
                this.MultipleInstancesForComponentPopupCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Opens a directory with the platform-specific file manager command and returns failure details.
        // ###########################################################################################
        private static bool TryOpenDirectory(string directory, out string failureDetails)
        {
            var attempts = new System.Collections.Generic.List<string>();

            if (OperatingSystem.IsWindows())
            {
                bool success = TryStartShellTarget("explorer.exe", attempts, directory);
                failureDetails = string.Join(" | ", attempts);
                return success;
            }

            if (OperatingSystem.IsMacOS())
            {
                bool success = TryStartCommandWithDiagnostics("open", attempts, directory);
                failureDetails = string.Join(" | ", attempts);
                return success;
            }

            if (OperatingSystem.IsLinux())
            {
                if (TryStartCommandWithDiagnostics("xdg-open", attempts, directory))
                {
                    failureDetails = string.Join(" | ", attempts);
                    return true;
                }

                if (TryStartCommandWithDiagnostics("gio", attempts, "open", directory))
                {
                    failureDetails = string.Join(" | ", attempts);
                    return true;
                }

                failureDetails = string.Join(" | ", attempts);
                return false;
            }

            attempts.Add("Unsupported operating system");
            failureDetails = string.Join(" | ", attempts);
            return false;
        }

        // ###########################################################################################
        // Starts an external command with arguments and records diagnostic details for each attempt.
        // ###########################################################################################
        private static bool TryStartCommandWithDiagnostics(
            string fileName,
            System.Collections.Generic.List<string> attempts,
            params string[] arguments)
        {
            string argumentText = arguments.Length == 0 ? "(none)" : string.Join(" ", arguments);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    RedirectStandardOutput = true
                };

                foreach (var argument in arguments)
                {
                    psi.ArgumentList.Add(argument);
                }

                using var process = Process.Start(psi);

                if (process == null)
                {
                    attempts.Add($"Command [{fileName}] args [{argumentText}] did not start a process");
                    return false;
                }

                if (process.WaitForExit(2000))
                {
                    string standardOutput = process.StandardOutput.ReadToEnd().Trim();
                    string standardError = process.StandardError.ReadToEnd().Trim();

                    if (process.ExitCode == 0)
                    {
                        attempts.Add($"Command [{fileName}] args [{argumentText}] succeeded with exit code [0]");
                        return true;
                    }

                    attempts.Add(
                        $"Command [{fileName}] args [{argumentText}] failed with exit code [{process.ExitCode}] output [{standardOutput}] error [{standardError}]");
                    return false;
                }

                attempts.Add($"Command [{fileName}] args [{argumentText}] started and is still running");
                return true;
            }
            catch (Exception ex)
            {
                attempts.Add($"Command [{fileName}] args [{argumentText}] threw [{ex.GetType().Name}: {ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // Starts a shell target and treats successful process creation as success.
        // This is used for Windows Explorer where exit codes are not reliable for this scenario.
        // ###########################################################################################
        private static bool TryStartShellTarget(
            string fileName,
            System.Collections.Generic.List<string> attempts,
            params string[] arguments)
        {
            string argumentText = arguments.Length == 0 ? "(none)" : string.Join(" ", arguments);

            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = fileName,
                    UseShellExecute = false
                };

                foreach (var argument in arguments)
                {
                    psi.ArgumentList.Add(argument);
                }

                using var process = Process.Start(psi);

                if (process == null)
                {
                    attempts.Add($"Command [{fileName}] args [{argumentText}] did not start a process");
                    return false;
                }

                attempts.Add($"Command [{fileName}] args [{argumentText}] started successfully");
                return true;
            }
            catch (Exception ex)
            {
                attempts.Add($"Command [{fileName}] args [{argumentText}] threw [{ex.GetType().Name}: {ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // Persists whether data should be fetched from the BETA manifest source instead of
        // the production source. Enabling it performs an immediate refresh so the selected source
        // takes effect right away, while disabling it applies on next application launch.
        // ###########################################################################################
        private async void OnDownloadDataFromTestSourceChanged(object? sender, RoutedEventArgs e)
        {
            bool isEnabled = this.DownloadDataFromTestSourceCheckBox.IsChecked == true;
            UserSettings.DownloadDataFromTestSource = isEnabled;

            if (!UserSettings.CheckDataOnLaunch)
            {
                this.UpdateDownloadDataFromTestSourceCheckBoxState();
                return;
            }

            if (!isEnabled)
            {
                return;
            }

            if (TopLevel.GetTopLevel(this) is Main mainWindow)
            {
                await mainWindow.CheckForDataUpdatesNowAsync();
            }
        }





    }
}