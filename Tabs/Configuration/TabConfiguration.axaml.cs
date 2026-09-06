using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
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

            this.PopulateWorklogCurrencyComboBox();

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
            this.WorklogCurrencyComboBox.SelectionChanged += this.OnWorklogCurrencySelectionChanged;
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
            string helpUrl = AppConfig.WikiPageUrl(AppConfig.WikiPageWorkbooks);

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
            string helpUrl = AppConfig.WikiPageUrl(AppConfig.WikiPageMiniPro);

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
        // Fills the currency drop-down from WorklogCurrency.Options and selects the row for the
        // stored code.
        //
        // Built in code rather than as ComboBoxItems in the markup like the Theme drop-down above:
        // there are 68 countries, the list is data that belongs beside the codes it maps to, and
        // hand-written markup rows would be a second copy of that table to keep sorted and in sync.
        //
        // The items are the Option values themselves, with DisplayMemberBinding rendering
        // DisplayName - so the selection handler reads a typed Option back rather than parsing a
        // country and a code out of a formatted string.
        //
        // Selection is set BEFORE the handler is subscribed (in the constructor), matching every
        // other control on this tab: a SelectionChanged raised while restoring the stored value
        // would write that same value straight back.
        // ###########################################################################################
        private void PopulateWorklogCurrencyComboBox()
        {
            this.WorklogCurrencyComboBox.DisplayMemberBinding =
                new Binding(nameof(WorklogCurrency.Option.DisplayName));
            this.WorklogCurrencyComboBox.ItemsSource = WorklogCurrency.Options;
            this.WorklogCurrencyComboBox.SelectedItem = WorklogCurrency.ResolveOption(UserSettings.WorklogCurrencyCode);
        }

        // ###########################################################################################
        // Persists the chosen country's currency CODE - not the country, which is only how the user
        // picks it (see WorklogCurrency). UserSettings raises WorklogCurrencyChanged from the setter,
        // so the Workbooks tab reprints its figures without this tab knowing it exists.
        // ###########################################################################################
        private void OnWorklogCurrencySelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.WorklogCurrencyComboBox.SelectedItem is not WorklogCurrency.Option option)
            {
                return;
            }

            UserSettings.WorklogCurrencyCode = option.Code;
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
        // The three "Open ... folder" buttons.
        //
        // These are three DIFFERENT folders, which is why they are three buttons rather than the
        // single "data/workbooks/log/settings" one they replaced: that button could only ever open
        // one of them, and named four things while opening the parent of two of them.
        //
        // Each resolves the path the app is ACTUALLY using rather than rebuilding the AppData
        // default, because "--data-root=" and "--workbooks-root=" can move the first two elsewhere.
        // Rebuilding the default would open a folder the app is not reading from, which is worse
        // than not offering the button - the user would be looking at stale files while reporting
        // that a change did not take effect.
        //
        // WHICH IS WHY AN UNRESOLVED ROOT IS REPORTED RATHER THAN GUESSED AT. An empty DataRoot or
        // WorkbookRoot means that layer never loaded - and the likeliest reason is a
        // "--workbooks-root=" pointing somewhere currently unreachable, an external drive being the
        // obvious case. Substituting the AppData default there would create an empty folder the app
        // is not using and present it as the user's own, which reads as "my workbooks are gone"
        // rather than as "that drive is not mounted". Naming the problem is the smaller harm.
        // ###########################################################################################
        private async void OnOpenDataFolderClick(object? sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(DataManager.DataRoot))
            {
                Logger.Warning("Open data folder: the data root is not resolved, so there is no folder to open");
                await this.ShowFolderUnavailableDialogAsync("data");
                return;
            }

            await this.OpenFolderOrExplainAsync(DataManager.DataRoot, "data");
        }

        private async void OnOpenWorkbooksFolderClick(object? sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(WorklogManager.WorkbookRoot))
            {
                Logger.Warning("Open workbooks folder: the workbook root is not resolved, so there is no folder to open");
                await this.ShowFolderUnavailableDialogAsync("workbooks");
                return;
            }

            await this.OpenFolderOrExplainAsync(WorklogManager.WorkbookRoot, "workbooks");
        }

        // The log, the crash log and the settings file all sit directly in the AppData folder, so
        // this one is the folder itself rather than anything below it.
        private async void OnOpenLogsFolderClick(object? sender, RoutedEventArgs e)
        {
            await this.OpenFolderOrExplainAsync(ResolveAppDataFolder(), "logs and settings");
        }

        // ###########################################################################################
        // The persistent AppData folder that survives Velopack updates - the parent of the default
        // data and workbook folders, and the home of the log, crash log and settings files.
        // ###########################################################################################
        private static string ResolveAppDataFolder()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            return Path.Combine(appData, AppConfig.AppFolderName);
        }

        // ###########################################################################################
        // Opens one folder, creating it first so a button never fails merely because nothing has
        // been written there yet, and falling back to a dialog naming the path when the platform's
        // file manager cannot be launched. "description" names the folder in the log line only.
        // ###########################################################################################
        private async System.Threading.Tasks.Task OpenFolderOrExplainAsync(string directory, string description)
        {
            try
            {
                Directory.CreateDirectory(directory);

                if (TryOpenDirectory(directory, out string failureDetails))
                {
                    return;
                }

                Logger.Warning(
                    $"Failed to open {description} folder: [{directory}] - launcher details: [{failureDetails}]");

                await this.ShowOpenAppDataFolderFailedDialogAsync(directory);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open {description} folder: [{directory}] - [{ex.Message}]");
                await this.ShowOpenAppDataFolderFailedDialogAsync(directory);
            }
        }

        // ###########################################################################################
        // Shown when a folder cannot be opened because the app never resolved where it is - an empty
        // DataManager.DataRoot or WorklogManager.WorkbookRoot.
        //
        // Deliberately says nothing about a path: the whole point is that there ISN'T one to name,
        // and printing the AppData default here would send the user to the same wrong folder that
        // opening it would have. The log is where the reason actually is, which is why it is the
        // thing pointed at.
        // ###########################################################################################
        private async System.Threading.Tasks.Task ShowFolderUnavailableDialogAsync(string description)
        {
            await this.ShowFolderDialogAsync(
                "Folder not available",
                $"The {description} folder could not be opened, because the application has not resolved where it is.",
                "This usually means the folder was moved with a command-line parameter and is not reachable right now - " +
                "an external drive that is not connected, for example. The log file records the folder it tried to use.",
                path: null);
        }

        // ###########################################################################################
        // Shows a dialog with the folder's path when the path is known but the platform's file
        // manager could not be launched for it.
        // ###########################################################################################
        private async System.Threading.Tasks.Task ShowOpenAppDataFolderFailedDialogAsync(string directory)
        {
            await this.ShowFolderDialogAsync(
                "Unable to open folder",
                "The application could not open the folder automatically.",
                "You can open it manually using this path:",
                directory);
        }

        // ###########################################################################################
        // The shared body of the two dialogs above: a title, two lines of explanation, and - when
        // there is a path worth showing - a selectable copy of it.
        // ###########################################################################################
        private async System.Threading.Tasks.Task ShowFolderDialogAsync(
            string title,
            string firstLine,
            string secondLine,
            string? path)
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
                Title = title,
                Width = 520,
                MinWidth = 420,
                CanResize = false,
                ShowInTaskbar = false,
                SizeToContent = SizeToContent.Height,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            closeButton.Click += (_, _) => dialog.Close();

            var body = new StackPanel { Spacing = 14 };

            body.Children.Add(new TextBlock
            {
                Text = firstLine,
                TextWrapping = TextWrapping.Wrap
            });

            body.Children.Add(new TextBlock
            {
                Text = secondLine,
                TextWrapping = TextWrapping.Wrap
            });

            // Only when there is a real path. An empty SelectableTextBlock would render as a blank
            // gap the user reads as something failing to load.
            if (!string.IsNullOrEmpty(path))
            {
                body.Children.Add(new SelectableTextBlock
                {
                    Text = path,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            body.Children.Add(closeButton);

            dialog.Content = new Border
            {
                Padding = new Thickness(18),
                Child = body
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