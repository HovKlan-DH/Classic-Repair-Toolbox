using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using DataHandling;
using System;
using System.Diagnostics;
using System.IO;

namespace CRT
{
    public partial class TabConfiguration : UserControl
    {
        public TabConfiguration()
        {
            this.InitializeComponent();

            // Determine if Dark Theme is actively evaluated during startup.
            var isDark = Application.Current?.RequestedThemeVariant == Avalonia.Styling.ThemeVariant.Dark ||
                         (Application.Current?.RequestedThemeVariant == Avalonia.Styling.ThemeVariant.Default &&
                          Application.Current?.ActualThemeVariant == Avalonia.Styling.ThemeVariant.Dark);

            this.ThemeToggleSwitch.IsChecked = isDark;
            this.MultipleInstancesForComponentPopupToggleSwitch.IsChecked = UserSettings.MultipleInstancesForComponentPopup;

            // Initialize configuration checkboxes — subscribe after setting initial values
            // to avoid triggering redundant saves during startup
            this.CheckVersionOnLaunchCheckBox.IsChecked = UserSettings.CheckVersionOnLaunch;
            this.CheckDataOnLaunchCheckBox.IsChecked = UserSettings.CheckDataOnLaunch;
            this.ShowDevelopmentVersionNotificationCheckBox.IsChecked = UserSettings.ShowDevelopmentVersionNotification;
            this.ValidateDataOnLaunchCheckBox.IsChecked = UserSettings.ValidateDataOnLaunch;
            this.DebugLoggingCheckBox.IsChecked = UserSettings.DebugLogging;

            this.ThemeToggleSwitch.IsCheckedChanged += this.OnThemeToggleSwitchChanged;
            this.MultipleInstancesForComponentPopupToggleSwitch.IsCheckedChanged += this.OnMultipleInstancesForComponentPopupChanged;
            this.CheckVersionOnLaunchCheckBox.IsCheckedChanged += this.OnCheckVersionOnLaunchChanged;
            this.CheckDataOnLaunchCheckBox.IsCheckedChanged += this.OnCheckDataOnLaunchChanged;
            this.ShowDevelopmentVersionNotificationCheckBox.IsCheckedChanged += this.OnShowDevelopmentVersionNotificationChanged;
            this.ValidateDataOnLaunchCheckBox.IsCheckedChanged += this.OnValidateDataOnLaunchChanged;
            this.DebugLoggingCheckBox.IsCheckedChanged += this.OnDebugLoggingChanged;
        }

        // ###########################################################################################
        // Applies and persists the selected application theme dynamically.
        // ###########################################################################################
        private void OnThemeToggleSwitchChanged(object? sender, RoutedEventArgs e)
        {
            var isDark = this.ThemeToggleSwitch.IsChecked == true;
            var newVariant = isDark ? Avalonia.Styling.ThemeVariant.Dark : Avalonia.Styling.ThemeVariant.Light;

            if (Application.Current != null)
            {
                Application.Current.RequestedThemeVariant = newVariant;
            }

            UserSettings.ThemeVariant = isDark ? "Dark" : "Light";
        }

        // ###########################################################################################
        // Persists the "Multiple instances for component popup" preference when the toggle is changed.
        // ###########################################################################################
        private void OnMultipleInstancesForComponentPopupChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.MultipleInstancesForComponentPopup = this.MultipleInstancesForComponentPopupToggleSwitch.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Check for new version at launch" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnCheckVersionOnLaunchChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.CheckVersionOnLaunch = this.CheckVersionOnLaunchCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Check for new or updated data at launch" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnCheckDataOnLaunchChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.CheckDataOnLaunch = this.CheckDataOnLaunchCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Show notification for DEVELOPMENT versions" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnShowDevelopmentVersionNotificationChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.ShowDevelopmentVersionNotification = this.ShowDevelopmentVersionNotificationCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Validate data at application launch" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnValidateDataOnLaunchChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.ValidateDataOnLaunch = this.ValidateDataOnLaunchCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Persists the "Enable debug logging" preference when the checkbox is toggled.
        // ###########################################################################################
        private void OnDebugLoggingChanged(object? sender, RoutedEventArgs e)
        {
            UserSettings.DebugLogging = this.DebugLoggingCheckBox.IsChecked == true;
        }

        // ###########################################################################################
        // Opens the persistent AppData folder that contains the log and settings files.
        // ###########################################################################################
        private void OnOpenAppDataFolderClick(object? sender, RoutedEventArgs e)
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = Path.Combine(appData, AppConfig.AppFolderName);

            try
            {
                Directory.CreateDirectory(directory);

                if (OperatingSystem.IsWindows())
                {
                    Process.Start(new ProcessStartInfo("explorer.exe", $"\"{directory}\"")
                    {
                        UseShellExecute = true
                    });
                }
                else if (OperatingSystem.IsMacOS())
                {
                    Process.Start("open", directory);
                }
                else
                {
                    Process.Start("xdg-open", directory);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to open app data folder - [{directory}] - [{ex.Message}]");
            }
        }
    }
}