using CRT;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Persisted user preferences model. Defaults to enabled for all online features.
    // ###########################################################################################
    internal sealed class UserSettingsData
    {
        [JsonPropertyName("checkVersionOnLaunch")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? CheckVersionOnLaunch { get; set; }

        [JsonPropertyName("checkDataOnLaunch")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? CheckDataOnLaunch { get; set; }

        [JsonPropertyName("showDevelopmentVersionNotification")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? ShowDevelopmentVersionNotification { get; set; }

        [JsonPropertyName("validateDataOnLaunch")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? ValidateDataOnLaunch { get; set; }

        [JsonPropertyName("debugLogging")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? DebugLogging { get; set; }

        [JsonPropertyName("multipleInstancesForComponentPopup")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? MultipleInstancesForComponentPopup { get; set; }

        [JsonPropertyName("leftPanelWidth")] public double LeftPanelWidth { get; set; } = 200.0;
        [JsonPropertyName("schematicsSplitterRatios")] public Dictionary<string, double> SchematicsSplitterRatios { get; set; } = new();
        [JsonPropertyName("selectedCategoriesByBoard")] public Dictionary<string, List<string>> SelectedCategoriesByBoard { get; set; } = new();
        [JsonPropertyName("lastHardware")] public string LastHardware { get; set; } = string.Empty;
        [JsonPropertyName("lastBoardByHardware")] public Dictionary<string, string> LastBoardByHardware { get; set; } = new();
        [JsonPropertyName("lastSchematicByBoard")] public Dictionary<string, string> LastSchematicByBoard { get; set; } = new();
        [JsonPropertyName("region")] public string Region { get; set; } = "PAL";
        [JsonPropertyName("theme")] public string ThemeVariant { get; set; } = "Default";
        [JsonPropertyName("hasWindowPlacement")] public bool HasWindowPlacement { get; set; } = false;
        [JsonPropertyName("windowState")] public string WindowState { get; set; } = "Normal";
        [JsonPropertyName("windowWidth")] public double WindowWidth { get; set; } = 1024.0;
        [JsonPropertyName("windowHeight")] public double WindowHeight { get; set; } = 768.0;
        [JsonPropertyName("windowX")] public int WindowX { get; set; } = 0;
        [JsonPropertyName("windowY")] public int WindowY { get; set; } = 0;
        [JsonPropertyName("windowScreenX")] public int WindowScreenX { get; set; } = 0;
        [JsonPropertyName("windowScreenY")] public int WindowScreenY { get; set; } = 0;
        [JsonPropertyName("windowScreenWidth")] public int WindowScreenWidth { get; set; } = 1920;
        [JsonPropertyName("windowScreenHeight")] public int WindowScreenHeight { get; set; } = 1080;
        [JsonPropertyName("windowScreenScaling")] public double WindowScreenScaling { get; set; } = 1.0;
        [JsonPropertyName("hasComponentInfoWindowLayout")] public bool HasComponentInfoWindowLayout { get; set; } = false;
        [JsonPropertyName("componentInfoWindowState")] public string ComponentInfoWindowState { get; set; } = "Normal";
        [JsonPropertyName("componentInfoWindowWidth")] public double ComponentInfoWindowWidth { get; set; } = 680.0;
        [JsonPropertyName("componentInfoWindowHeight")] public double ComponentInfoWindowHeight { get; set; } = 420.0;
        [JsonPropertyName("componentInfoWindowLeftColumnRatio")] public double ComponentInfoWindowLeftColumnRatio { get; set; } = 0.5;
        [JsonPropertyName("componentInfoWindowThumbnailRowHeight")] public double ComponentInfoWindowThumbnailRowHeight { get; set; } = 100.0;
        [JsonPropertyName("componentInfoScrollAction")] public string ComponentInfoScrollAction { get; set; } = "Image change";
        [JsonPropertyName("schematicsLabelBoard")] public bool SchematicsLabelBoard { get; set; } = false;
        [JsonPropertyName("schematicsLabelTechnical")] public bool SchematicsLabelTechnical { get; set; } = false;
        [JsonPropertyName("schematicsLabelFriendly")] public bool SchematicsLabelFriendly { get; set; } = false;
        [JsonPropertyName("schematicsLabelSelectedOnly")] public bool SchematicsLabelSelectedOnly { get; set; } = false;
        [JsonPropertyName("schematicsLabelsPanelExpanded")] public bool SchematicsLabelsPanelExpanded { get; set; } = true;
        [JsonPropertyName("blinkSelected")] public bool BlinkSelected { get; set; } = false;
        [JsonPropertyName("schematicsOrderByBoard")] public Dictionary<string, List<string>> SchematicsOrderByBoard { get; set; } = new();
        [JsonPropertyName("contactEmail")] public string ContactEmail { get; set; } = string.Empty;

        [JsonPropertyName("lastOscilloscopeVendor")] public string LastOscilloscopeVendor { get; set; } = string.Empty;
        [JsonPropertyName("lastOscilloscopeSeriesByVendor")] public Dictionary<string, string> LastOscilloscopeSeriesByVendor { get; set; } = new();
        [JsonPropertyName("oscilloscopeHost")] public string OscilloscopeHost { get; set; } = "192.168.0.100";
        [JsonPropertyName("oscilloscopePort")] public int OscilloscopePort { get; set; } = 5025;
        [JsonPropertyName("componentInfoKeyboardHandling")] public string ComponentInfoKeyboardHandling { get; set; } = "Control image pin selection";

    }

    // ###########################################################################################
    // Loads and saves user preferences to a JSON file in the same folder as the log file.
    // Call Load() once at startup before any settings are read.
    // ###########################################################################################
    public static class UserSettings
    {
        private static UserSettingsData _data = new();
        private static string _settingsFilePath = string.Empty;

        public static bool CheckVersionOnLaunch
        {
            get => _data.CheckVersionOnLaunch ?? true;
            set
            {
                _data.CheckVersionOnLaunch = value;
                Logger.Info($"Setting changed: [CheckVersionOnLaunch] [{value}]");
                Save();
            }
        }

        public static string ComponentInfoKeyboardHandling
        {
            get => _data.ComponentInfoKeyboardHandling;
            set
            {
                _data.ComponentInfoKeyboardHandling = value;
                Logger.Info($"Setting changed: [ComponentInfoKeyboardHandling] [{value}]");
                Save();
            }
        }

        public static bool CheckDataOnLaunch
        {
            get => _data.CheckDataOnLaunch ?? true;
            set
            {
                _data.CheckDataOnLaunch = value;
                Logger.Info($"Setting changed: [CheckDataOnLaunch] [{value}]");
                Save();
            }
        }

        public static bool ShowDevelopmentVersionNotification
        {
            get => _data.ShowDevelopmentVersionNotification ?? false; // Default is false
            set
            {
                _data.ShowDevelopmentVersionNotification = value;
                Logger.Info($"Setting changed: [ShowDevelopmentVersionNotification] [{value}]");
                Save();
            }
        }

        public static bool ValidateDataOnLaunch
        {
            get => _data.ValidateDataOnLaunch ?? false;
            set
            {
                _data.ValidateDataOnLaunch = value;
                Logger.Info($"Setting changed: [ValidateDataOnLaunch] [{value}]");
                Save();
            }
        }

        public static bool DebugLogging
        {
            get => _data.DebugLogging ?? false;
            set
            {
                _data.DebugLogging = value;
                Logger.IsDebugEnabled = value;
                Logger.Info($"Setting changed: [DebugLogging] [{value}]");
                Save();
            }
        }

        public static bool MultipleInstancesForComponentPopup
        {
            get => _data.MultipleInstancesForComponentPopup ?? false;
            set
            {
                _data.MultipleInstancesForComponentPopup = value;
                Logger.Info($"Setting changed: [MultipleInstancesForComponentPopup] [{value}]");
                Save();
            }
        }

        public static double LeftPanelWidth
        {
            get => _data.LeftPanelWidth;
            set
            {
                _data.LeftPanelWidth = value;
                Logger.Info($"Setting changed: [LeftPanelWidth] [{value:F1}]");
                Save();
            }
        }

        public static string ThemeVariant
        {
            get => _data.ThemeVariant;
            set
            {
                _data.ThemeVariant = value;
                Logger.Info($"Setting changed: [Theme] [{value}]");
                Save();
            }
        }

        public static string ContactEmail
        {
            get => _data.ContactEmail ?? string.Empty;
            set
            {
                string normalized = value?.Trim() ?? string.Empty;

                if (string.Equals(_data.ContactEmail, normalized, StringComparison.Ordinal))
                    return;

                _data.ContactEmail = normalized;
                Logger.Info($"Setting changed: [ContactEmail] [{(string.IsNullOrWhiteSpace(normalized) ? "empty" : "set")}]");
                Save();
            }
        }

        public static string Region
        {
            get => _data.Region;
            set
            {
                _data.Region = value;
                Logger.Info($"Setting changed: [Region] [{value}]");
                Save();
            }
        }

        public static bool SchematicsLabelBoard
        {
            get => _data.SchematicsLabelBoard;
            set { _data.SchematicsLabelBoard = value; Save(); }
        }

        public static bool SchematicsLabelTechnical
        {
            get => _data.SchematicsLabelTechnical;
            set { _data.SchematicsLabelTechnical = value; Save(); }
        }

        public static bool SchematicsLabelFriendly
        {
            get => _data.SchematicsLabelFriendly;
            set { _data.SchematicsLabelFriendly = value; Save(); }
        }

        public static bool SchematicsLabelSelectedOnly
        {
            get => _data.SchematicsLabelSelectedOnly;
            set { _data.SchematicsLabelSelectedOnly = value; Save(); }
        }

        public static bool SchematicsLabelsPanelExpanded
        {
            get => _data.SchematicsLabelsPanelExpanded;
            set { _data.SchematicsLabelsPanelExpanded = value; Save(); }
        }

        public static bool BlinkSelected
        {
            get => _data.BlinkSelected;
            set
            {
                _data.BlinkSelected = value;
                Logger.Info($"Setting changed: [BlinkSelected] [{value}]");
                Save();
            }
        }

        // Window placement — read-only; written atomically via SaveWindowPlacement
        public static bool HasWindowPlacement => _data.HasWindowPlacement;
        public static string WindowState => _data.WindowState;
        public static double WindowWidth => _data.WindowWidth;
        public static double WindowHeight => _data.WindowHeight;
        public static int WindowX => _data.WindowX;
        public static int WindowY => _data.WindowY;
        public static int WindowScreenX => _data.WindowScreenX;
        public static int WindowScreenY => _data.WindowScreenY;
        public static int WindowScreenWidth => _data.WindowScreenWidth;
        public static int WindowScreenHeight => _data.WindowScreenHeight;
        public static double WindowScreenScaling => _data.WindowScreenScaling;

        // Component info window layout — read-only; written atomically via SaveComponentInfoWindowLayout
        public static bool HasComponentInfoWindowLayout => _data.HasComponentInfoWindowLayout;
        public static string ComponentInfoWindowState => _data.ComponentInfoWindowState;
        public static double ComponentInfoWindowWidth => _data.ComponentInfoWindowWidth;
        public static double ComponentInfoWindowHeight => _data.ComponentInfoWindowHeight;
        public static double ComponentInfoWindowLeftColumnRatio => _data.ComponentInfoWindowLeftColumnRatio;
        public static double ComponentInfoWindowThumbnailRowHeight => _data.ComponentInfoWindowThumbnailRowHeight;

        public static string ComponentInfoScrollAction
        {
            get => _data.ComponentInfoScrollAction;
            set
            {
                _data.ComponentInfoScrollAction = value;
                Logger.Info($"Setting changed: [ComponentInfoScrollAction] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Saves component info window layout values atomically in a single disk write.
        // ###########################################################################################
        public static void SaveComponentInfoWindowLayout(string state, double width, double height, double leftColumnRatio, double thumbnailRowHeight)
        {
            _data.HasComponentInfoWindowLayout = true;
            _data.ComponentInfoWindowState = state;
            _data.ComponentInfoWindowWidth = width;
            _data.ComponentInfoWindowHeight = height;
            _data.ComponentInfoWindowLeftColumnRatio = leftColumnRatio;
            _data.ComponentInfoWindowThumbnailRowHeight = thumbnailRowHeight;
            Logger.Info($"Setting changed: [ComponentInfoWindowLayout] [{state}] [{width:F0}x{height:F0}] [LeftRatio: {leftColumnRatio:F3}] [ThumbnailHeight: {thumbnailRowHeight:F1}]");
            Save();
        }

        // ###########################################################################################
        // Returns the saved schematics splitter ratio for the given board key.
        // Defaults to 0.5 (equal split) when no saved value exists.
        // ###########################################################################################
        public static double GetSchematicsSplitterRatio(string boardKey)
            => _data.SchematicsSplitterRatios.TryGetValue(boardKey, out var ratio) ? ratio : 0.5;

        // ###########################################################################################
        // Persists the schematics splitter ratio for the given board key.
        // ###########################################################################################
        public static void SetSchematicsSplitterRatio(string boardKey, double ratio)
        {
            _data.SchematicsSplitterRatios[boardKey] = ratio;
            Logger.Info($"Setting changed: [SchematicsSplitterRatio] [{boardKey}] [{ratio:F3}]");
            Save();
        }

        // ###########################################################################################
        // Returns the saved selected categories for the given board key.
        // Returns null when no selection has been saved yet (caller should default to all selected).
        // ###########################################################################################
        public static List<string>? GetSelectedCategories(string boardKey)
            => _data.SelectedCategoriesByBoard.TryGetValue(boardKey, out var categories) ? categories : null;

        // ###########################################################################################
        // Persists the selected category list for the given board key.
        // ###########################################################################################
        public static void SetSelectedCategories(string boardKey, List<string> categories)
        {
            _data.SelectedCategoriesByBoard[boardKey] = categories;
            Logger.Info($"Setting changed: [SelectedCategories] [{boardKey}] [{categories.Count} selected]");
            Save();
        }

        // ###########################################################################################
        // Saves all window placement values atomically in a single disk write.
        // ###########################################################################################
        public static void SaveWindowPlacement(string state, double width, double height, int x, int y, int screenX, int screenY, int screenWidth, int screenHeight, double screenScaling)
        {
            _data.HasWindowPlacement = true;
            _data.WindowState = state;
            _data.WindowWidth = width;
            _data.WindowHeight = height;
            _data.WindowX = x;
            _data.WindowY = y;
            _data.WindowScreenX = screenX;
            _data.WindowScreenY = screenY;
            _data.WindowScreenWidth = screenWidth;
            _data.WindowScreenHeight = screenHeight;
            _data.WindowScreenScaling = screenScaling;
            Logger.Info($"Setting changed: [WindowPlacement] [{state}] [{width:F0}x{height:F0}] [Pos: {x},{y}] [Screen: {screenX},{screenY} {screenWidth}x{screenHeight} @{screenScaling:F2}x]");
            Save();
        }

        // ###########################################################################################
        // Resolves the settings file path and loads persisted values.
        // Falls back to defaults silently on any failure.
        // ###########################################################################################
        public static void Load()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var directory = Path.Combine(appData, AppConfig.AppFolderName);
                Directory.CreateDirectory(directory);
                _settingsFilePath = Path.Combine(directory, AppConfig.SettingsFileName);

                // Get and log system information completely independently of the configuration state
                var os = RuntimeInformation.OSDescription;
                var osHighLevel = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? "Windows"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "macOS"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.Linux) ? "Linux"
                    : RuntimeInformation.IsOSPlatform(OSPlatform.FreeBSD) ? "FreeBSD"
                    : "Unknown";

                var archDescription = RuntimeInformation.ProcessArchitecture switch
                {
                    Architecture.X64 => "64-bit",
                    Architecture.X86 => "32-bit",
                    Architecture.Arm64 => "ARM 64-bit",
                    Architecture.Arm => "ARM 32-bit",
                    var a => a.ToString()
                };

                Logger.Info($"System information:");
                Logger.Info($"    Operating system is [{osHighLevel}] version [{os}]");
                Logger.Info($"    Using CPU architecture [{archDescription}]");
                Logger.Info($"    Self-contained .NET Runtime used [{RuntimeInformation.FrameworkDescription}]");

                // Now evaluate settings
                if (!File.Exists(_settingsFilePath))
                {
                    Logger.Info("Configuration file not found - using defaults");
                    return;
                }

                var json = File.ReadAllText(_settingsFilePath);
                var loaded = JsonSerializer.Deserialize<UserSettingsData>(json);
                if (loaded != null)
                {
                    _data = loaded;
                    Logger.IsDebugEnabled = DebugLogging;

                    Logger.Info("Settings loaded:");
                    Logger.Info($"    Configuration:");
                    Logger.Info($"        [Theme] [{ThemeVariant}]");
                    Logger.Info($"        [OpenMultiplePopups] [{MultipleInstancesForComponentPopup}]");
                    Logger.Info($"        [CheckDataOnLaunch] [{CheckDataOnLaunch}]");
                    Logger.Info($"        [CheckVersionOnLaunch] [{CheckVersionOnLaunch}]");
                    Logger.Info($"        [AllowBetaNotification] [{ShowDevelopmentVersionNotification}]");
                    Logger.Info($"        [ValidateDataOnLaunch] [{ValidateDataOnLaunch}]");
                    Logger.Info($"        [DebugLogging] [{DebugLogging}]");
                    Logger.Info($"    Various other settings:");
                    Logger.Info($"        [BlinkSelected] [{BlinkSelected}]");
                    Logger.Info($"        [ComponentInfoWindowLayout] [{_data.ComponentInfoWindowState}] [{_data.ComponentInfoWindowWidth:F0}x{_data.ComponentInfoWindowHeight:F0}] [LeftRatio: {_data.ComponentInfoWindowLeftColumnRatio:F3}] [ThumbnailHeight: {_data.ComponentInfoWindowThumbnailRowHeight:F1}]");
                    Logger.Info($"        [ComponentInfoKeyboardHandling] [{ComponentInfoKeyboardHandling}]");
                    Logger.Info($"        [ComponentInfoScrollAction] [{ComponentInfoScrollAction}]");
                    Logger.Info($"        [ContactEmail] [{(string.IsNullOrWhiteSpace(ContactEmail) ? "empty" : "set")}]");
                    Logger.Info($"        [LeftPanelWidth] [{_data.LeftPanelWidth:F1}]");
                    Logger.Info($"        [Region] [{Region}]");
                    Logger.Info($"        [SchematicsLabelsPanelExpanded] [{SchematicsLabelsPanelExpanded}]");
                    Logger.Info($"        [SchematicsLabelBoard] [{SchematicsLabelBoard}]");
                    Logger.Info($"        [SchematicsLabelTechnical] [{SchematicsLabelTechnical}]");
                    Logger.Info($"        [SchematicsLabelFriendly] [{SchematicsLabelFriendly}]");
                    Logger.Info($"        [SchematicsLabelSelectedOnly] [{SchematicsLabelSelectedOnly}]");
                    Logger.Info($"        [WindowPlacement] [{_data.WindowState}] [{_data.WindowWidth:F0}x{_data.WindowHeight:F0}]");
                    Logger.Info($"    Various hardware/board specific settings:");
                    Logger.Info($"        [LastBoardByHardware] [{_data.LastBoardByHardware.Count} entries]");
                    Logger.Info($"        [LastHardware] [{_data.LastHardware}]");
                    Logger.Info($"        [LastSchematicByBoard] [{_data.LastSchematicByBoard.Count} entries]");
                    Logger.Info($"        [SchematicsSplitterRatios] [{_data.SchematicsSplitterRatios.Count} entries]");
                    Logger.Info($"        [SelectedCategoriesByBoard] [{_data.SelectedCategoriesByBoard.Count} entries]");
                    Logger.Info($"        [LastOscilloscopeSeriesByVendor] [{_data.LastOscilloscopeSeriesByVendor.Count} entries]");
                    Logger.Info($"        [LastOscilloscopeVendor] [{_data.LastOscilloscopeVendor}]");
                    Logger.Info($"        [OscilloscopeHost] [{_data.OscilloscopeHost}]");
                    Logger.Info($"        [OscilloscopePort] [{_data.OscilloscopePort}]");
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load settings: [{ex.Message}] - using defaults");
            }
        }

        // ###########################################################################################
        // Serializes current settings and writes them to the JSON file.
        // ###########################################################################################
        private static void Save()
        {
            if (string.IsNullOrEmpty(_settingsFilePath))
                return;

            try
            {
                var json = JsonSerializer.Serialize(_data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to save settings: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Returns the last selected hardware name, or empty when none has been saved.
        // ###########################################################################################
        public static string GetLastHardware()
            => _data.LastHardware ?? string.Empty;

        // ###########################################################################################
        // Persists the last selected hardware name.
        // ###########################################################################################
        public static void SetLastHardware(string hardwareName)
        {
            if (string.IsNullOrWhiteSpace(hardwareName))
                return;

            if (string.Equals(_data.LastHardware, hardwareName, StringComparison.OrdinalIgnoreCase))
                return;

            _data.LastHardware = hardwareName;
            Logger.Info($"Setting changed: [LastHardware] [{hardwareName}]");
            Save();
        }

        // ###########################################################################################
        // Returns the saved last board for a hardware, or null when not found.
        // ###########################################################################################
        public static string? GetLastBoardForHardware(string hardwareName)
        {
            if (string.IsNullOrWhiteSpace(hardwareName))
                return null;

            var match = _data.LastBoardByHardware.FirstOrDefault(kvp =>
                string.Equals(kvp.Key, hardwareName, StringComparison.OrdinalIgnoreCase));

            return string.IsNullOrEmpty(match.Key) ? null : match.Value;
        }

        // ###########################################################################################
        // Persists the last selected board for a specific hardware.
        // ###########################################################################################
        public static void SetLastBoardForHardware(string hardwareName, string boardName)
        {
            if (string.IsNullOrWhiteSpace(hardwareName) || string.IsNullOrWhiteSpace(boardName))
                return;

            var existingKey = _data.LastBoardByHardware.Keys.FirstOrDefault(k =>
                string.Equals(k, hardwareName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(existingKey) &&
                string.Equals(_data.LastBoardByHardware[existingKey], boardName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!string.IsNullOrEmpty(existingKey) &&
                !string.Equals(existingKey, hardwareName, StringComparison.Ordinal))
            {
                _data.LastBoardByHardware.Remove(existingKey);
            }

            _data.LastBoardByHardware[hardwareName] = boardName;
            Logger.Info($"Setting changed: [LastBoardByHardware] [{hardwareName}] [{boardName}]");
            Save();
        }

        // ###########################################################################################
        // Returns the saved last schematic name for a board key, or null when not found.
        // ###########################################################################################
        public static string? GetLastSchematicForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return null;

            return _data.LastSchematicByBoard.TryGetValue(boardKey, out var schematic)
                ? schematic
                : null;
        }

        // ###########################################################################################
        // Persists the last selected schematic name for a board key.
        // ###########################################################################################
        public static void SetLastSchematicForBoard(string boardKey, string schematicName)
        {
            if (string.IsNullOrWhiteSpace(boardKey) || string.IsNullOrWhiteSpace(schematicName))
                return;

            if (_data.LastSchematicByBoard.TryGetValue(boardKey, out var existingSchematic) &&
                string.Equals(existingSchematic, schematicName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            _data.LastSchematicByBoard[boardKey] = schematicName;
            Logger.Info($"Setting changed: [LastSchematicByBoard] [{boardKey}] [{schematicName}]");
            Save();
        }

        public static string OscilloscopeHost
        {
            get => string.IsNullOrWhiteSpace(_data.OscilloscopeHost) ? "192.168.0.100" : _data.OscilloscopeHost;
            set
            {
                if (string.Equals(_data.OscilloscopeHost, value, StringComparison.Ordinal))
                    return;

                _data.OscilloscopeHost = value;
                Logger.Info($"Setting changed: [OscilloscopeHost] [{(string.IsNullOrWhiteSpace(value) ? "empty" : value)}]");
                Save();
            }
        }

        public static int OscilloscopePort
        {
            get => _data.OscilloscopePort;
            set
            {
                if (_data.OscilloscopePort == value)
                    return;

                _data.OscilloscopePort = value;
                Logger.Info($"Setting changed: [OscilloscopePort] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Returns the last selected vendor name, or empty when none has been saved.
        // ###########################################################################################
        public static string GetLastOscilloscopeVendor()
            => _data.LastOscilloscopeVendor ?? string.Empty;

        // ###########################################################################################
        // Persists the last selected vendor name.
        // ###########################################################################################
        public static void SetLastOscilloscopeVendor(string vendorName)
        {
            if (string.IsNullOrWhiteSpace(vendorName))
                return;

            if (string.Equals(_data.LastOscilloscopeVendor, vendorName, StringComparison.OrdinalIgnoreCase))
                return;

            _data.LastOscilloscopeVendor = vendorName;
            Logger.Info($"Setting changed: [LastOscilloscopeVendor] [{vendorName}]");
            Save();
        }

        // ###########################################################################################
        // Returns the saved last series for a vendor, or null when not found.
        // ###########################################################################################
        public static string? GetLastOscilloscopeSeriesForVendor(string vendorName)
        {
            if (string.IsNullOrWhiteSpace(vendorName))
                return null;

            var match = _data.LastOscilloscopeSeriesByVendor.FirstOrDefault(kvp =>
                string.Equals(kvp.Key, vendorName, StringComparison.OrdinalIgnoreCase));

            return string.IsNullOrEmpty(match.Key) ? null : match.Value;
        }

        // ###########################################################################################
        // Persists the last selected series for a specific vendor.
        // ###########################################################################################
        public static void SetLastOscilloscopeSeriesForVendor(string vendorName, string seriesName)
        {
            if (string.IsNullOrWhiteSpace(vendorName) || string.IsNullOrWhiteSpace(seriesName))
                return;

            var existingKey = _data.LastOscilloscopeSeriesByVendor.Keys.FirstOrDefault(k =>
                string.Equals(k, vendorName, StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrEmpty(existingKey) &&
                string.Equals(_data.LastOscilloscopeSeriesByVendor[existingKey], seriesName, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            if (!string.IsNullOrEmpty(existingKey) &&
                !string.Equals(existingKey, vendorName, StringComparison.Ordinal))
            {
                _data.LastOscilloscopeSeriesByVendor.Remove(existingKey);
            }

            _data.LastOscilloscopeSeriesByVendor[vendorName] = seriesName;
            Logger.Info($"Setting changed: [LastOscilloscopeSeriesByVendor] [{vendorName}] [{seriesName}]");
            Save();
        }

        // ###########################################################################################
        // Retrieves the saved thumbnail sorting layout for a board, if present.
        // ###########################################################################################
        public static List<string>? GetSchematicsOrder(string boardKey)
            => _data.SchematicsOrderByBoard.TryGetValue(boardKey, out var order) ? order : null;

        // ###########################################################################################
        // Persists custom dragged thumbnail layout ordering globally.
        // ###########################################################################################
        public static void SetSchematicsOrder(string boardKey, List<string> order)
        {
            if (string.IsNullOrWhiteSpace(boardKey) || order == null)
                return;

            _data.SchematicsOrderByBoard[boardKey] = order;
            Logger.Info($"Setting changed: [SchematicsOrder] [{boardKey}] [{order.Count} sequenced]");
            Save();
        }

    }
}