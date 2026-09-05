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

        [JsonPropertyName("enableNetworkConnectedOscilloscopeTab")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? EnableNetworkConnectedOscilloscopeTab { get; set; }

        [JsonPropertyName("enableMiniproExperimentalMode")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? EnableMiniproExperimentalMode { get; set; }

        [JsonPropertyName("enableMiniproExperimentalDemoMode")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? EnableMiniproExperimentalDemoMode { get; set; }

        [JsonPropertyName("enableWorklog")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? EnableWorklog { get; set; }

        [JsonPropertyName("workbooksScope")] public string WorkbooksScope { get; set; } = "CurrentBoard";

        [JsonPropertyName("worklogCommentsSortNewestFirst")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? WorklogCommentsSortNewestFirst { get; set; }

        [JsonPropertyName("worklogWorkDoneSortNewestFirst")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? WorklogWorkDoneSortNewestFirst { get; set; }

        [JsonPropertyName("worklogShowEntriesChecked")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? WorklogShowEntriesChecked { get; set; }

        [JsonPropertyName("contributorMode")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? ContributorMode { get; set; }

        [JsonPropertyName("leftPanelWidth")] public double LeftPanelWidth { get; set; } = 200.0;
        [JsonPropertyName("workbooksLeftPanelWidth")] public double WorkbooksLeftPanelWidth { get; set; } = 200.0;
        [JsonPropertyName("workbooksEntryListWidth")] public double WorkbooksEntryListWidth { get; set; } = 347.0;

        // Whether the Workbooks tab's workbook-summary strip is expanded to its full breakdown.
        // Collapsed by default: the one-line headline is always visible either way, and starting
        // expanded would push the board pane down for every user to show detail most do not need
        // on every visit.
        [JsonPropertyName("workbooksSummaryExpanded")] public bool WorkbooksSummaryExpanded { get; set; }
        [JsonPropertyName("schematicsSplitterRatios")] public Dictionary<string, double> SchematicsSplitterRatios { get; set; } = new();
        [JsonPropertyName("selectedCategoriesByBoard")] public Dictionary<string, List<string>> SelectedCategoriesByBoard { get; set; } = new();
        [JsonPropertyName("activeWorkbookIdByBoard")] public Dictionary<string, int> ActiveWorkbookIdByBoard { get; set; } = new();
        [JsonPropertyName("lastHardware")] public string LastHardware { get; set; } = string.Empty;
        [JsonPropertyName("lastBoardByHardware")] public Dictionary<string, string> LastBoardByHardware { get; set; } = new();
        [JsonPropertyName("lastSchematicByBoard")] public Dictionary<string, string> LastSchematicByBoard { get; set; } = new();
        [JsonPropertyName("region")] public string Region { get; set; } = "PAL";
        [JsonPropertyName("theme")] public string ThemeVariant { get; set; } = "Light";
        [JsonPropertyName("userThemeColors")] public Dictionary<string, string> UserThemeColors { get; set; } = new();
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
        [JsonPropertyName("componentInfoWindowX")] public int ComponentInfoWindowX { get; set; } = 0;
        [JsonPropertyName("componentInfoWindowY")] public int ComponentInfoWindowY { get; set; } = 0;
        [JsonPropertyName("hasWorklogEntryWindowLayout")] public bool HasWorklogEntryWindowLayout { get; set; } = false;
        [JsonPropertyName("worklogEntryWindowState")] public string WorklogEntryWindowState { get; set; } = "Normal";
        [JsonPropertyName("worklogEntryWindowWidth")] public double WorklogEntryWindowWidth { get; set; } = 1200.0;
        [JsonPropertyName("worklogEntryWindowHeight")] public double WorklogEntryWindowHeight { get; set; } = 800.0;
        [JsonPropertyName("worklogEntryWindowX")] public int WorklogEntryWindowX { get; set; } = 0;
        [JsonPropertyName("worklogEntryWindowY")] public int WorklogEntryWindowY { get; set; } = 0;

        // The top-left of the SCREEN the window was last on, independent of WorklogEntryWindowX/Y
        // (the Normal-state restore position). Needed only for the Maximized case: restoring a
        // maximized window by setting WindowState alone maximizes it on whatever monitor the OS
        // last remembers, which is wrong once the window was dragged to - and maximized on -
        // another monitor without ever being un-maximized there (WorklogEntryWindowX/Y then never
        // updates, since position is only tracked in Normal state). Restoring instead moves the
        // window onto this screen first, then maximizes - the same two-step Main.axaml.cs's own
        // window placement uses for the same reason.
        [JsonPropertyName("worklogEntryWindowScreenX")] public int WorklogEntryWindowScreenX { get; set; } = 0;
        [JsonPropertyName("worklogEntryWindowScreenY")] public int WorklogEntryWindowScreenY { get; set; } = 0;

        [JsonPropertyName("worklogEntryWindowLeftColumnRatio")] public double WorklogEntryWindowLeftColumnRatio { get; set; } = 0.6;
        [JsonPropertyName("componentInfoScrollAction")] public string ComponentInfoScrollAction { get; set; } = "Image change";
        [JsonPropertyName("schematicsLabelBoard")] public bool SchematicsLabelBoard { get; set; } = false;
        [JsonPropertyName("schematicsLabelTechnical")] public bool SchematicsLabelTechnical { get; set; } = false;
        [JsonPropertyName("schematicsLabelFriendly")] public bool SchematicsLabelFriendly { get; set; } = false;
        [JsonPropertyName("schematicsLabelSelectedOnly")] public bool SchematicsLabelSelectedOnly { get; set; } = false;
        [JsonPropertyName("schematicsLabelsPanelExpanded")] public bool SchematicsLabelsPanelExpanded { get; set; } = true;
        [JsonPropertyName("schematicsBoardSettingsPanelExpanded")] public bool SchematicsBoardSettingsPanelExpanded { get; set; } = true;
        [JsonPropertyName("schematicsGlobalSettingsPanelExpanded")] public bool SchematicsGlobalSettingsPanelExpanded { get; set; } = true;
        [JsonPropertyName("schematicsNetConnectionsPanelExpanded")] public bool SchematicsNetConnectionsPanelExpanded { get; set; } = true;
        [JsonPropertyName("schematicsShowTracesOnComponentSelect")] public bool SchematicsShowTracesOnComponentSelect { get; set; } = true;
        [JsonPropertyName("blinkSelected")] public bool BlinkSelected { get; set; } = false;
        [JsonPropertyName("schematicsOrderByBoard")] public Dictionary<string, List<string>> SchematicsOrderByBoard { get; set; } = new();
        [JsonPropertyName("contactEmail")] public string ContactEmail { get; set; } = string.Empty;

        [JsonPropertyName("oscilloscopeVendor")] public string LastOscilloscopeVendor { get; set; } = string.Empty;
        [JsonPropertyName("oscilloscopeSeriesByVendor")] public Dictionary<string, string> LastOscilloscopeSeriesByVendor { get; set; } = new();
        [JsonPropertyName("oscilloscopeHost")] public string OscilloscopeHost { get; set; } = "192.168.0.100";
        [JsonPropertyName("oscilloscopePort")] public int OscilloscopePort { get; set; } = 5025;
        [JsonPropertyName("oscilloscopeAutoConnect")] public bool OscilloscopeAutoConnect { get; set; } = false;
        [JsonPropertyName("componentInfoKeyboardHandling")] public string ComponentInfoKeyboardHandling { get; set; } = "Control oscilloscope";
        [JsonPropertyName("oscilloscopeImageFolder")] public string OscilloscopeImageFolder { get; set; } = string.Empty;
        [JsonPropertyName("oscilloscopeSyncEnabled")] public bool? ComponentInfoOscilloscopeSyncEnabled { get; set; }
        [JsonPropertyName("schematicsHoverHighlightsTracesByBoard")] public Dictionary<string, bool> SchematicsHoverHighlightsTracesByBoard { get; set; } = new();
        [JsonPropertyName("contributorModeByBoard")] public Dictionary<string, bool> ContributorModeByBoard { get; set; } = new();
        [JsonPropertyName("interactiveCadTraceHoverMode")] public string InteractiveCadTraceHoverMode { get; set; } = "Always";
        [JsonPropertyName("schematicsMarkPin1OnSelectedComponentByBoard")] public Dictionary<string, bool> SchematicsMarkPin1OnSelectedComponentByBoard { get; set; } = new();
        [JsonPropertyName("schematicsShowTracesOnSelectedComponent")] public bool SchematicsShowTracesOnSelectedComponent { get; set; } = true;
        [JsonPropertyName("schematicsShowTracesOnSelectedComponentByBoard")] public Dictionary<string, bool> SchematicsShowTracesOnSelectedComponentByBoard { get; set; } = new();
        [JsonPropertyName("schematicsImportantSignalsPanelExpanded")] public bool SchematicsImportantSignalsPanelExpanded { get; set; } = true;
        [JsonPropertyName("schematicsShowOppositeSideTraces")] public bool SchematicsShowOppositeSideTraces { get; set; } = true;
        [JsonPropertyName("schematicsShowZones")] public bool SchematicsShowZones { get; set; } = true;
        [JsonPropertyName("downloadDataFromTestSource")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? DownloadDataFromTestSource { get; set; }

        [JsonPropertyName("allowDeletionOfOrphanAndNonUsedFiles")]
        [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
        public bool? AllowDeletionOfOrphanAndNonUsedFiles { get; set; }

    }

    // ###########################################################################################
    // Loads and saves user preferences to a JSON file in the same folder as the log file.
    // Call Load() once at startup before any settings are read.
    // ###########################################################################################
    public static class UserSettings
    {
        private static UserSettingsData _data = new();
        private static string _settingsFilePath = string.Empty;

        private static readonly Dictionary<string, string> UserThemeDefaultColors = new(StringComparer.Ordinal)
        {
            { "Bg", "#E0E4D9" },
            { "Fg", "Black" },
            { "Border", "Black" },
            { "Form_Bg", "White" },
            { "Form_Fg", "Black" },
            { "Form_Border", "Gray" },
            { "Form_Selected_Bg", "LightBlue" },
            { "Form_Selected_Fg", "White" },
            { "Form_Selected_Hover_Bg", "SkyBlue" },
            { "Splitter", "Gray" },
            { "Button_Bg", "LightGray" },
            { "Button_Fg", "Black" },
            { "Button_Border", "DarkGray" },
            { "Button_Ok_Bg", "#4F8A5B" },
            { "Button_Ok_Fg", "White" },
            { "Button_Ok_Border", "#3E6E48" },
            { "Button_Cancel_Bg", "IndianRed" },
            { "Button_Cancel_Fg", "White" },
            { "Button_Cancel_Border", "#B85C5C" },
            { "Button_Highlighted_Bg", "IndianRed" },
            { "Button_Highlighted_Fg", "White" },
            { "Button_Highlighted_Border", "White" },
            { "Panel_Border", "Gray" },
            { "Link_Fg", "Blue" },
            { "LinkHover_Fg", "Black" },
            { "Table_Header_Bg", "LightGray" },
            { "Table_Header_Fg", "Black" },
            { "Table_Bg", "White" },
            { "Table_Fg", "Black" },
            { "Table_BorderRowLine", "#EEEEEE" },
            { "Table_RowHover_Bg", "#E1F0FA" },
            { "Text_Success_Fg", "#547131" },
            { "Text_Fail_Fg", "IndianRed" },
            { "Main_TabUnderline_Selected", "IndianRed" },
            { "Main_TabUnderline_Hover", "#ecc3c3" },
            { "Main_Banner_Bg", "#4C1515" },
            { "Main_Banner_Fg", "White" },
            { "Main_BannerButtonsHover_Bg", "#4F3528" },
            { "Main_BannerButtonsHover_Fg", "#DDE0DA" },
            { "Schematics_Name_Bg", "Khaki" },
            { "Schematics_Name_Fg", "Black" },
            { "Schematics_Name_Border", "Black" },
            { "Schematics_Region_PAL_Bg", "IndianRed" },
            { "Schematics_Region_PAL_Fg", "White" },
            { "Schematics_Region_PAL_Border", "Black" },
            { "Schematics_Region_NTSC_Bg", "#CBD6E5" },
            { "Schematics_Region_NTSC_Fg", "Black" },
            { "Schematics_Region_NTSC_Border", "Black" },
            { "Schematics_ComponentHover_Bg", "Khaki" },
            { "Schematics_ComponentHover_Fg", "Black" },
            { "Schematics_ComponentHover_Border", "Black" },
            { "Schematics_Panels_Bg", "#F2F2F2" },
            { "Schematics_Panels_Fg", "Black" },
            { "Schematics_Panels_Border", "Gray" },
            { "Schematics_ComponentLabel_Bg", "#222113" },
            { "Schematics_ComponentLabel_Fg", "#ECEEEF" },
            { "Schematics_ComponentLabel_Border", "Gray" },
            { "Schematics_Panel_TracesVisible_Trace_Border", "Gray" },
            { "Schematics_FirstPin", "Orange" },
            { "Schematics_TracePalette_Color1", "#F40000" },
            { "Schematics_TracePalette_Color2", "#1564F4" },
            { "Schematics_TracePalette_Color3", "#0BC419" },
            { "Schematics_TracePalette_Color4", "#E6E600" },
            { "Thumbnail_Border", "Black" },
            { "Thumbnail_BorderSelected", "IndianRed" },
            { "Thumbnail_Bg", "Khaki" },
            { "Thumbnail_Fg", "Black" },
            { "Thumbnail_Selection_Bg", "#4682B4" },
            { "Thumbnail_Selection_Fg", "White" },
            { "ComponentPopup_Name_Bg", "Khaki" },
            { "ComponentPopup_Name_Fg", "Black" },
            { "ComponentPopup_Name_Border", "Black" },
            { "ComponentPopup_Pin_Bg", "Khaki" },
            { "ComponentPopup_Pin_Fg", "Black" },
            { "ComponentPopup_Pin_Border", "Black" },
            { "ComponentPopup_OscilloscopeExpected_Bg", "#B0D8EC" },
            { "ComponentPopup_OscilloscopeExpected_Fg", "Black" },
            { "ComponentPopup_OscilloscopeExpected_Border", "Black" },
            { "ComponentPopup_OscilloscopeSyncData_Bg", "LightGray" },
            { "ComponentPopup_OscilloscopeSyncData_Fg", "Black" },
            { "ComponentPopup_OscilloscopeSyncData_Border", "Black" },
            { "ComponentPopup_ImageXofY_Bg", "#BB000000" },
            { "ComponentPopup_ImageXofY_Border", "Black" },
            { "ComponentPopup_ImageXofY_Fg", "White" },
            { "ComponentPopup_KeyboardTyping_Bg", "#778461" },
            { "ComponentPopup_KeyboardTyping_Fg", "White" },
            { "ComponentPopup_KeyboardTyping_Border", "Black" },
            { "ComponentPopup_ImageTransparent_Bg", "White" },
            { "ComponentPopup_BulletList_Fg", "Red" },
            { "Feedback_TrashIcon_Fg", "IndianRed" }
        };

        public static bool AllowDeletionOfOrphanAndNonUsedFiles
        {
            get => _data.AllowDeletionOfOrphanAndNonUsedFiles ?? false;
            set
            {
                bool currentValue = _data.AllowDeletionOfOrphanAndNonUsedFiles ?? false;
                if (currentValue == value)
                    return;

                _data.AllowDeletionOfOrphanAndNonUsedFiles = value;
                Logger.Info($"Setting changed: [AllowDeletionOfOrphanAndNonUsedFiles] [{value}]");
                Save();
            }
        }

        public static bool DownloadDataFromTestSource
        {
            get => _data.DownloadDataFromTestSource ?? false;
            set
            {
                bool currentValue = _data.DownloadDataFromTestSource ?? false;
                if (currentValue == value)
                    return;

                _data.DownloadDataFromTestSource = value;
                Logger.Info($"Setting changed: [DownloadDataFromTestSource] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Returns whether opposite-side PCB traces should be rendered while viewing a PCB replica.
        // ###########################################################################################
        public static bool SchematicsShowOppositeSideTraces
        {
            get => _data.SchematicsShowOppositeSideTraces;
            set
            {
                if (_data.SchematicsShowOppositeSideTraces == value)
                    return;

                _data.SchematicsShowOppositeSideTraces = value;
                Logger.Info($"Setting changed: [SchematicsShowOppositeSideTraces] [{value}]");
                Save();
            }
        }

        public static bool SchematicsImportantSignalsPanelExpanded
        {
            get => _data.SchematicsImportantSignalsPanelExpanded;
            set
            {
                _data.SchematicsImportantSignalsPanelExpanded = value;
                Save();
            }
        }

        // ###########################################################################################
        // Returns whether KiCad copper zones are shown in the schematics overlay.
        // Defaults to true unless explicitly disabled by the user.
        // ###########################################################################################
        public static bool SchematicsShowZones
        {
            get => _data.SchematicsShowZones;
            set
            {
                if (_data.SchematicsShowZones == value)
                    return;

                _data.SchematicsShowZones = value;
                Logger.Info($"Setting changed: [SchematicsShowZones] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Returns whether selected-component trace preview is enabled globally.
        // ###########################################################################################
        public static bool SchematicsShowTracesOnComponentSelect
        {
            get => _data.SchematicsShowTracesOnComponentSelect;
            set
            {
                if (_data.SchematicsShowTracesOnComponentSelect == value)
                    return;

                _data.SchematicsShowTracesOnComponentSelect = value;
                Logger.Info($"Setting changed: [SchematicsShowTracesOnComponentSelect] [{value}]");
                Save();
            }
        }

        public static bool SchematicsNetConnectionsPanelExpanded
        {
            get => _data.SchematicsNetConnectionsPanelExpanded;
            set { _data.SchematicsNetConnectionsPanelExpanded = value; Save(); }
        }

        // ###########################################################################################
        // Returns true when global contributor mode is enabled.
        // ###########################################################################################
        private static bool IsContributorModeEnabled()
        {
            return _data.ContributorMode ?? false;
        }

        // ###########################################################################################
        // Migrates the removed legacy configuration checkboxes into contributor mode once.
        // ###########################################################################################
        private static void MigrateLegacyContributorModeSettings()
        {
            if (_data.ContributorMode.HasValue)
            {
                return;
            }

            if (_data.ValidateDataOnLaunch == true &&
                _data.DebugLogging == true)
            {
                _data.ContributorMode = true;
            }
        }

        // ###########################################################################################
        // Keeps the removed legacy settings synchronized with contributor mode.
        // ###########################################################################################
        private static void SyncContributorModeLinkedSettings()
        {
            bool isEnabled = IsContributorModeEnabled();

            _data.ValidateDataOnLaunch = isEnabled;
            _data.DebugLogging = isEnabled;
            Logger.IsDebugEnabled = isEnabled;
        }

        private static readonly HashSet<string> UserThemeColorResourceKeys = new(StringComparer.Ordinal)
        {
            "Schematics_TracePalette_Color1",
            "Schematics_TracePalette_Color2",
            "Schematics_TracePalette_Color3",
            "Schematics_TracePalette_Color4"
        };

        public static event Action? InteractiveCadTraceHoverModeChanged;

        public static string InteractiveCadTraceHoverMode
        {
            get => string.Equals(_data.InteractiveCadTraceHoverMode, "Disabled", StringComparison.OrdinalIgnoreCase)
                ? "Disabled"
                : string.Equals(_data.InteractiveCadTraceHoverMode, "HoldShift", StringComparison.OrdinalIgnoreCase)
                    ? "HoldShift"
                    : "Always";
            set
            {
                string normalized = string.Equals(value, "Disabled", StringComparison.OrdinalIgnoreCase)
                    ? "Disabled"
                    : string.Equals(value, "HoldShift", StringComparison.OrdinalIgnoreCase)
                        ? "HoldShift"
                        : "Always";

                if (string.Equals(_data.InteractiveCadTraceHoverMode, normalized, StringComparison.Ordinal))
                    return;

                _data.InteractiveCadTraceHoverMode = normalized;
                Logger.Info($"Setting changed: [InteractiveCadTraceHoverMode] [{normalized}]");
                Save();
                InteractiveCadTraceHoverModeChanged?.Invoke();
            }
        }

        public static bool ComponentInfoOscilloscopeSyncEnabled
        {
            get => _data.ComponentInfoOscilloscopeSyncEnabled ?? true;
            set
            {
                _data.ComponentInfoOscilloscopeSyncEnabled = value;
                Logger.Info($"Setting changed: [OscilloscopeSyncEnabled] [{value}]");
                Save();
            }
        }

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

        public static bool SchematicsLabelsPanelExpanded
        {
            get => _data.SchematicsLabelsPanelExpanded;
            set { _data.SchematicsLabelsPanelExpanded = value; Save(); }
        }

        public static bool SchematicsBoardSettingsPanelExpanded
        {
            get => _data.SchematicsBoardSettingsPanelExpanded;
            set { _data.SchematicsBoardSettingsPanelExpanded = value; Save(); }
        }

        public static bool SchematicsGlobalSettingsPanelExpanded
        {
            get => _data.SchematicsGlobalSettingsPanelExpanded;
            set { _data.SchematicsGlobalSettingsPanelExpanded = value; Save(); }
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

        public static event Action<bool>? CheckDataOnLaunchChanged;


        public static bool CheckDataOnLaunch
        {
            get => _data.CheckDataOnLaunch ?? true;
            set
            {
                bool currentValue = _data.CheckDataOnLaunch ?? true;
                if (currentValue == value)
                    return;

                _data.CheckDataOnLaunch = value;
                Logger.Info($"Setting changed: [CheckDataOnLaunch] [{value}]");
                Save();
                CheckDataOnLaunchChanged?.Invoke(value);
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
            get => IsContributorModeEnabled();
            set
            {
                _data.ContributorMode = value;
                SyncContributorModeLinkedSettings();
                Logger.Info($"Setting changed: [ValidateDataOnLaunch] [{value}]");
                Save();
            }
        }

        public static bool DebugLogging
        {
            get => IsContributorModeEnabled();
            set
            {
                _data.ContributorMode = value;
                SyncContributorModeLinkedSettings();
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

        public static bool EnableNetworkConnectedOscilloscopeTab
        {
            get => _data.EnableNetworkConnectedOscilloscopeTab ?? true; // Default is true
            set
            {
                _data.EnableNetworkConnectedOscilloscopeTab = value;
                Logger.Info($"Setting changed: [EnableNetworkConnectedOscilloscopeTab] [{value}]");
                Save();
            }
        }

        public static bool EnableMiniproExperimentalMode
        {
            get => _data.EnableMiniproExperimentalMode ?? true; // Default is true
            set
            {
                _data.EnableMiniproExperimentalMode = value;
                Logger.Info($"Setting changed: [EnableMiniproExperimentalMode] [{value}]");
                Save();
            }
        }

        public static bool EnableMiniproExperimentalDemoMode
        {
            get => _data.EnableMiniproExperimentalDemoMode ?? false; // Default is false
            set
            {
                _data.EnableMiniproExperimentalDemoMode = value;
                Logger.Info($"Setting changed: [EnableMiniproExperimentalDemoMode] [{value}]");
                Save();
            }
        }

        public static bool EnableWorklog
        {
            get => _data.EnableWorklog ?? true; // Default is true
            set
            {
                _data.EnableWorklog = value;
                Logger.Info($"Setting changed: [EnableWorklog] [{value}]");
                Save();
            }
        }

        public static event Action? WorkbooksScopeChanged;

        // ###########################################################################################
        // Which workbooks the Workbooks tab's left-hand list shows: every workbook on every board
        // ("AllBoards"), or only the ones for the currently selected board ("CurrentBoard", the
        // original and default behaviour). Set from the Configuration tab's radio group below
        // "Enable Worklog".
        // ###########################################################################################
        public static string WorkbooksScope
        {
            get => string.Equals(_data.WorkbooksScope, "AllBoards", StringComparison.OrdinalIgnoreCase)
                ? "AllBoards"
                : "CurrentBoard";
            set
            {
                string normalized = string.Equals(value, "AllBoards", StringComparison.OrdinalIgnoreCase)
                    ? "AllBoards"
                    : "CurrentBoard";

                if (string.Equals(_data.WorkbooksScope, normalized, StringComparison.Ordinal))
                    return;

                _data.WorkbooksScope = normalized;
                Logger.Info($"Setting changed: [WorkbooksScope] [{normalized}]");
                Save();
                WorkbooksScopeChanged?.Invoke();
            }
        }

        // ###########################################################################################
        // Sort order for the worklog entry editor's Comments list. Defaults to newest-first, matching
        // the editor's in-memory default before this was persisted.
        // ###########################################################################################
        public static bool WorklogCommentsSortNewestFirst
        {
            get => _data.WorklogCommentsSortNewestFirst ?? true;
            set
            {
                if ((_data.WorklogCommentsSortNewestFirst ?? true) == value)
                    return;

                _data.WorklogCommentsSortNewestFirst = value;
                Logger.Info($"Setting changed: [WorklogCommentsSortNewestFirst] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Sort order for the worklog entry editor's Work done list. Defaults to newest-first, matching
        // the editor's in-memory default before this was persisted.
        // ###########################################################################################
        public static bool WorklogWorkDoneSortNewestFirst
        {
            get => _data.WorklogWorkDoneSortNewestFirst ?? true;
            set
            {
                if ((_data.WorklogWorkDoneSortNewestFirst ?? true) == value)
                    return;

                _data.WorklogWorkDoneSortNewestFirst = value;
                Logger.Info($"Setting changed: [WorklogWorkDoneSortNewestFirst] [{value}]");
                Save();
            }
        }

        // ###########################################################################################
        // Whether the top-bar's "Show worklogs" checkbox should default to checked. Defaults to
        // true so the entries list overlay shows automatically. This is the checkbox's own
        // preferred value, not whether it is currently showing entries for the board on screen -
        // RefreshWorklogBar forces it unchecked whenever there is no workbook, or a different one
        // than it was showing, to avoid displaying entries for a workbook the user is no longer
        // looking at.
        // ###########################################################################################
        public static bool WorklogShowEntriesChecked
        {
            get => _data.WorklogShowEntriesChecked ?? true;
            set
            {
                if ((_data.WorklogShowEntriesChecked ?? true) == value)
                    return;

                _data.WorklogShowEntriesChecked = value;
                Logger.Info($"Setting changed: [WorklogShowEntriesChecked] [{value}]");
                Save();
            }
        }

        public static bool ContributorMode
        {
            get => _data.ContributorMode ?? false;
            set
            {
                _data.ContributorMode = value;
                SyncContributorModeLinkedSettings();
                Logger.Info($"Setting changed: [ContributorMode] [{value}]");
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

        public static double WorkbooksLeftPanelWidth
        {
            get => _data.WorkbooksLeftPanelWidth;
            set
            {
                _data.WorkbooksLeftPanelWidth = value;
                Logger.Info($"Setting changed: [WorkbooksLeftPanelWidth] [{value:F1}]");
                Save();
            }
        }

        public static double WorkbooksEntryListWidth
        {
            get => _data.WorkbooksEntryListWidth;
            set
            {
                _data.WorkbooksEntryListWidth = value;
                Logger.Info($"Setting changed: [WorkbooksEntryListWidth] [{value:F1}]");
                Save();
            }
        }

        public static bool WorkbooksSummaryExpanded
        {
            get => _data.WorkbooksSummaryExpanded;
            set
            {
                _data.WorkbooksSummaryExpanded = value;
                Logger.Info($"Setting changed: [WorkbooksSummaryExpanded] [{value}]");
                Save();
            }
        }

        public static string ThemeVariant
        {
            get => string.Equals(_data.ThemeVariant, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark"
                : string.Equals(_data.ThemeVariant, "UserPreference", StringComparison.OrdinalIgnoreCase) ? "UserPreference"
                : "Light";
            set
            {
                var normalized = string.Equals(value, "Dark", StringComparison.OrdinalIgnoreCase) ? "Dark"
                    : string.Equals(value, "UserPreference", StringComparison.OrdinalIgnoreCase) ? "UserPreference"
                    : "Light";

                if (string.Equals(_data.ThemeVariant, normalized, StringComparison.Ordinal))
                    return;

                _data.ThemeVariant = normalized;
                Logger.Info($"Setting changed: [Theme] [{normalized}]");
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

        // ###########################################################################################
        // Returns whether selected-component trace preview is enabled globally.
        // ###########################################################################################
        public static bool SchematicsShowTracesOnSelectedComponent
        {
            get => _data.SchematicsShowTracesOnSelectedComponent;
            set
            {
                if (_data.SchematicsShowTracesOnSelectedComponent == value)
                    return;

                _data.SchematicsShowTracesOnSelectedComponent = value;
                Logger.Info($"Setting changed: [SchematicsShowTracesOnSelectedComponent] [{value}]");
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
        public static int ComponentInfoWindowX => _data.ComponentInfoWindowX;
        public static int ComponentInfoWindowY => _data.ComponentInfoWindowY;

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
        public static void SaveComponentInfoWindowLayout(string state, double width, double height, double leftColumnRatio, double thumbnailRowHeight, int x, int y)
        {
            _data.HasComponentInfoWindowLayout = true;
            _data.ComponentInfoWindowState = state;
            _data.ComponentInfoWindowWidth = width;
            _data.ComponentInfoWindowHeight = height;
            _data.ComponentInfoWindowLeftColumnRatio = leftColumnRatio;
            _data.ComponentInfoWindowThumbnailRowHeight = thumbnailRowHeight;
            _data.ComponentInfoWindowX = x;
            _data.ComponentInfoWindowY = y;
            Logger.Info($"Setting changed: [ComponentInfoWindowLayout] [{state}] [{width:F0}x{height:F0}] [LeftRatio: {leftColumnRatio:F3}] [ThumbnailHeight: {thumbnailRowHeight:F1}] [Position: {x},{y}]");
            Save();
        }

        // Worklog entry window layout - read-only; written atomically via SaveWorklogEntryWindowLayout.
        // The splitter is stored as the left column's SHARE of the two content columns rather than a
        // pixel width, so reopening at a different window size keeps the same proportions instead of
        // leaving one side squeezed. 0.6 is the 3*:2* the XAML declares.
        public static bool HasWorklogEntryWindowLayout => _data.HasWorklogEntryWindowLayout;
        public static string WorklogEntryWindowState => _data.WorklogEntryWindowState;
        public static double WorklogEntryWindowWidth => _data.WorklogEntryWindowWidth;
        public static double WorklogEntryWindowHeight => _data.WorklogEntryWindowHeight;
        public static int WorklogEntryWindowX => _data.WorklogEntryWindowX;
        public static int WorklogEntryWindowY => _data.WorklogEntryWindowY;
        public static int WorklogEntryWindowScreenX => _data.WorklogEntryWindowScreenX;
        public static int WorklogEntryWindowScreenY => _data.WorklogEntryWindowScreenY;
        public static double WorklogEntryWindowLeftColumnRatio => _data.WorklogEntryWindowLeftColumnRatio;

        // ###########################################################################################
        // Saves the worklog entry editor's window placement atomically in a single disk write.
        //
        // Width/height/x/y are always the NORMAL (non-maximized) values - see the caller. Storing a
        // maximized window's bounds here would mean un-maximizing it later restored it to full
        // screen size, with no memory of the size the user actually chose.
        //
        // screenX/screenY are the top-left of whichever screen the window is CURRENTLY on, tracked
        // independently of x/y - see WorklogEntryWindowScreenX's own comment for why a maximized
        // window needs this separately from its Normal-state restore position.
        // ###########################################################################################
        public static void SaveWorklogEntryWindowLayout(
            string state, double width, double height, int x, int y, int screenX, int screenY, double leftColumnRatio)
        {
            _data.HasWorklogEntryWindowLayout = true;
            _data.WorklogEntryWindowState = state;
            _data.WorklogEntryWindowWidth = width;
            _data.WorklogEntryWindowHeight = height;
            _data.WorklogEntryWindowX = x;
            _data.WorklogEntryWindowY = y;
            _data.WorklogEntryWindowScreenX = screenX;
            _data.WorklogEntryWindowScreenY = screenY;
            _data.WorklogEntryWindowLeftColumnRatio = leftColumnRatio;
            Logger.Info($"Setting changed: [WorklogEntryWindowLayout] [{state}] [{width:F0}x{height:F0}] [Position: {x},{y}] [Screen: {screenX},{screenY}] [LeftRatio: {leftColumnRatio:F3}]");
            Save();
        }

        // ###########################################################################################
        // Returns true when the given board already has a persisted schematics splitter ratio.
        // ###########################################################################################
        public static bool HasSchematicsSplitterRatio(string boardKey)
            => !string.IsNullOrWhiteSpace(boardKey) &&
               _data.SchematicsSplitterRatios.ContainsKey(boardKey);

        // ###########################################################################################
        // Returns the saved schematics splitter ratio for the given board key.
        // Defaults to 0.5 (equal split) for legacy callers when no saved value exists.
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
        // The workbook the user explicitly "activated" from the Workbooks tab for the given board -
        // the one the worklog bar shows, "Show worklogs" draws, and "Add worklog" writes new entries
        // into, in place of the board's newest workbook.
        //
        // Returns null when nothing has been activated (the common case: most boards are worked on
        // one workbook at a time, and the newest-workbook fallback in Main.RefreshWorklogBar already
        // covers that). The caller is responsible for checking the id still names a real workbook on
        // this board - a workbook can be deleted from disk by hand, and this dictionary is not told.
        // ###########################################################################################
        public static int? GetActiveWorkbookId(string boardKey)
            => _data.ActiveWorkbookIdByBoard.TryGetValue(boardKey, out var id) ? id : null;

        // ###########################################################################################
        // Persists the activated workbook id for the given board key.
        // ###########################################################################################
        public static void SetActiveWorkbookId(string boardKey, int workbookId)
        {
            _data.ActiveWorkbookIdByBoard[boardKey] = workbookId;
            Logger.Info($"Setting changed: [ActiveWorkbookId] [{boardKey}] [#{workbookId}]");
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
        // Returns the persisted user-defined theme colors loaded from configuration.
        // ###########################################################################################
        public static IReadOnlyDictionary<string, string> GetUserThemeColors()
            => _data.UserThemeColors;

        // ###########################################################################################
        // Returns whether a resource key should be applied as a raw Color instead of a brush.
        // ###########################################################################################
        public static bool IsUserThemeColorResourceKey(string resourceKey)
            => UserThemeColorResourceKeys.Contains(resourceKey);

        // ###########################################################################################
        // Ensures all JSON-backed user theme color entries exist, seeding missing values from Light.
        // ###########################################################################################
        private static bool EnsureUserThemeColors()
        {
            _data.UserThemeColors ??= new Dictionary<string, string>();

            var changed = false;

            foreach (var entry in UserThemeDefaultColors)
            {
                if (!_data.UserThemeColors.TryGetValue(entry.Key, out var value) || string.IsNullOrWhiteSpace(value))
                {
                    _data.UserThemeColors[entry.Key] = entry.Value;
                    changed = true;
                }
            }

            return changed;
        }

        // ###########################################################################################
        // Resolves the settings file path in the user's AppData folder and loads persisted values.
        // Falls back to defaults silently on any failure.
        // ###########################################################################################
        public static void Load()
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var directory = Path.Combine(appData, AppConfig.AppFolderName);
                Directory.CreateDirectory(directory);
                LoadFrom(Path.Combine(directory, AppConfig.SettingsFileName));
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load settings: [{ex.Message}] - using defaults");
            }
        }

        // ###########################################################################################
        // Loads persisted values from an explicit settings file and makes that file the save target.
        // Load() resolves the real AppData location and calls this; splitting the two keeps the
        // "where do settings live" decision out of the load logic, and lets the test suite point at
        // a temporary file instead of the user's real settings.
        // Falls back to defaults silently on any failure.
        // ###########################################################################################
        internal static void LoadFrom(string settingsFilePath)
        {
            try
            {
                _settingsFilePath = settingsFilePath;

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
                Logger.Info($"    CPU architecture is [{archDescription}]");
                Logger.Info($"    Using self-contained .NET Runtime [{RuntimeInformation.FrameworkDescription}]");

                // Now evaluate settings
                if (!File.Exists(_settingsFilePath))
                {
                    var createdUserThemeDefaults = EnsureUserThemeColors();

                    Logger.Info("Configuration file not found - using defaults");

                    if (createdUserThemeDefaults)
                    {
                        Save();
                        Logger.Info($"Created [UserThemeColors] defaults [{_data.UserThemeColors.Count} entries]");
                    }

                    return;
                }

                var json = File.ReadAllText(_settingsFilePath);
                var loaded = JsonSerializer.Deserialize<UserSettingsData>(json);
                if (loaded != null)
                {
                    _data = loaded;

                    MigrateLegacyContributorModeSettings();

                    var addedUserThemeDefaults = EnsureUserThemeColors();
                    SyncContributorModeLinkedSettings();

                    Logger.Info("Settings loaded:");
                    Logger.Info($"    Configuration:");
                    Logger.Info($"        [Theme] [{ThemeVariant}]");
                    Logger.Info($"        [OpenMultiplePopups] [{MultipleInstancesForComponentPopup}]");
                    Logger.Info($"        [EnableNetworkConnectedOscilloscopeTab] [{EnableNetworkConnectedOscilloscopeTab}]");
                    Logger.Info($"        [EnableMiniproExperimentalMode] [{EnableMiniproExperimentalMode}]");
                    Logger.Info($"        [EnableWorklog] [{EnableWorklog}]");
                    Logger.Info($"        [WorkbooksScope] [{WorkbooksScope}]");
                    Logger.Info($"        [WorklogCommentsSortNewestFirst] [{WorklogCommentsSortNewestFirst}]");
                    Logger.Info($"        [WorklogWorkDoneSortNewestFirst] [{WorklogWorkDoneSortNewestFirst}]");
                    Logger.Info($"        [WorklogShowEntriesChecked] [{WorklogShowEntriesChecked}]");
                    Logger.Info($"        [CheckDataOnLaunch] [{CheckDataOnLaunch}]");
                    Logger.Info($"        [DownloadDataFromTestSource] [{DownloadDataFromTestSource}]");
                    Logger.Info($"        [AllowDeletionOfOrphanAndNonUsedFiles] [{AllowDeletionOfOrphanAndNonUsedFiles}]");
                    Logger.Info($"        [CheckVersionOnLaunch] [{CheckVersionOnLaunch}]");
                    Logger.Info($"        [AllowBetaNotification] [{ShowDevelopmentVersionNotification}]");
                    Logger.Info($"        [ValidateDataOnLaunch] [{ValidateDataOnLaunch}]");
                    Logger.Info($"        [DebugLogging] [{DebugLogging}]");
                    Logger.Info($"    Various other settings:");
                    Logger.Info($"        [BlinkSelected] [{BlinkSelected}]");
                    Logger.Info($"        [ComponentInfoWindowLayout] [{_data.ComponentInfoWindowState}] [{_data.ComponentInfoWindowWidth:F0}x{_data.ComponentInfoWindowHeight:F0}] [LeftRatio: {_data.ComponentInfoWindowLeftColumnRatio:F3}] [ThumbnailHeight: {_data.ComponentInfoWindowThumbnailRowHeight:F1}] [Position: {_data.ComponentInfoWindowX},{_data.ComponentInfoWindowY}]");
                    Logger.Info($"        [ComponentInfoKeyboardHandling] [{ComponentInfoKeyboardHandling}]");
                    Logger.Info($"        [ComponentInfoScrollAction] [{ComponentInfoScrollAction}]");
                    Logger.Info($"        [ContactEmail] [{(string.IsNullOrWhiteSpace(ContactEmail) ? "empty" : "set")}]");
                    Logger.Info($"        [ContributorMode] [{ContributorMode}]");
                    Logger.Info($"        [InteractiveCadTraceHoverMode] [{InteractiveCadTraceHoverMode}]");
                    Logger.Info($"        [LeftPanelWidth] [{_data.LeftPanelWidth:F1}]");
                    Logger.Info($"        [WorkbooksLeftPanelWidth] [{_data.WorkbooksLeftPanelWidth:F1}]");
                    Logger.Info($"        [WorkbooksEntryListWidth] [{_data.WorkbooksEntryListWidth:F1}]");
                    Logger.Info($"        [WorkbooksSummaryExpanded] [{_data.WorkbooksSummaryExpanded}]");
                    Logger.Info($"        [Region] [{Region}]");
                    Logger.Info($"        [SchematicsLabelsPanelExpanded] [{SchematicsLabelsPanelExpanded}]");
                    Logger.Info($"        [SchematicsBoardSettingsPanelExpanded] [{SchematicsBoardSettingsPanelExpanded}]");
                    Logger.Info($"        [SchematicsGlobalSettingsPanelExpanded] [{SchematicsGlobalSettingsPanelExpanded}]");
                    Logger.Info($"        [SchematicsShowTracesOnComponentSelect] [{SchematicsShowTracesOnComponentSelect}]");
                    Logger.Info($"        [SchematicsNetConnectionsPanelExpanded] [{SchematicsNetConnectionsPanelExpanded}]");
                    Logger.Info($"        [SchematicsImportantSignalsPanelExpanded] [{SchematicsImportantSignalsPanelExpanded}]");
                    Logger.Info($"        [SchematicsLabelBoard] [{SchematicsLabelBoard}]");
                    Logger.Info($"        [SchematicsLabelTechnical] [{SchematicsLabelTechnical}]");
                    Logger.Info($"        [SchematicsLabelFriendly] [{SchematicsLabelFriendly}]");
                    Logger.Info($"        [SchematicsLabelSelectedOnly] [{SchematicsLabelSelectedOnly}]");
                    Logger.Info($"        [SchematicsShowTracesOnSelectedComponent] [{SchematicsShowTracesOnSelectedComponent}]");
                    Logger.Info($"        [UserThemeColors] [{_data.UserThemeColors.Count} entries]");
                    Logger.Info($"        [WindowPlacement] [{_data.WindowState}] [{_data.WindowWidth:F0}x{_data.WindowHeight:F0}]");
                    Logger.Info($"    Various hardware/board specific settings:");
                    Logger.Info($"        [LastBoardByHardware] [{_data.LastBoardByHardware.Count} entries]");
                    Logger.Info($"        [LastHardware] [{_data.LastHardware}]");
                    Logger.Info($"        [LastSchematicByBoard] [{_data.LastSchematicByBoard.Count} entries]");
                    Logger.Info($"        [SchematicsSplitterRatios] [{_data.SchematicsSplitterRatios.Count} entries]");
                    Logger.Info($"        [SchematicsHoverHighlightsTracesByBoard] [{_data.SchematicsHoverHighlightsTracesByBoard.Count} entries]");
                    Logger.Info($"        [ContributorModeByBoard] [{_data.ContributorModeByBoard.Count} entries]");
                    Logger.Info($"        [SelectedCategoriesByBoard] [{_data.SelectedCategoriesByBoard.Count} entries]");
                    Logger.Info($"        [ActiveWorkbookIdByBoard] [{_data.ActiveWorkbookIdByBoard.Count} entries]");
                    Logger.Info($"        [OscilloscopeSeriesByVendor] [{_data.LastOscilloscopeSeriesByVendor.Count} entries]");
                    Logger.Info($"        [OscilloscopeVendor] [{_data.LastOscilloscopeVendor}]");
                    Logger.Info($"        [OscilloscopeHost] [{_data.OscilloscopeHost}]");
                    Logger.Info($"        [OscilloscopePort] [{_data.OscilloscopePort}]");
                    Logger.Info($"        [OscilloscopeSyncEnabled] [{ComponentInfoOscilloscopeSyncEnabled}]");

                    if (addedUserThemeDefaults)
                    {
                        Save();
                        Logger.Info("Missing [UserThemeColors] entries were added to configuration");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to load settings: [{ex.Message}] - using defaults");
            }
        }

        private static readonly JsonSerializerOptions SaveJsonOptions = new() { WriteIndented = true };

        // ###########################################################################################
        // Serializes current settings and writes them to the JSON file. The write goes to a
        // sibling ".tmp" file which is then swapped over the real one (the same pattern
        // OnlineServices.DownloadFileAsync uses): an in-place write truncates first, so a crash,
        // power loss or full disk mid-write would wipe every preference at once and the next
        // launch would silently fall back to defaults. With the swap, a failed save loses only
        // that one save - the previous file stays intact and parseable.
        // ###########################################################################################
        private static void Save()
        {
            if (string.IsNullOrEmpty(_settingsFilePath))
                return;

            AtomicJsonFile.Write(_settingsFilePath, _data, SaveJsonOptions, "settings");
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
        // Persists whether oscilloscope auto-connect is enabled.
        // ###########################################################################################
        public static bool OscilloscopeAutoConnect
        {
            get => _data.OscilloscopeAutoConnect;
            set
            {
                if (_data.OscilloscopeAutoConnect == value)
                    return;

                _data.OscilloscopeAutoConnect = value;
                Logger.Info($"Setting changed: [OscilloscopeAutoConnect] [{value}]");
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
            Logger.Info($"Setting changed: [OscilloscopeVendor] [{vendorName}]");
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
            Logger.Info($"Setting changed: [OscilloscopeSeriesByVendor] [{vendorName}] [{seriesName}]");
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

        // ###########################################################################################
        // Persists the selected oscilloscope image folder path.
        // When no explicit folder has been saved, the effective default becomes the current data root.
        // ###########################################################################################
        public static string OscilloscopeImageFolder
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_data.OscilloscopeImageFolder))
                {
                    return _data.OscilloscopeImageFolder;
                }

                return string.IsNullOrWhiteSpace(DataManager.DataRoot)
                    ? string.Empty
                    : DataManager.DataRoot;
            }
            set
            {
                value ??= string.Empty;

                if (string.Equals(_data.OscilloscopeImageFolder, value, StringComparison.Ordinal))
                {
                    return;
                }

                _data.OscilloscopeImageFolder = value;
                Logger.Info($"Setting changed: [OscilloscopeImageFolder] [{(string.IsNullOrWhiteSpace(value) ? "empty" : value)}]");
                Save();
            }
        }

        // ###########################################################################################
        // Returns whether contributor mode is enabled for the given board.
        // Falls back to the legacy global setting when a board-specific value is not yet saved.
        // ###########################################################################################
        public static bool GetContributorModeForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return _data.ContributorMode ?? false;

            return _data.ContributorModeByBoard.TryGetValue(boardKey, out var enabled)
                ? enabled
                : (_data.ContributorMode ?? false);
        }

        // ###########################################################################################
        // Persists contributor mode for the given board.
        // ###########################################################################################
        public static void SetContributorModeForBoard(string boardKey, bool enabled)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            _data.ContributorModeByBoard[boardKey] = enabled;
            Logger.Info($"Setting changed: [ContributorModeByBoard] [{boardKey}] [{enabled}]");
            Save();
        }

        // ###########################################################################################
        // Reloads only the user preference theme colors from the settings file while the app runs.
        // ###########################################################################################
        public static void ReloadUserThemeColors()
        {
            if (string.IsNullOrWhiteSpace(_settingsFilePath))
            {
                return;
            }

            try
            {
                if (!File.Exists(_settingsFilePath))
                {
                    var createdDefaults = EnsureUserThemeColors();

                    if (createdDefaults)
                    {
                        Save();
                        Logger.Info($"Created [UserThemeColors] defaults [{_data.UserThemeColors.Count} entries]");
                    }

                    return;
                }

                var json = File.ReadAllText(_settingsFilePath);
                var loaded = JsonSerializer.Deserialize<UserSettingsData>(json);

                if (loaded == null)
                {
                    Logger.Warning("Failed to reload user theme colors: settings file could not be deserialized");
                    return;
                }

                _data.UserThemeColors = loaded.UserThemeColors ?? new Dictionary<string, string>();

                var addedMissingDefaults = EnsureUserThemeColors();

                if (addedMissingDefaults)
                {
                    Save();
                    Logger.Info("Missing [UserThemeColors] entries were added to configuration during reload");
                }

                Logger.Info($"Reloaded [UserThemeColors] [{_data.UserThemeColors.Count} entries]");
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to reload user theme colors: [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Returns whether pin 1 should be specially marked for selected components on the given board.
        // Defaults to true unless explicitly disabled for that board.
        // ###########################################################################################
        public static bool GetSchematicsMarkPin1OnSelectedComponentForBoard(string boardKey)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return true;

            return _data.SchematicsMarkPin1OnSelectedComponentByBoard.TryGetValue(boardKey, out var enabled)
                ? enabled
                : true;
        }

        // ###########################################################################################
        // Persists whether pin 1 should be specially marked for selected components on the given board.
        // ###########################################################################################
        public static void SetSchematicsMarkPin1OnSelectedComponentForBoard(string boardKey, bool enabled)
        {
            if (string.IsNullOrWhiteSpace(boardKey))
                return;

            _data.SchematicsMarkPin1OnSelectedComponentByBoard[boardKey] = enabled;
            Logger.Info($"Setting changed: [SchematicsMarkPin1OnSelectedComponent] [{boardKey}] [{enabled}]");
            Save();
        }

    }
}