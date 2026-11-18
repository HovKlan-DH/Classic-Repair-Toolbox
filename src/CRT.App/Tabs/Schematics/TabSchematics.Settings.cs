using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// Board-level and global setting rows shown in the Schematics tab, and restoring/applying
// those settings when the board or a toggle changes.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private bool thisSuppressBoardSettingsChanged;

    private bool thisSuppressGlobalSettingsChanged;

    // ###########################################################################################
    // Handles row clicks for the temporary KiCad calibration visibility toggle that hides or shows
    // the rendered traces and pads while keeping the calibration box visible.
    // ###########################################################################################
    private void OnGlobalShowCalibrationTracesAndPadsRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckGlobalShowCalibrationTracesAndPads.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckGlobalShowCalibrationTracesAndPads.IsChecked =
                this.CheckGlobalShowCalibrationTracesAndPads.IsChecked != true;

            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Handle manual row clicks for scaled label visibilities.
    // ###########################################################################################
    private void OnLabelBoardRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelBoard.IsChecked = !this.CheckLabelBoard.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelTechnicalRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelTechnical.IsChecked = !this.CheckLabelTechnical.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelFriendlyRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelFriendly.IsChecked = !this.CheckLabelFriendly.IsChecked;
            e.Handled = true;
        }
    }

    private void OnLabelSelectedOnlyRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckLabelSelectedOnly.IsChecked = !this.CheckLabelSelectedOnly.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Handle manual row clicks for board-specific schematic settings.
    // ###########################################################################################
    private void OnGlobalHoverHighlightsTracesRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckGlobalHoverHighlightsTraces.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckGlobalHoverHighlightsTraces.IsChecked = !this.CheckGlobalHoverHighlightsTraces.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Restores the settings panel's rows from persisted configuration - none is per board since "Mark first pin on component"
    // went (2026-11-17).
    // ###########################################################################################
    private void RestoreBoardSettings()
    {
        this.thisSuppressBoardSettingsChanged = true;
        this.thisSuppressGlobalSettingsChanged = true;

        this.CheckBoardShowTracesOnSelectedComponent.IsChecked = UserSettings.SchematicsShowTracesOnSelectedComponent;
        this.CheckGlobalShowTracesOnComponentSelect.IsChecked = UserSettings.SchematicsShowTracesOnComponentSelect;
        this.CheckGlobalShowOppositeSideTraces.IsChecked = UserSettings.SchematicsShowOppositeSideTraces;
        this.CheckGlobalShowZones.IsChecked = UserSettings.SchematicsShowZones;

        bool isInteractiveCadTraceHoverEnabled =
            !string.Equals(UserSettings.InteractiveCadTraceHoverMode, "Disabled", StringComparison.Ordinal);

        bool isInteractiveCadTraceHoverHoldShiftMode =
            string.Equals(UserSettings.InteractiveCadTraceHoverMode, "HoldShift", StringComparison.Ordinal);

        this.CheckGlobalHoverHighlightsTraces.IsChecked = isInteractiveCadTraceHoverEnabled;

        bool shouldRestoreHoldShiftCheckState =
            isInteractiveCadTraceHoverEnabled ||
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked is null;

        if (shouldRestoreHoldShiftCheckState)
        {
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked = isInteractiveCadTraceHoverHoldShiftMode;
        }

        this.UpdateSchematicsInteractiveCadTraceHoverHoldShiftVisualState(isInteractiveCadTraceHoverEnabled);

        this.thisSuppressGlobalSettingsChanged = false;
        this.thisSuppressBoardSettingsChanged = false;

        this.UpdateInteractiveCadTraceHoverModeUi();
    }

    // ###########################################################################################
    // Returns true when the current schematic allows hover-driven KiCad trace highlighting.
    // ###########################################################################################
    private bool IsBoardHoverHighlightsTracesEnabled()
    {
        if (!this.HasCurrentSchematicKiCadTraces())
        {
            return false;
        }

        return UserSettings.InteractiveCadTraceHoverMode switch
        {
            "Disabled" => false,
            "HoldShift" => this.thisIsInteractiveCadTraceHoverShiftPressed,
            _ => true
        };
    }

    // ###########################################################################################
    // Whether pin 1 of the components selected or pointed at is marked in orange: ALWAYS, on a schematic image showing the
    // board's KiCad pads (owner decision, 2026-11-17: "remove the "Mark first pin on component" checkbox, but keep its
    // functionality enabled per default, so everyone gets this" - a box per board, ticked unless unticked, until then; the
    // pins numbered with SHIFT held mark it on their own badge, over it).
    // ###########################################################################################
    private bool IsMarkingFirstPin() => this.HasCurrentSchematicKiCadPcbPadData();

    // ###########################################################################################
    // Handle manual row clicks for board-specific selected-component trace preview.
    // ###########################################################################################
    private void OnBoardShowTracesOnSelectedComponentRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckBoardShowTracesOnSelectedComponent.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckBoardShowTracesOnSelectedComponent.IsChecked = !this.CheckBoardShowTracesOnSelectedComponent.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Returns true when hovered components should preview the same traces as a selected component.
    // Uses the global checkbox state instead of a board-specific setting.
    // ###########################################################################################
    private bool IsBoardShowTracesOnSelectedComponentEnabled()
    {
        if (!this.HasCurrentSchematicKiCadTraces())
        {
            return false;
        }

        return UserSettings.SchematicsShowTracesOnSelectedComponent;
    }

    // ###########################################################################################
    // Syncs the copied global settings controls from persisted user settings.
    // ###########################################################################################
    private void UpdateGlobalSettingsControls()
    {
        this.thisSuppressGlobalSettingsChanged = true;

        bool isInteractiveCadTraceHoverEnabled =
            !string.Equals(UserSettings.InteractiveCadTraceHoverMode, "Disabled", StringComparison.Ordinal);

        bool isInteractiveCadTraceHoverHoldShiftMode =
            string.Equals(UserSettings.InteractiveCadTraceHoverMode, "HoldShift", StringComparison.Ordinal);

        this.CheckGlobalHoverHighlightsTraces.IsChecked = isInteractiveCadTraceHoverEnabled;
        this.CheckGlobalShowOppositeSideTraces.IsChecked = UserSettings.SchematicsShowOppositeSideTraces;
        this.CheckGlobalShowZones.IsChecked = UserSettings.SchematicsShowZones;

        bool shouldRestoreHoldShiftCheckState =
            isInteractiveCadTraceHoverEnabled ||
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked is null;

        if (shouldRestoreHoldShiftCheckState)
        {
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked = isInteractiveCadTraceHoverHoldShiftMode;
        }

        this.UpdateSchematicsInteractiveCadTraceHoverHoldShiftVisualState(isInteractiveCadTraceHoverEnabled);

        this.thisSuppressGlobalSettingsChanged = false;
    }

    // ###########################################################################################
    // Handle manual row clicks for selected-component trace preview.
    // ###########################################################################################
    private void OnGlobalShowTracesOnComponentSelectRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckGlobalShowTracesOnComponentSelect.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckGlobalShowTracesOnComponentSelect.IsChecked = !this.CheckGlobalShowTracesOnComponentSelect.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Handle manual row clicks for opposite-side PCB trace preview.
    // ###########################################################################################
    private void OnGlobalShowOppositeSideTracesRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckGlobalShowOppositeSideTraces.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckGlobalShowOppositeSideTraces.IsChecked = !this.CheckGlobalShowOppositeSideTraces.IsChecked;
            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Handle manual row clicks for global KiCad zone visibility.
    // ###########################################################################################
    private void OnGlobalShowZonesRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.CheckGlobalShowZones.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.CheckGlobalShowZones.IsChecked = !this.CheckGlobalShowZones.IsChecked;
            e.Handled = true;
        }
    }
}