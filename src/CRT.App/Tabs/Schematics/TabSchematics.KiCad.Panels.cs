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
// The two KiCad side panels: 'Important signals' and 'Net connections' - building their
// rows, syncing selection, and their expand/clear/header state.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private readonly Dictionary<string, HashSet<string>> thisImportantSignalNetNamesByDisplayName =
    new(StringComparer.OrdinalIgnoreCase);

    private readonly List<string> thisImportantSignalDisplayNames = new();

    private readonly HashSet<string> thisSelectedImportantSignalDisplayNames =
        new(StringComparer.OrdinalIgnoreCase);

    private string thisLastKiCadNetConnectionsSignature = string.Empty;

    // ###########################################################################################
    // Rebuilds the grouped "Important signals" panel from the current board data and available KiCad
    // net names. Only display groups that can resolve to at least one real KiCad normalized net name.
    // Contributor mode logs duplicate groups, duplicate mappings, unresolved rows, and final totals.
    // ###########################################################################################
    private void RebuildImportantSignalsPanel()
    {
        this.thisImportantSignalNetNamesByDisplayName.Clear();
        this.thisImportantSignalDisplayNames.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();

        this.UpdateImportantSignalsHeaderText();
        this.ImportantSignalsListBox.ItemsSource = null;
        this.ImportantSignalsListBox.SelectedItems?.Clear();
        this.ImportantSignalsPanel.IsVisible = false;
        this.UpdateImportantSignalsClearButtonState(false);

        if (!this.HasCurrentSchematicKiCadTraces())
        {
            return;
        }

        var boardData = this.MainWindow?.CurrentBoardData;
        var root = this.thisKiCadProject?.Root;

        if (boardData == null || root == null || root.Pcb.Count == 0)
        {
            if (UserSettings.ContributorMode)
            {
                Logger.Warning("Important signals debug: panel rebuild aborted because board data or KiCad PCB data is unavailable");
            }

            return;
        }

        var knownNetNames = new HashSet<string>(
            root.Pcb
                .SelectMany(pcb => pcb.Nets.List)
                .SelectMany(net => new[]
                {
                    net.Name?.Trim() ?? string.Empty,
                    net.NormalizedName?.Trim() ?? string.Empty
                })
                .Where(netName => !string.IsNullOrWhiteSpace(netName)),
            StringComparer.OrdinalIgnoreCase);

        var normalizedNameByAnyKnownName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pcb in root.Pcb)
        {
            foreach (var net in pcb.Nets.List)
            {
                string normalizedName = net.NormalizedName?.Trim() ?? string.Empty;
                string name = net.Name?.Trim() ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(normalizedName))
                {
                    normalizedNameByAnyKnownName[normalizedName] = normalizedName;
                }

                if (!string.IsNullOrWhiteSpace(name) && !string.IsNullOrWhiteSpace(normalizedName))
                {
                    normalizedNameByAnyKnownName[name] = normalizedName;
                }
            }
        }

        bool isContributorMode = UserSettings.ContributorMode;

        if (isContributorMode)
        {
            Logger.Info($"Important signals debug: Excel rows loaded [{boardData.KiCadImportantSignals.Count}]");
            Logger.Info($"Important signals debug: loaded [{knownNetNames.Count}] raw KiCad net names");
            Logger.Info($"Important signals debug: built [{normalizedNameByAnyKnownName.Count}] KiCad net name mappings");

            var normalizedNetNames = root.Pcb
                .SelectMany(pcb => pcb.Nets.List)
                .Select(net => net.NormalizedName?.Trim() ?? string.Empty)
                .Where(netName => !string.IsNullOrWhiteSpace(netName))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(netName => netName, StringComparer.OrdinalIgnoreCase)
                .ToList();

            Logger.Info($"Important signals debug: loaded [{normalizedNetNames.Count}] normalized KiCad net names");

            foreach (string normalizedNetName in normalizedNetNames)
            {
                Logger.Info($"Important signals debug: normalized KiCad net name [{normalizedNetName}]");
            }
        }

        var excelRowCountByDisplayName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var seenResolvedDisplayNameToNetPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var duplicateResolvedDisplayNameToNetPairs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        int invalidRowCount = 0;
        int unresolvedRowCount = 0;
        int resolvedRowCount = 0;

        foreach (var entry in boardData.KiCadImportantSignals)
        {
            string displayName = entry.DisplayName?.Trim() ?? string.Empty;
            string kiCadNetName = entry.KiCadNetName?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(displayName) || string.IsNullOrWhiteSpace(kiCadNetName))
            {
                invalidRowCount++;

                if (isContributorMode)
                {
                    Logger.Warning(
                        $"Important signals debug: skipped invalid Excel row with DisplayName=[{displayName}] KiCadNetName=[{kiCadNetName}]");
                }

                continue;
            }

            if (excelRowCountByDisplayName.TryGetValue(displayName, out int currentDisplayNameCount))
            {
                excelRowCountByDisplayName[displayName] = currentDisplayNameCount + 1;
            }
            else
            {
                excelRowCountByDisplayName[displayName] = 1;
            }

            if (!normalizedNameByAnyKnownName.TryGetValue(kiCadNetName, out string? normalizedNetName) ||
                string.IsNullOrWhiteSpace(normalizedNetName))
            {
                unresolvedRowCount++;

                if (isContributorMode)
                {
                    Logger.Warning(
                        $"Important signals debug: Excel KiCad net name [{kiCadNetName}] for display name [{displayName}] did not match any loaded KiCad net name");
                }

                continue;
            }

            if (!this.thisImportantSignalNetNamesByDisplayName.TryGetValue(displayName, out var netNames))
            {
                netNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                this.thisImportantSignalNetNamesByDisplayName[displayName] = netNames;
                this.thisImportantSignalDisplayNames.Add(displayName);
            }

            string resolvedPairKey = $"{displayName}\u001F{normalizedNetName}";

            if (!seenResolvedDisplayNameToNetPairs.Add(resolvedPairKey))
            {
                duplicateResolvedDisplayNameToNetPairs.Add(resolvedPairKey);

                if (isContributorMode)
                {
                    Logger.Warning(
                        $"Important signals debug: duplicate resolved mapping detected for display name [{displayName}] -> normalized net [{normalizedNetName}]");
                }
            }

            netNames.Add(normalizedNetName);
            resolvedRowCount++;
        }

        if (isContributorMode)
        {
            foreach (var duplicateDisplayName in excelRowCountByDisplayName
                         .Where(entry => entry.Value > 1)
                         .OrderBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase))
            {
                int resolvedNetCount = this.thisImportantSignalNetNamesByDisplayName.TryGetValue(duplicateDisplayName.Key, out var resolvedNetNames)
                    ? resolvedNetNames.Count
                    : 0;

                Logger.Info(
                    $"Important signals debug: display name [{duplicateDisplayName.Key}] appears in [{duplicateDisplayName.Value}] Excel rows and resolves to [{resolvedNetCount}] unique KiCad nets");
            }

            Logger.Info($"Important signals debug: invalid Excel rows skipped [{invalidRowCount}]");
            Logger.Info($"Important signals debug: unresolved Excel rows [{unresolvedRowCount}]");
            Logger.Info($"Important signals debug: resolved Excel rows [{resolvedRowCount}]");
            Logger.Info($"Important signals debug: duplicate resolved mappings [{duplicateResolvedDisplayNameToNetPairs.Count}]");
            Logger.Info($"Important signals debug: resolved display groups [{this.thisImportantSignalDisplayNames.Count}]");
        }

        if (this.thisImportantSignalDisplayNames.Count == 0)
        {
            if (isContributorMode)
            {
                Logger.Warning("Important signals debug: no Excel signal rows could be resolved to loaded KiCad net names");
            }

            return;
        }

        var items = this.thisImportantSignalDisplayNames
            .Select(displayName =>
            {
                var mappedNetNames = this.thisImportantSignalNetNamesByDisplayName.TryGetValue(displayName, out var netNames)
                    ? netNames.OrderBy(netName => netName, StringComparer.OrdinalIgnoreCase).ToList()
                    : new List<string>();

                return new ImportantSignalListItem
                {
                    DisplayName = displayName,
                    ToolTipText = string.Join(", ", mappedNetNames)
                };
            })
            .ToList();

        this.UpdateImportantSignalsHeaderText();
        this.ImportantSignalsListBox.ItemsSource = items;
        this.ImportantSignalsPanel.IsVisible = true;
        this.UpdateImportantSignalsClearButtonState(true);

        if (isContributorMode)
        {
            Logger.Info($"Important signals debug: built [{items.Count}] visible important signal groups");

            foreach (var item in items)
            {
                Logger.Info($"Important signals debug: display name [{item.DisplayName}] -> [{item.ToolTipText}]");
            }
        }
    }

    // ###########################################################################################
    // Synchronizes the currently selected important signal display names from the visible list control.
    // ###########################################################################################
    private void SyncSelectedImportantSignalsFromList()
    {
        this.thisSelectedImportantSignalDisplayNames.Clear();

        foreach (var item in this.ImportantSignalsListBox.SelectedItems?.Cast<ImportantSignalListItem>() ?? Enumerable.Empty<ImportantSignalListItem>())
        {
            string displayName = item.DisplayName?.Trim() ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(displayName))
            {
                this.thisSelectedImportantSignalDisplayNames.Add(displayName);
            }
        }

        this.UpdateImportantSignalsHeaderText();
        this.UpdateImportantSignalsClearButtonState(this.ImportantSignalsPanel.IsVisible);
    }

    // ###########################################################################################
    // Clears the selected important signal list items and also clears any matching explicit KiCad
    // net selections so overlapping net picks can be cleared from either panel.
    // ###########################################################################################
    private void ClearImportantSignalsSelection()
    {
        var selectedNetNames = this.BuildSelectedImportantSignalNetNames();

        this.ImportantSignalsListBox.SelectedItems?.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();
        this.ClearKiCadTraceSelectionsForNetNames(selectedNetNames);
        this.UpdateImportantSignalsHeaderText();
        this.UpdateImportantSignalsClearButtonState(this.ImportantSignalsPanel.IsVisible);
        this.RefreshKiCadOverlay(forceImmediate: true);
        this.RefreshBlinkStateFromCurrentSelection();
    }

    // ###########################################################################################
    // Builds the effective set of normalized KiCad net names selected through the "Important signals"
    // panel by expanding each selected display group to all mapped KiCad net names.
    // ###########################################################################################
    private HashSet<string> BuildSelectedImportantSignalNetNames()
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (string displayName in this.thisSelectedImportantSignalDisplayNames)
        {
            if (!this.thisImportantSignalNetNamesByDisplayName.TryGetValue(displayName, out var netNames))
            {
                continue;
            }

            result.UnionWith(netNames.Where(netName => !string.IsNullOrWhiteSpace(netName)));
        }

        return result;
    }

    // ###########################################################################################
    // Returns true when the pointer is currently inside the KiCad net connections panel bounds.
    // This prevents hover and click handling from reacting to traces or components behind it.
    // ###########################################################################################
    private bool IsPointerInsideKiCadNetConnectionsPanel(Point containerPoint)
    {
        if (!this.KiCadNetConnectionsPanel.IsVisible)
        {
            return false;
        }

        Point? translatedTopLeft = this.KiCadNetConnectionsPanel.TranslatePoint(new Point(0, 0), this.SchematicsContainer);
        if (!translatedTopLeft.HasValue)
        {
            return false;
        }

        var panelRect = new Rect(translatedTopLeft.Value, this.KiCadNetConnectionsPanel.Bounds.Size);
        return panelRect.Contains(containerPoint);
    }

    // ###########################################################################################
    // Clears transient hover state while the pointer is inside the KiCad net connections panel.
    // Explicitly selected or locked KiCad nets remain visible and are not affected.
    // ###########################################################################################
    private void ClearTransientHoverForKiCadNetConnectionsPanel()
    {
        this.SetHoveredComponentBoardLabel(null);
        this.SetHoveredKiCadNet(null);
        this.thisHoveredKiCadPadNumber = null;

        this.ResetSchematicsHoverLabelToDefaultAppearance();
        this.SchematicsHoverLabelBorder.IsVisible = false;
        this.SchematicsHoverLabelText.Text = string.Empty;
        this.SchematicsHoverPadBorder.IsVisible = false;
        this.SchematicsHoverPadText.Text = string.Empty;
        this.SchematicsContainer.Cursor = Cursor.Default;

        if (this.MainWindow != null)
        {
            this.MainWindow.isHoveringComponent = false;
        }
    }

    // ###########################################################################################
    // Updates the panel displaying connected components and pins for a supplied active net set.
    // This avoids recomputing hover-preview net names multiple times within one overlay refresh.
    // Keeps the clear button visible even when the panel body is collapsed, as long as the panel
    // itself is populated with netlist content.
    // ###########################################################################################
    private void UpdateKiCadNetConnectionsPanel(IReadOnlyCollection<string> activeNets)
    {
        var activeNetLookup = activeNets as HashSet<string> ??
                              new HashSet<string>(activeNets, StringComparer.OrdinalIgnoreCase);

        if (!this.HasCurrentSchematicKiCadTraces() ||
            activeNetLookup.Count == 0 ||
            this.thisKiCadProject?.Root == null)
        {
            this.thisLastKiCadNetConnectionsSignature = string.Empty;
            this.KiCadNetConnectionsHeaderTextBlock.Text = "Netlist names";
            this.KiCadNetConnectionsNetNameText.Text = string.Empty;
            this.KiCadNetConnectionsList.ItemsSource = null;
            this.KiCadNetConnectionsPanel.IsVisible = false;
            this.UpdateKiCadNetConnectionsClearButtonState(false);
            return;
        }

        var sortedNetNames = activeNetLookup
            .OrderBy(netName => netName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        string signature = string.Join("\u001F", sortedNetNames);

        if (string.Equals(this.thisLastKiCadNetConnectionsSignature, signature, StringComparison.Ordinal))
        {
            this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
                ? $"Netlist names ({sortedNetNames.Count})"
                : "Netlist names";
            this.KiCadNetConnectionsPanel.IsVisible = true;
            this.UpdateKiCadNetConnectionsClearButtonState(true);
            return;
        }

        var boardComponents = this.MainWindow?.CurrentBoardData?.Components;
        var orderLookup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        if (boardComponents != null)
        {
            for (int i = 0; i < boardComponents.Count; i++)
            {
                string boardLabel = boardComponents[i].BoardLabel?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(boardLabel) && !orderLookup.ContainsKey(boardLabel))
                {
                    orderLookup[boardLabel] = i;
                }
            }
        }

        var connections = new List<(string ConnStr, int OrderIndex, int PadIndex)>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pcb in this.thisKiCadProject.Root.Pcb)
        {
            foreach (var footprint in pcb.Footprints)
            {
                string refName = footprint.Reference?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(refName))
                {
                    continue;
                }

                foreach (var pad in footprint.Pads)
                {
                    string netName = pad.Net?.NormalizedName?.Trim() ?? string.Empty;

                    if (!activeNetLookup.Contains(netName))
                    {
                        continue;
                    }

                    string padNum = pad.Number?.Trim() ?? "?";

                    string connStr = activeNetLookup.Count > 1
                        ? $"{refName} pin {padNum} [{netName}]"
                        : $"{refName} pin {padNum}";

                    if (!seen.Add(connStr))
                    {
                        continue;
                    }

                    int orderIndex = orderLookup.TryGetValue(refName, out int idx) ? idx : int.MaxValue;
                    int padIndex = int.TryParse(padNum, out int parsedPad) ? parsedPad : int.MaxValue;

                    connections.Add((connStr, orderIndex, padIndex));
                }
            }
        }

        if (connections.Count == 0)
        {
            this.thisLastKiCadNetConnectionsSignature = signature;
            this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
                ? $"Netlist names ({sortedNetNames.Count})"
                : "Netlist names";
            this.KiCadNetConnectionsNetNameText.Text = string.Join(Environment.NewLine, sortedNetNames);
            this.KiCadNetConnectionsList.ItemsSource = null;
            this.KiCadNetConnectionsPanel.IsVisible = false;
            this.UpdateKiCadNetConnectionsClearButtonState(false);
            return;
        }

        var sortedConnections = connections
            .OrderBy(connection => connection.OrderIndex)
            .ThenBy(connection => connection.PadIndex)
            .ThenBy(connection => connection.ConnStr, StringComparer.OrdinalIgnoreCase)
            .Select(connection => connection.ConnStr)
            .ToList();

        this.thisLastKiCadNetConnectionsSignature = signature;
        this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
            ? $"Netlist names ({sortedNetNames.Count})"
            : "Netlist names";
        this.KiCadNetConnectionsNetNameText.Text = string.Join(Environment.NewLine, sortedNetNames);
        this.KiCadNetConnectionsList.ItemsSource = sortedConnections;
        this.KiCadNetConnectionsPanel.IsVisible = true;
        this.UpdateKiCadNetConnectionsClearButtonState(true);
    }

    // ###########################################################################################
    // Applies the current expand/collapse state to the KiCad net connections panel content.
    // ###########################################################################################
    private void UpdateKiCadNetConnectionsPanelExpandedState(bool isExpanded)
    {
        this.KiCadNetConnectionsContentPanel.IsVisible = isExpanded;
        this.ToggleKiCadNetConnectionsPanelButton.Content = isExpanded ? "Collapse" : "Expand";
    }

    // ###########################################################################################
    // Updates the visibility and enabled state of the KiCad trace-clear button.
    // The button only clears explicit KiCad net picks and Important signal selections.
    // ###########################################################################################
    private void UpdateKiCadNetConnectionsClearButtonState(bool hasVisibleNetlistContent)
    {
        bool hasAnythingToClear =
            this.thisLockedKiCadNetNames.Count > 0 ||
            this.thisSelectedImportantSignalDisplayNames.Count > 0;

        this.ClearKiCadTraceSelectionButton.IsVisible = hasAnythingToClear;
        this.ClearKiCadTraceSelectionButton.IsEnabled = hasAnythingToClear;
    }

    // ###########################################################################################
    // Applies the current expand/collapse state to the Important signals panel content.
    // ###########################################################################################
    private void UpdateImportantSignalsPanelExpandedState(bool isExpanded)
    {
        this.ImportantSignalsListPanel.IsVisible = isExpanded;
        this.ToggleImportantSignalsPanelButton.Content = isExpanded ? "Collapse" : "Expand";
    }

    // ###########################################################################################
    // Updates the visibility and enabled state of the Important signals clear button.
    // The button stays visible whenever the panel has content, but is disabled when nothing is selected.
    // When the panel is collapsed and nothing is selected, the button is hidden to save space.
    // ###########################################################################################
    private void UpdateImportantSignalsClearButtonState(bool hasVisibleImportantSignalsContent)
    {
        bool hasAnythingToClear = this.thisSelectedImportantSignalDisplayNames.Count > 0;
        bool isExpanded = this.ImportantSignalsListPanel.IsVisible;

        bool shouldShowButton = hasVisibleImportantSignalsContent && (isExpanded || hasAnythingToClear);

        this.ClearImportantSignalsButton.IsVisible = shouldShowButton;
        this.ClearImportantSignalsButton.IsEnabled = hasAnythingToClear;
    }

    // ###########################################################################################
    // Updates the Important signals header so it shows only the selected-signal count.
    // When nothing is selected, the header falls back to the base panel title.
    // ###########################################################################################
    private void UpdateImportantSignalsHeaderText()
    {
        int selectedCount = this.thisSelectedImportantSignalDisplayNames.Count;

        this.ImportantSignalsHeaderTextBlock.Text = selectedCount > 0
            ? $"Important signals ({selectedCount})"
            : "Important signals";
    }
}