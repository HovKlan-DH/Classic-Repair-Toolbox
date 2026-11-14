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

    // The board those lines were made from - compared by reference, never by a hash code two boards can share (code review,
    // 2026-11-13: a board read again under the old one's identity hash kept the old names).
    private BoardData? thisLastKiCadNetConnectionsBoard;

    // Whether those lines showed the panel - none for a net with no pads on the board.
    private bool thisLastKiCadNetConnectionsShown;

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
            Logger.Debug("Important signals: panel rebuild aborted because board data or KiCad PCB data is unavailable");
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

        // The one rule for how a row's "KiCad net name" finds its net - shared with the Drafts tab's
        // KiCad data window, which lists the names and the rows that find none (2026-10-10).
        var nets = root.Pcb
            .SelectMany(pcb => pcb.Nets.List)
            .Select(net => (net.Name, net.NormalizedName, net.IsNoNet))
            .ToList();

        var normalizedNameByAnyKnownName = KiCadNetNames.Resolver(nets.Select(net => (net.Name, net.NormalizedName)));

        // Written always, as debug lines, since "Write a detailed log" went (owner request, 2026-11-07) - the counts, and
        // each row that finds no net; never every net's name, hundreds per board shown.
        Logger.Debug(
            $"Important signals: [{boardData.KiCadImportantSignals.Count}] rows, [{knownNetNames.Count}] KiCad net names, " +
            $"[{normalizedNameByAnyKnownName.Count}] name mappings, [{KiCadNetNames.NormalizedNames(nets).Count}] normalized nets");

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
                Logger.Debug($"Important signals: skipped a row with DisplayName=[{displayName}] KiCadNetName=[{kiCadNetName}]");
                continue;
            }

            if (!normalizedNameByAnyKnownName.TryGetValue(kiCadNetName, out string? normalizedNetName) ||
                string.IsNullOrWhiteSpace(normalizedNetName))
            {
                unresolvedRowCount++;
                Logger.Debug($"Important signals: KiCad net name [{kiCadNetName}] of [{displayName}] matches no KiCad net of the board");
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
                Logger.Debug($"Important signals: [{displayName}] -> [{normalizedNetName}] is listed twice");
            }

            netNames.Add(normalizedNetName);
            resolvedRowCount++;
        }

        Logger.Debug(
            $"Important signals: [{resolvedRowCount}] rows found their net, [{unresolvedRowCount}] found none, [{invalidRowCount}] skipped, " +
            $"[{duplicateResolvedDisplayNameToNetPairs.Count}] listed twice, [{this.thisImportantSignalDisplayNames.Count}] signals shown");

        if (this.thisImportantSignalDisplayNames.Count == 0)
        {
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
            this.KiCadNetConnectionsHeaderTextBlock.Text = "KiCad nets";
            this.KiCadNetConnectionsNetNameText.Text = string.Empty;
            this.KiCadNetConnectionsList.ItemsSource = null;
            this.KiCadNetConnectionsPanel.IsVisible = false;
            this.UpdateKiCadNetConnectionsClearButtonState(false);
            return;
        }

        var sortedNetNames = activeNetLookup
            .OrderBy(netName => netName, StringComparer.OrdinalIgnoreCase)
            .ToList();

        // The nets, and what the lines name each component by: the board's rows in the region shown (KiCadNetConnections) -
        // another region, or the board read again after an edit, makes them again.
        var boardShown = this.MainWindow?.CurrentBoardData;
        string region = this.MainWindow?.LocalRegion?.Trim() ?? string.Empty;
        string signature = string.Join("\u001F", sortedNetNames) + "\u001E" + region;

        if (string.Equals(this.thisLastKiCadNetConnectionsSignature, signature, StringComparison.Ordinal) &&
            ReferenceEquals(this.thisLastKiCadNetConnectionsBoard, boardShown))
        {
            // As the lines were made: a net with no pads on the board shows no panel, at the next refresh too (code review,
            // 2026-11-13: it came back empty under "KiCad nets (1)", "Clear" on).
            this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
                ? $"KiCad nets ({sortedNetNames.Count})"
                : "KiCad nets";
            this.KiCadNetConnectionsPanel.IsVisible = this.thisLastKiCadNetConnectionsShown;
            this.UpdateKiCadNetConnectionsClearButtonState(this.thisLastKiCadNetConnectionsShown);
            return;
        }

        // "[U5 | CIA 1 | 6581] pin 13", the label and the pin bold - each component named as its component info window's
        // title names it (owner requests, 2026-11-13; KiCadNetConnections).
        List<KiCadNetConnectionLine> sortedConnections = KiCadNetConnections.Lines(
            this.thisKiCadProject.Root.Pcb
                .SelectMany(pcb => pcb.Footprints)
                .SelectMany(footprint => footprint.Pads.Select(pad => new KiCadPadOnNet(
                    footprint.Reference ?? string.Empty,
                    pad.Number ?? string.Empty,
                    pad.Net?.NormalizedName ?? string.Empty))),
            activeNetLookup,
            boardShown?.Components,
            region);

        if (sortedConnections.Count == 0)
        {
            this.thisLastKiCadNetConnectionsSignature = signature;
            this.thisLastKiCadNetConnectionsBoard = boardShown;
            this.thisLastKiCadNetConnectionsShown = false;
            this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
                ? $"KiCad nets ({sortedNetNames.Count})"
                : "KiCad nets";
            this.KiCadNetConnectionsNetNameText.Text = string.Join(Environment.NewLine, sortedNetNames);
            this.KiCadNetConnectionsList.ItemsSource = null;
            this.KiCadNetConnectionsPanel.IsVisible = false;
            this.UpdateKiCadNetConnectionsClearButtonState(false);
            return;
        }

        this.thisLastKiCadNetConnectionsSignature = signature;
        this.thisLastKiCadNetConnectionsBoard = boardShown;
        this.thisLastKiCadNetConnectionsShown = true;
        this.KiCadNetConnectionsHeaderTextBlock.Text = sortedNetNames.Count > 0
            ? $"KiCad nets ({sortedNetNames.Count})"
            : "KiCad nets";
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
            ? $"Important KiCad signals ({selectedCount})"
            : "Important KiCad signals";
    }
}