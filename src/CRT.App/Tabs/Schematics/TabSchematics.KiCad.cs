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
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// KiCad project state: loading the project for the current board, mapping board labels to
// nets and references, the current selection/lock sets, and the per-board runtime cache scopes.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // KiCad import / overlay
    private KiCadProjectBundle? thisKiCadProject;

    private readonly HashSet<string> thisSelectedKiCadReferences = new(StringComparer.OrdinalIgnoreCase);

    private readonly HashSet<string> thisSelectedKiCadNormalizedNetNames = new(StringComparer.OrdinalIgnoreCase);

    private readonly HashSet<string> thisLockedKiCadNetNames = new(StringComparer.OrdinalIgnoreCase);

    private string thisCurrentKiCadRuntimeCacheScopeKey = string.Empty;

    private readonly LinkedList<string> thisKiCadRuntimeCacheScopeLru = new();

    private readonly Dictionary<string, KiCadRuntimeCacheScope> thisKiCadRuntimeCacheScopeByKey =
        new(StringComparer.OrdinalIgnoreCase);

    private int thisKiCadProjectLoadVersion;

    // ###########################################################################################
    // Returns true when the requested schematic contains any explicitly selected KiCad net content.
    // Only explicit KiCad trace selections participate here, so component-driven thumbnail matching
    // stays reliable and depends on real component/image presence instead of shared net names.
    // Zones are included so copper pours participate in thumbnail dimming and selection presence.
    // ###########################################################################################
    private bool DoesSchematicContainSelectedKiCadContent(string schematicName)
    {
        if (string.IsNullOrWhiteSpace(schematicName) || this.thisKiCadProject == null)
        {
            return false;
        }

        var selectedNets = this.BuildKiCadThumbnailDimmingNetNames();
        if (selectedNets.Count == 0)
        {
            return false;
        }

        var view = this.ResolveKiCadViewForSchematicName(schematicName);
        if (view == null)
        {
            return false;
        }

        if (string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
        {
            var root = this.thisKiCadProject.Root;
            if (view.SourceIndex < 0 || view.SourceIndex >= root.Pcb.Count)
            {
                return false;
            }

            var pcb = root.Pcb[view.SourceIndex];

            foreach (var net in pcb.Nets.List)
            {
                string normalizedName = net.NormalizedName?.Trim() ?? string.Empty;
                string netId = net.Id?.Trim() ?? string.Empty;

                if (string.IsNullOrWhiteSpace(normalizedName) ||
                    !selectedNets.Contains(normalizedName) ||
                    string.IsNullOrWhiteSpace(netId))
                {
                    continue;
                }

                if (!pcb.HighlightIndex.TryGetValue(netId, out var bucket))
                {
                    continue;
                }

                if (bucket.Pads.Count > 0 ||
                    bucket.Segments.Count > 0 ||
                    bucket.Vias.Count > 0 ||
                    bucket.Arcs.Count > 0 ||
                    bucket.Zones.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        if (string.Equals(view.Type, "schematic", StringComparison.OrdinalIgnoreCase))
        {
            if (!this.thisKiCadProject.SchematicNetPathIndexBySchematicIndex.TryGetValue(view.SourceIndex, out var indexByNet))
            {
                return false;
            }

            foreach (string selectedNet in selectedNets)
            {
                if (!indexByNet.TryGetValue(selectedNet, out var resolvedPaths))
                {
                    continue;
                }

                if (resolvedPaths.Any(path => path.Points.Count >= 2))
                {
                    return true;
                }
            }
        }

        return false;
    }

    // ###########################################################################################
    // Resolves any schematic thumbnail name to the corresponding KiCad project view.
    // Top and bottom replica pages are matched explicitly, while schematic pages are matched by
    // exact display name first and by page ordinal second.
    // ###########################################################################################
    private KiCadProjectView? ResolveKiCadViewForSchematicName(string schematicName)
    {
        var project = this.thisKiCadProject?.Root.Project;
        if (project == null || project.Views.Count == 0 || string.IsNullOrWhiteSpace(schematicName))
        {
            return null;
        }

        string targetName = schematicName;
        if (this.schematicByName.TryGetValue(schematicName, out var entry) && !string.IsNullOrWhiteSpace(entry.CadName))
        {
            targetName = entry.CadName.Trim();
        }

        var exact = project.Views.FirstOrDefault(view =>
            string.Equals(view.DisplayName, targetName, StringComparison.OrdinalIgnoreCase));

        if (exact != null)
        {
            return exact;
        }

        if (string.Equals(targetName, "Top (replica)", StringComparison.OrdinalIgnoreCase))
        {
            return project.Views.FirstOrDefault(view =>
                string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase));
        }

        if (string.Equals(targetName, "Bottom (replica)", StringComparison.OrdinalIgnoreCase))
        {
            return project.Views.FirstOrDefault(view =>
                string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase));
        }

        if (KiCadLayerGeometry.TryExtractSchematicPageOrdinal(targetName, out int pageOrdinal))
        {
            var schematicViews = project.Views
                .Where(view => string.Equals(view.Type, "schematic", StringComparison.OrdinalIgnoreCase))
                .OrderBy(view => view.SourceIndex)
                .ToList();

            int targetIndex = pageOrdinal - 1;
            if (targetIndex >= 0 && targetIndex < schematicViews.Count)
            {
                return schematicViews[targetIndex];
            }
        }

        return null;
    }

    // ###########################################################################################
    // Loads raw KiCad overlay data for the current board without blocking the initial board UI.
    // Establishes a persistent per-board runtime cache scope so revisiting the same KiCad project
    // can reuse already-built render and hover caches across board switches in this session.
    // ###########################################################################################
    public async Task LoadKiCadProjectForCurrentBoardAsync()
    {
        int loadVersion = unchecked(++this.thisKiCadProjectLoadVersion);
        string expectedBoardKey = this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;

        this.thisKiCadProject = null;
        this.thisCurrentKiCadRuntimeCacheScopeKey = this.BuildCurrentKiCadRuntimeCacheScopeKey();
        this.thisKiCadSchematicHoverHitTestCacheByKey.Clear();

        lock (this.thisKiCadPcbNetRenderCacheSync)
        {
            this.thisKiCadPcbNetRenderBuildTaskByKey.Clear();
        }

        lock (this.thisKiCadPcbHoverHitTestCacheSync)
        {
            this.thisKiCadPcbHoverHitTestBuildTaskByKey.Clear();
        }

        this.thisSelectedKiCadReferences.Clear();
        this.thisSelectedKiCadNormalizedNetNames.Clear();
        this.thisLockedKiCadNetNames.Clear();
        this.thisHoveredKiCadNetName = null;
        this.thisHoveredKiCadPadNumber = null;
        this.thisIsKiCadOverlayRefreshQueued = false;
        this.thisKiCadOverlayRefreshRequestVersion = 0;
        this.thisKiCadOverlayLastRenderedVersion = 0;
        this.thisLastKiCadNetConnectionsSignature = string.Empty;
        this.thisLastThumbnailHighlightSignature = string.Empty;
        this.thisImportantSignalNetNamesByDisplayName.Clear();
        this.thisImportantSignalDisplayNames.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();
        this.ResetKiCadHoverHitTestThrottle();
        this.ClearKiCadOverlay();

        this.KiCadNetConnectionsPanel.IsVisible = false;
        this.KiCadNetConnectionsHeaderTextBlock.Text = "Netlist name";
        this.KiCadNetConnectionsList.ItemsSource = null;
        this.UpdateKiCadNetConnectionsClearButtonState(false);

        this.UpdateImportantSignalsHeaderText();
        this.ImportantSignalsListBox.ItemsSource = null;
        this.ImportantSignalsListBox.SelectedItems?.Clear();
        this.ImportantSignalsPanel.IsVisible = false;
        this.UpdateImportantSignalsClearButtonState(false);

        this.RestoreBoardSettings(expectedBoardKey);

        var rawPaths = this.MainWindow?.GetCurrentBoardKiCadRawPaths() ?? new List<string>();
        if (rawPaths.Count == 0)
        {
            this.SetKiCadInitializingIndicatorVisible(false);
            this.thisCurrentKiCadRuntimeCacheScopeKey = string.Empty;
            this.RebuildImportantSignalsPanel();
            this.RestoreBoardSettings(expectedBoardKey);
            this.RefreshKiCadOverlay();
            return;
        }

        this.SetKiCadInitializingIndicatorVisible(true);

        var boardEntry = this.MainWindow?.GetCurrentBoardEntry();
        string hardwareName = boardEntry?.HardwareName ?? string.Empty;
        string boardName = boardEntry?.BoardName ?? string.Empty;

        KiCadProjectBundle? loadedProject = null;

        try
        {
            loadedProject = await Task.Run(async () =>
                await KiCadProjectLoader.LoadRawAsync(rawPaths, hardwareName, boardName));
        }
        catch (Exception ex)
        {
            Logger.Warning($"Failed to load KiCad project in background - [{ex}]");
        }

        if (loadVersion != this.thisKiCadProjectLoadVersion)
        {
            // A newer load is already running and owns the indicator, so leave it showing for that one.
            return;
        }

        string currentBoardKey = this.MainWindow?.GetCurrentBoardKey() ?? string.Empty;
        if (!string.Equals(expectedBoardKey, currentBoardKey, StringComparison.OrdinalIgnoreCase))
        {
            // The board moved on without starting another KiCad load, so nothing is initializing.
            this.SetKiCadInitializingIndicatorVisible(false);
            return;
        }

        this.thisKiCadProject = loadedProject;

        var activeScope = this.GetOrCreateCurrentKiCadRuntimeCacheScope();

        lock (this.thisKiCadPcbNetRenderCacheSync)
        {
            this.thisKiCadPcbNetRenderCacheByKey.Clear();
            this.thisKiCadPcbNetRenderBuildTaskByKey.Clear();

            if (activeScope != null)
            {
                foreach (var pair in activeScope.NetRenderCacheByKey)
                {
                    this.thisKiCadPcbNetRenderCacheByKey[pair.Key] = pair.Value;
                }

                foreach (var pair in activeScope.NetRenderBuildTaskByKey)
                {
                    this.thisKiCadPcbNetRenderBuildTaskByKey[pair.Key] = pair.Value;
                }
            }
        }

        lock (this.thisKiCadPcbHoverHitTestCacheSync)
        {
            this.thisKiCadPcbHoverHitTestCacheByKey.Clear();
            this.thisKiCadPcbHoverHitTestBuildTaskByKey.Clear();

            if (activeScope != null)
            {
                foreach (var pair in activeScope.HoverHitTestCacheByKey)
                {
                    this.thisKiCadPcbHoverHitTestCacheByKey[pair.Key] = pair.Value;
                }

                foreach (var pair in activeScope.HoverHitTestBuildTaskByKey)
                {
                    this.thisKiCadPcbHoverHitTestBuildTaskByKey[pair.Key] = pair.Value;
                }
            }
        }

        this.thisSelectedKiCadReferences.Clear();
        this.thisSelectedKiCadNormalizedNetNames.Clear();
        this.thisLockedKiCadNetNames.Clear();
        this.thisHoveredKiCadNetName = null;
        this.thisHoveredKiCadPadNumber = null;
        this.thisLastKiCadNetConnectionsSignature = string.Empty;
        this.thisLastThumbnailHighlightSignature = string.Empty;
        this.thisImportantSignalNetNamesByDisplayName.Clear();
        this.thisImportantSignalDisplayNames.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();
        this.ResetKiCadHoverHitTestThrottle();

        this.RebuildImportantSignalsPanel();
        this.RestoreBoardSettings(currentBoardKey);
        this.RefreshKiCadOverlay();

        this.SetKiCadInitializingIndicatorVisible(false);
    }

    // ###########################################################################################
    // Shows or hides the bottom-right "KiCad data initializing..." indicator.
    //
    // A board's KiCad project loads in the background so the schematic image can appear straight
    // away, which leaves a few seconds where traces, nets and pads do not respond yet. The
    // indicator is what tells the user that gap is data still arriving rather than a dead overlay.
    // ###########################################################################################
    public void SetKiCadInitializingIndicatorVisible(bool isVisible)
    {
        this.KiCadInitializingIndicator.IsVisible = isVisible;
    }

    // ###########################################################################################
    // Updates the active KiCad selection from the currently selected board-label references and
    // derives the corresponding normalized PCB net names from footprint pad assignments.
    // ###########################################################################################
    private void UpdateKiCadSelectionFromBoardLabels(IEnumerable<string> boardLabels)
    {
        this.thisSelectedKiCadReferences.Clear();
        this.thisSelectedKiCadNormalizedNetNames.Clear();

        foreach (string boardLabel in boardLabels
                     .Where(label => !string.IsNullOrWhiteSpace(label))
                     .Select(label => label.Trim()))
        {
            this.thisSelectedKiCadReferences.Add(boardLabel);
        }

        foreach (string netName in this.BuildKiCadNormalizedNetNamesForReferences(this.thisSelectedKiCadReferences))
        {
            this.thisSelectedKiCadNormalizedNetNames.Add(netName);
        }

        this.RefreshKiCadOverlay();
    }

    // ###########################################################################################
    // Derives normalized PCB net names from the selected footprint references by reading the net
    // assignments on the pads of matching footprints.
    // ###########################################################################################
    private HashSet<string> BuildKiCadNormalizedNetNamesForReferences(IEnumerable<string> references)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var root = this.thisKiCadProject?.Root;
        if (root == null || root.Pcb.Count == 0)
        {
            return result;
        }

        var referenceSet = new HashSet<string>(
            references.Where(reference => !string.IsNullOrWhiteSpace(reference)),
            StringComparer.OrdinalIgnoreCase);

        if (referenceSet.Count == 0)
        {
            return result;
        }

        var pcb = root.Pcb[0];

        foreach (var footprint in pcb.Footprints)
        {
            string reference = footprint.Reference?.Trim() ?? string.Empty;
            if (!referenceSet.Contains(reference))
            {
                continue;
            }

            foreach (var pad in footprint.Pads)
            {
                string normalizedName = pad.Net?.NormalizedName?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(normalizedName))
                {
                    result.Add(normalizedName);
                }
            }
        }

        return result;
    }

    // ###########################################################################################
    // Resolves the active Excel/image schematic selection to the corresponding KiCad project view.
    // Top and bottom replica pages are matched explicitly, while schematic pages are matched by
    // exact display name first and by page ordinal second.
    // ###########################################################################################
    private KiCadProjectView? ResolveKiCadViewForCurrentSchematic()
    {
        return this.ResolveKiCadViewForSchematicName(this.GetCurrentSchematicName());
    }

    // ###########################################################################################
    // Clears all explicit KiCad trace selections and any selected Important signals so shared nets
    // can always be cleared from the Netlist names panel regardless of where they were added from.
    // ###########################################################################################
    private void ClearAllKiCadTraceSelections()
    {
        this.thisLockedKiCadNetNames.Clear();
        this.thisHoveredKiCadNetName = null;
        this.thisHoveredKiCadPadNumber = null;

        this.ImportantSignalsListBox.SelectedItems?.Clear();
        this.thisSelectedImportantSignalDisplayNames.Clear();
        this.UpdateImportantSignalsHeaderText();
        this.UpdateImportantSignalsClearButtonState(this.ImportantSignalsPanel.IsVisible);

        this.SchematicsHoverPadBorder.IsVisible = false;
        this.SchematicsHoverPadText.Text = string.Empty;

        this.RefreshKiCadOverlay();
        this.RefreshBlinkStateFromCurrentSelection();
    }

    // ###########################################################################################
    // Returns true when the currently shown schematic has KiCad overlay data available.
    // ###########################################################################################
    private bool HasCurrentSchematicKiCadTraces()
    {
        if (this.thisKiCadProject?.Root?.Project == null)
        {
            return false;
        }

        var view = this.ResolveKiCadViewForCurrentSchematic();
        if (view == null)
        {
            return false;
        }

        if (string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
        {
            return view.SourceIndex >= 0 &&
                   view.SourceIndex < this.thisKiCadProject.Root.Pcb.Count;
        }

        if (string.Equals(view.Type, "schematic", StringComparison.OrdinalIgnoreCase))
        {
            return view.SourceIndex >= 0 &&
                   view.SourceIndex < this.thisKiCadProject.Root.Schematics.Count;
        }

        return false;
    }

    // ###########################################################################################
    // Returns true when the current schematic is a KiCad-backed PCB replica view with pad data.
    // Pin-1 marking is only available there because the current KiCad JSON exposes pad geometry.
    // ###########################################################################################
    private bool HasCurrentSchematicKiCadPcbPadData()
    {
        if (this.thisKiCadProject?.Root?.Project == null)
        {
            return false;
        }

        var view = this.ResolveKiCadViewForCurrentSchematic();
        if (view == null)
        {
            return false;
        }

        if (!string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return view.SourceIndex >= 0 &&
               view.SourceIndex < this.thisKiCadProject.Root.Pcb.Count;
    }

    // ###########################################################################################
    // Builds the effective KiCad net-name set used for trace preview in the current schematic.
    // Includes selected-component nets only when the global select option is enabled, and includes
    // hovered-component nets when selected-component trace preview is enabled.
    // Also includes manually selected important-signal net groups.
    // ###########################################################################################
    private HashSet<string> BuildActiveKiCadTracePreviewNetNames()
    {
        var activeNets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (UserSettings.SchematicsShowTracesOnComponentSelect)
        {
            foreach (string netName in this.thisSelectedKiCadNormalizedNetNames)
            {
                activeNets.Add(netName);
            }
        }

        if (this.IsBoardShowTracesOnSelectedComponentEnabled() &&
            !string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel))
        {
            foreach (string netName in this.BuildKiCadNormalizedNetNamesForReferences(
                new[] { this.thisHoveredComponentBoardLabel.Trim() }))
            {
                activeNets.Add(netName);
            }
        }

        foreach (string netName in this.BuildSelectedImportantSignalNetNames())
        {
            activeNets.Add(netName);
        }

        foreach (string lockedNet in this.thisLockedKiCadNetNames)
        {
            activeNets.Add(lockedNet);
        }

        string? activeHoveredKiCadNetName = this.GetActiveHoveredKiCadNetName();
        if (!string.IsNullOrWhiteSpace(activeHoveredKiCadNetName))
        {
            activeNets.Add(activeHoveredKiCadNetName);
        }

        return activeNets;
    }

    // ###########################################################################################
    // Builds the effective KiCad reference set used for PCB trace preview traversal.
    // Includes the hovered component when selected-component trace preview is enabled.
    // ###########################################################################################
    private HashSet<string> BuildActiveKiCadTracePreviewReferences()
    {
        var activeReferences = new HashSet<string>(this.thisSelectedKiCadReferences, StringComparer.OrdinalIgnoreCase);

        if (this.IsBoardShowTracesOnSelectedComponentEnabled() &&
            !string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel))
        {
            activeReferences.Add(this.thisHoveredComponentBoardLabel.Trim());
        }

        return activeReferences;
    }

    // ###########################################################################################
    // Builds the normalized KiCad net-name set for one optional component reference.
    // Returns an empty set when no valid board label is supplied.
    // ###########################################################################################
    private HashSet<string> BuildKiCadNormalizedNetNamesForSingleReference(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
        {
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }

        return this.BuildKiCadNormalizedNetNamesForReferences(new[] { reference.Trim() });
    }

    // ###########################################################################################
    // Builds a stable runtime cache scope key for the current board's raw KiCad files.
    // The key matches the raw loader identity by including file paths and last-write timestamps.
    // ###########################################################################################
    private string BuildCurrentKiCadRuntimeCacheScopeKey()
    {
        var rawPaths = this.MainWindow?.GetCurrentBoardKiCadRawPaths() ?? new List<string>();

        var existingPaths = rawPaths
            .Where(path => !string.IsNullOrWhiteSpace(path))
            .Select(System.IO.Path.GetFullPath)
            .Where(System.IO.File.Exists)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (existingPaths.Count == 0)
        {
            return string.Empty;
        }

        return string.Join(
            "\u001E",
            existingPaths.Select(path =>
                $"{path}|{System.IO.File.GetLastWriteTimeUtc(path).Ticks.ToString(CultureInfo.InvariantCulture)}"));
    }

    // ###########################################################################################
    // Returns the active KiCad runtime cache scope for the current board and updates the LRU order.
    // Creates a new scope when this KiCad project has not been seen before in the current session.
    // ###########################################################################################
    private KiCadRuntimeCacheScope? GetOrCreateCurrentKiCadRuntimeCacheScope()
    {
        string scopeKey = this.BuildCurrentKiCadRuntimeCacheScopeKey();
        this.thisCurrentKiCadRuntimeCacheScopeKey = scopeKey;

        if (string.IsNullOrWhiteSpace(scopeKey))
        {
            return null;
        }

        if (!this.thisKiCadRuntimeCacheScopeByKey.TryGetValue(scopeKey, out var scope))
        {
            scope = new KiCadRuntimeCacheScope();
            this.thisKiCadRuntimeCacheScopeByKey[scopeKey] = scope;
        }

        this.TouchKiCadRuntimeCacheScope(scopeKey);
        this.TrimKiCadRuntimeCacheScopes();

        return scope;
    }

    // ###########################################################################################
    // Marks one runtime cache scope as recently used so older boards can be evicted first.
    // ###########################################################################################
    private void TouchKiCadRuntimeCacheScope(string scopeKey)
    {
        if (string.IsNullOrWhiteSpace(scopeKey))
        {
            return;
        }

        var node = this.thisKiCadRuntimeCacheScopeLru.First;
        while (node != null)
        {
            var next = node.Next;

            if (string.Equals(node.Value, scopeKey, StringComparison.OrdinalIgnoreCase))
            {
                this.thisKiCadRuntimeCacheScopeLru.Remove(node);
                break;
            }

            node = next;
        }

        this.thisKiCadRuntimeCacheScopeLru.AddFirst(scopeKey);
    }

    // ###########################################################################################
    // Limits the number of retained KiCad runtime cache scopes so switching across many boards
    // does not grow memory usage without bound during one application session.
    // ###########################################################################################
    private void TrimKiCadRuntimeCacheScopes()
    {
        const int maximumRetainedScopes = 4;

        while (this.thisKiCadRuntimeCacheScopeLru.Count > maximumRetainedScopes)
        {
            var lastNode = this.thisKiCadRuntimeCacheScopeLru.Last;
            if (lastNode == null)
            {
                break;
            }

            string scopeKey = lastNode.Value;
            this.thisKiCadRuntimeCacheScopeLru.RemoveLast();

            if (string.Equals(scopeKey, this.thisCurrentKiCadRuntimeCacheScopeKey, StringComparison.OrdinalIgnoreCase))
            {
                this.thisKiCadRuntimeCacheScopeLru.AddFirst(scopeKey);
                break;
            }

            this.thisKiCadRuntimeCacheScopeByKey.Remove(scopeKey);
        }
    }

    // ###########################################################################################
    // Clears explicit KiCad trace selections that match the supplied normalized net names.
    // This allows overlapping Important signal and manual KiCad net selections to clear each other.
    // ###########################################################################################
    private void ClearKiCadTraceSelectionsForNetNames(IReadOnlyCollection<string> normalizedNetNames)
    {
        if (normalizedNetNames == null || normalizedNetNames.Count == 0)
        {
            return;
        }

        var targetNetNames = new HashSet<string>(
            normalizedNetNames
                .Where(netName => !string.IsNullOrWhiteSpace(netName))
                .Select(netName => netName.Trim()),
            StringComparer.OrdinalIgnoreCase);

        if (targetNetNames.Count == 0)
        {
            return;
        }

        this.thisLockedKiCadNetNames.RemoveWhere(netName => targetNetNames.Contains(netName));

        string hoveredNetName = this.thisHoveredKiCadNetName?.Trim() ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(hoveredNetName) &&
            targetNetNames.Contains(hoveredNetName))
        {
            this.thisHoveredKiCadNetName = null;
            this.thisHoveredKiCadPadNumber = null;
            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;
        }
    }
}