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
// Hover hit-testing over the KiCad overlay: the spatial hit-test caches, hover throttling,
// the hovered net/pad state, and the interactive trace hover mode UI.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private string? thisHoveredKiCadNetName;

    private string? thisHoveredKiCadPadNumber;

    private bool thisIsInteractiveCadTraceHoverShiftPressed;

    private readonly object thisKiCadPcbHoverHitTestCacheSync = new();

    private readonly Dictionary<string, Task> thisKiCadPcbHoverHitTestBuildTaskByKey = new(StringComparer.OrdinalIgnoreCase);

    private Point thisLastKiCadHoverHitTestContainerPoint = new(double.NaN, double.NaN);

    private long thisLastKiCadHoverHitTestTimestamp;

    private readonly Dictionary<string, KiCadPcbHoverHitTestCache> thisKiCadPcbHoverHitTestCacheByKey = new(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<string, KiCadSchematicHoverHitTestCache> thisKiCadSchematicHoverHitTestCacheByKey =
        new(StringComparer.OrdinalIgnoreCase);

    // ###########################################################################################
    // Triggers a visual overlay refresh when the hovered KiCad net changes.
    // ###########################################################################################
    private void SetHoveredKiCadNet(string? netName)
    {
        if (string.Equals(this.thisHoveredKiCadNetName, netName, StringComparison.OrdinalIgnoreCase))
            return;

        this.thisHoveredKiCadNetName = netName;

        // Immediate on purpose. This was briefly debounced while a hover rebuild cost ~262 ms, but
        // the per-net primitive cache brought that down to ~32 ms, and the delay was worse than the
        // cost it avoided: moving slowly across traces kept restarting the quiet period, so the
        // highlight only appeared once the pointer stopped and traces in between were skipped.
        this.RefreshKiCadOverlay(forceImmediate: true);
    }

    // ###########################################################################################
    // Enumerates all schematic label types that can identify one net on a schematic page.
    // ###########################################################################################
    private static IEnumerable<KiCadSchematicLabel> EnumerateKiCadSchematicNetLabels(KiCadSchematic schematic)
    {
        foreach (var label in schematic.Labels.Local)
        {
            yield return label;
        }

        foreach (var label in schematic.Labels.Global)
        {
            yield return label;
        }

        foreach (var label in schematic.Labels.Hierarchical)
        {
            yield return label;
        }
    }

    // ###########################################################################################
    // Hit-tests KiCad schematic nets using a cached local-space spatial index for labels and
    // resolved wire segments so hover remains responsive on dense schematic pages.
    // ###########################################################################################
    private void HitTestKiCadSchematicOverlayForHover(KiCadProjectView view, Point localPoint)
    {
        if (this.thisKiCadProject == null ||
            view.SourceIndex < 0 ||
            view.SourceIndex >= this.thisKiCadProject.Root.Schematics.Count)
        {
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;
            return;
        }

        var schematic = this.thisKiCadProject.Root.Schematics[view.SourceIndex];
        var contentRect = this.GetImageContentRect();
        var worldBounds = this.GetKiCadSchematicWorldBounds(schematic);

        if (contentRect.Width <= 0 ||
            contentRect.Height <= 0 ||
            worldBounds.Width <= 0 ||
            worldBounds.Height <= 0)
        {
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;
            return;
        }

        string currentSchematicName = this.GetCurrentSchematicName();
        var calibration = this.GetKiCadViewCalibration(currentSchematicName);

        var cache = this.GetOrCreateKiCadSchematicHoverHitTestCache(
            view,
            schematic,
            worldBounds,
            contentRect,
            calibration,
            currentSchematicName);

        if (cache == null)
        {
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;
            return;
        }

        double hitThresholdLocal = Math.Max(3.0, 8.0 / Math.Max(0.0001, this.schematicsMatrix.M11));

        int minCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.X - hitThresholdLocal, cache.CellSizeLocal);
        int maxCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.X + hitThresholdLocal, cache.CellSizeLocal);
        int minCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.Y - hitThresholdLocal, cache.CellSizeLocal);
        int maxCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.Y + hitThresholdLocal, cache.CellSizeLocal);

        string? bestNetName = null;
        double bestDistance = double.MaxValue;

        var testedLabelIndices = new HashSet<int>();
        var testedSegmentIndices = new HashSet<int>();

        for (int cellY = minCellY; cellY <= maxCellY; cellY++)
        {
            for (int cellX = minCellX; cellX <= maxCellX; cellX++)
            {
                long cellKey = KiCadHoverIndex.BuildKiCadHoverCellKey(cellX, cellY);

                if (cache.LabelIndicesByCell.TryGetValue(cellKey, out var labelIndices))
                {
                    foreach (int labelIndex in labelIndices)
                    {
                        if (!testedLabelIndices.Add(labelIndex))
                        {
                            continue;
                        }

                        var candidate = cache.LabelCandidates[labelIndex];
                        double dx = candidate.LocalPoint.X - localPoint.X;
                        double dy = candidate.LocalPoint.Y - localPoint.Y;
                        double distance = Math.Sqrt((dx * dx) + (dy * dy));

                        if (distance <= hitThresholdLocal && distance < bestDistance)
                        {
                            bestDistance = distance;
                            bestNetName = candidate.NormalizedNetName;
                        }
                    }
                }

                if (cache.SegmentIndicesByCell.TryGetValue(cellKey, out var segmentIndices))
                {
                    foreach (int segmentIndex in segmentIndices)
                    {
                        if (!testedSegmentIndices.Add(segmentIndex))
                        {
                            continue;
                        }

                        var candidate = cache.SegmentCandidates[segmentIndex];

                        double distance = PolygonGeometry.DistanceToSegment(
                            localPoint,
                            candidate.StartLocal.X,
                            candidate.StartLocal.Y,
                            candidate.EndLocal.X,
                            candidate.EndLocal.Y);

                        if (distance <= hitThresholdLocal && distance < bestDistance)
                        {
                            bestDistance = distance;
                            bestNetName = candidate.NormalizedNetName;
                        }
                    }
                }
            }
        }

        this.SetHoveredKiCadNet(string.IsNullOrWhiteSpace(bestNetName) ? null : bestNetName);
        this.thisHoveredKiCadPadNumber = null;
    }

    // ###########################################################################################
    // Performs KiCad hover hit-testing for both PCB and schematic views.
    // PCB pages use the existing spatial cache, while schematic pages resolve the nearest net
    // by checking net-label anchors and rendered wire/polyline paths in the active sheet.
    // ###########################################################################################
    private void HitTestKiCadOverlayForHover(Point pointerInContainer)
    {
        var view = this.ResolveKiCadViewForCurrentSchematic();
        if (view == null || this.thisKiCadProject == null || this.currentFullResBitmap == null)
        {
            this.SetHoveredKiCadNet(null);
            return;
        }

        if (!RectGeometry.TryInvert(this.schematicsMatrix, out var inv))
        {
            this.SetHoveredKiCadNet(null);
            return;
        }

        var localPoint = new Point(
            (pointerInContainer.X * inv.M11) + (pointerInContainer.Y * inv.M21) + inv.M31,
            (pointerInContainer.X * inv.M12) + (pointerInContainer.Y * inv.M22) + inv.M32);

        bool isTop = string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase);
        bool isBottom = string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase);
        bool isSchematic = string.Equals(view.Type, "schematic", StringComparison.OrdinalIgnoreCase);

        if (isSchematic)
        {
            this.HitTestKiCadSchematicOverlayForHover(view, localPoint);
            return;
        }

        if (!isTop && !isBottom)
        {
            this.SetHoveredKiCadNet(null);
            return;
        }

        string requiredLayer = isBottom ? "B.Cu" : "F.Cu";
        var pcb = this.thisKiCadProject.Root.Pcb.ElementAtOrDefault(view.SourceIndex);
        if (pcb == null)
        {
            this.SetHoveredKiCadNet(null);
            return;
        }

        var contentRect = this.GetImageContentRect();
        var worldBounds = this.GetKiCadPcbWorldBounds(pcb);
        string currentSchematicName = this.GetCurrentSchematicName();
        var calibration = this.GetKiCadViewCalibration(currentSchematicName);

        if (!this.TryMapLocalToKiCadWorld(localPoint, worldBounds, contentRect, calibration, out var worldPoint))
        {
            this.SetHoveredKiCadNet(null);
            return;
        }

        var cache = this.GetOrCreateKiCadPcbHoverHitTestCache(pcb, view.SourceIndex, requiredLayer);
        if (cache == null)
        {
            this.SetHoveredKiCadNet(null);
            this.thisHoveredKiCadPadNumber = null;
            return;
        }

        const double zoneHoverToleranceWorld = 0.4;
        double searchRadiusWorld = Math.Max(0.8, Math.Max(cache.MaxHitRadiusWorld, zoneHoverToleranceWorld));

        int minCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(worldPoint.X - searchRadiusWorld, cache.CellSizeWorld);
        int maxCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(worldPoint.X + searchRadiusWorld, cache.CellSizeWorld);
        int minCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(worldPoint.Y - searchRadiusWorld, cache.CellSizeWorld);
        int maxCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(worldPoint.Y + searchRadiusWorld, cache.CellSizeWorld);

        var testedPadIndices = new HashSet<int>();
        var testedZoneIndices = new HashSet<int>();
        var testedSegmentIndices = new HashSet<int>();
        var testedViaIndices = new HashSet<int>();

        double closestPadDist = double.MaxValue;
        KiCadNetRef? bestPadNet = null;
        string? bestPadNumber = null;

        for (int cellY = minCellY; cellY <= maxCellY; cellY++)
        {
            for (int cellX = minCellX; cellX <= maxCellX; cellX++)
            {
                long cellKey = KiCadHoverIndex.BuildKiCadHoverCellKey(cellX, cellY);

                if (!cache.PadIndicesByCell.TryGetValue(cellKey, out var padIndices))
                {
                    continue;
                }

                foreach (int padIndex in padIndices)
                {
                    if (!testedPadIndices.Add(padIndex))
                    {
                        continue;
                    }

                    var candidate = cache.PadCandidates[padIndex];

                    double dx = candidate.CenterWorld.X - worldPoint.X;
                    double dy = candidate.CenterWorld.Y - worldPoint.Y;
                    double dist = Math.Sqrt((dx * dx) + (dy * dy));

                    if (dist < candidate.HitRadiusWorld && dist < closestPadDist)
                    {
                        closestPadDist = dist;
                        bestPadNet = candidate.Net;
                        bestPadNumber = candidate.PadNumber;
                    }
                }
            }
        }

        if (bestPadNet != null)
        {
            string? foundPadNet = bestPadNet.NormalizedName?.Trim();
            this.SetHoveredKiCadNet(string.IsNullOrWhiteSpace(foundPadNet) ? null : foundPadNet);
            this.thisHoveredKiCadPadNumber = bestPadNumber?.Trim();
            return;
        }

        double closestZoneDist = double.MaxValue;
        KiCadNetRef? bestZoneNet = null;

        for (int cellY = minCellY; cellY <= maxCellY; cellY++)
        {
            for (int cellX = minCellX; cellX <= maxCellX; cellX++)
            {
                long cellKey = KiCadHoverIndex.BuildKiCadHoverCellKey(cellX, cellY);

                if (!cache.ZoneIndicesByCell.TryGetValue(cellKey, out var zoneIndices))
                {
                    continue;
                }

                foreach (int zoneIndex in zoneIndices)
                {
                    if (!testedZoneIndices.Add(zoneIndex))
                    {
                        continue;
                    }

                    var candidate = cache.ZoneCandidates[zoneIndex];

                    if (!PolygonGeometry.IsPointInOrNearZone(
                            worldPoint,
                            candidate.PolygonsWorld,
                            zoneHoverToleranceWorld,
                            out double zoneDistanceWorld))
                    {
                        continue;
                    }

                    if (zoneDistanceWorld < closestZoneDist)
                    {
                        closestZoneDist = zoneDistanceWorld;
                        bestZoneNet = candidate.Net;
                    }
                }
            }
        }

        if (bestZoneNet != null)
        {
            string? foundZoneNet = bestZoneNet.NormalizedName?.Trim();
            this.SetHoveredKiCadNet(string.IsNullOrWhiteSpace(foundZoneNet) ? null : foundZoneNet);
            this.thisHoveredKiCadPadNumber = null;
            return;
        }

        double closestDist = double.MaxValue;
        KiCadNetRef? bestNet = null;

        for (int cellY = minCellY; cellY <= maxCellY; cellY++)
        {
            for (int cellX = minCellX; cellX <= maxCellX; cellX++)
            {
                long cellKey = KiCadHoverIndex.BuildKiCadHoverCellKey(cellX, cellY);

                if (cache.SegmentIndicesByCell.TryGetValue(cellKey, out var segmentIndices))
                {
                    foreach (int segmentIndex in segmentIndices)
                    {
                        if (!testedSegmentIndices.Add(segmentIndex))
                        {
                            continue;
                        }

                        var candidate = cache.SegmentCandidates[segmentIndex];

                        double dist = PolygonGeometry.DistanceToSegment(
                            worldPoint,
                            candidate.StartWorld.X,
                            candidate.StartWorld.Y,
                            candidate.EndWorld.X,
                            candidate.EndWorld.Y);

                        if (dist < closestDist && dist < candidate.HitRadiusWorld)
                        {
                            closestDist = dist;
                            bestNet = candidate.Net;
                        }
                    }
                }

                if (cache.ViaIndicesByCell.TryGetValue(cellKey, out var viaIndices))
                {
                    foreach (int viaIndex in viaIndices)
                    {
                        if (!testedViaIndices.Add(viaIndex))
                        {
                            continue;
                        }

                        var candidate = cache.ViaCandidates[viaIndex];

                        double dx = candidate.CenterWorld.X - worldPoint.X;
                        double dy = candidate.CenterWorld.Y - worldPoint.Y;
                        double dist = Math.Sqrt((dx * dx) + (dy * dy));

                        if (dist < closestDist && dist < candidate.HitRadiusWorld)
                        {
                            closestDist = dist;
                            bestNet = candidate.Net;
                        }
                    }
                }
            }
        }

        string? foundNet = bestNet?.NormalizedName?.Trim();
        this.SetHoveredKiCadNet(string.IsNullOrWhiteSpace(foundNet) ? null : foundNet);
        this.thisHoveredKiCadPadNumber = null;
    }

    // ###########################################################################################
    // Updates cached SHIFT state for interactive KiCad hover highlighting.
    // ###########################################################################################
    private void UpdateInteractiveCadTraceHoverShiftState(KeyModifiers modifiers)
    {
        bool isShiftPressed = modifiers.HasFlag(KeyModifiers.Shift);

        if (this.thisIsInteractiveCadTraceHoverShiftPressed == isShiftPressed)
        {
            return;
        }

        this.thisIsInteractiveCadTraceHoverShiftPressed = isShiftPressed;
        this.RefreshKiCadHoverPadUi();
        this.RefreshKiCadOverlay();
    }

    // ###########################################################################################
    // Refreshes the transient KiCad pad hover label based on the active hover mode.
    // ###########################################################################################
    private void RefreshKiCadHoverPadUi()
    {
        string hoveredPadNumber = this.GetActiveHoveredKiCadPadNumber()?.Trim() ?? string.Empty;
        string hoveredNetName = this.GetActiveHoveredKiCadNetName()?.Trim() ?? string.Empty;

        if (!string.IsNullOrWhiteSpace(hoveredPadNumber))
        {
            this.SchematicsHoverPadText.Text = string.IsNullOrWhiteSpace(hoveredNetName)
                ? hoveredPadNumber
                : $"{hoveredPadNumber} | {hoveredNetName}";
            this.SchematicsHoverPadBorder.IsVisible = true;
        }
        else
        {
            this.SchematicsHoverPadBorder.IsVisible = false;
            this.SchematicsHoverPadText.Text = string.Empty;
        }
    }

    // ###########################################################################################
    // Reacts to global interactive CAD trace hover mode changes from configuration.
    // ###########################################################################################
    private void OnInteractiveCadTraceHoverModeChanged()
    {
        Dispatcher.UIThread.Post(() =>
        {
            this.UpdateGlobalSettingsControls();
            this.UpdateInteractiveCadTraceHoverModeUi();
            this.RefreshKiCadHoverPadUi();
            this.RefreshKiCadOverlay();
        });
    }

    // ###########################################################################################
    // Updates schematics settings visibility for global and board-specific CAD trace options.
    // Also exposes the temporary calibration-only traces-and-pads toggle while KiCad calibration
    // mode is active so alignment can be checked against the underlying image.
    // ###########################################################################################
    private void UpdateInteractiveCadTraceHoverModeUi()
    {
        bool hasBoard = !string.IsNullOrWhiteSpace(this.MainWindow?.GetCurrentBoardKey());
        bool hasKiCadTraces = this.HasCurrentSchematicKiCadTraces();
        bool hasKiCadPcbPadData = this.HasCurrentSchematicKiCadPcbPadData();
        bool isHoverHighlightEnabled = this.CheckGlobalHoverHighlightsTraces.IsChecked == true;
        bool isHoldShiftEnabled = hasKiCadTraces && isHoverHighlightEnabled;

        var currentView = this.ResolveKiCadViewForCurrentSchematic();
        bool isCurrentViewPcb =
            currentView != null &&
            (string.Equals(currentView.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(currentView.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase));

        bool isCalibrationTraceToggleVisible =
            hasKiCadTraces &&
            this.thisIsKiCadTraceCalibrationMode;

        this.BoardMarkPin1OnSelectedComponentRow.IsVisible = hasBoard && hasKiCadPcbPadData;
        this.CheckBoardMarkPin1OnSelectedComponent.IsEnabled = hasBoard && hasKiCadPcbPadData;

        this.GlobalHoverHighlightsTracesRow.IsVisible = hasKiCadTraces;
        this.CheckGlobalHoverHighlightsTraces.IsEnabled = hasKiCadTraces;

        this.GlobalShowTracesOnComponentSelectRow.IsVisible = hasKiCadTraces;
        this.GlobalShowTracesOnComponentSelectRow.IsEnabled = hasKiCadTraces;
        this.GlobalShowTracesOnComponentSelectRow.Opacity = hasKiCadTraces ? 1.0 : 0.55;
        this.GlobalShowTracesOnComponentSelectRow.Cursor = hasKiCadTraces
            ? new Cursor(StandardCursorType.Hand)
            : Cursor.Default;
        this.CheckGlobalShowTracesOnComponentSelect.IsEnabled = hasKiCadTraces;

        this.GlobalShowOppositeSideTracesRow.IsVisible = isCurrentViewPcb;
        this.GlobalShowOppositeSideTracesRow.IsEnabled = isCurrentViewPcb;
        this.GlobalShowOppositeSideTracesRow.Opacity = isCurrentViewPcb ? 1.0 : 0.55;
        this.GlobalShowOppositeSideTracesRow.Cursor = isCurrentViewPcb
            ? new Cursor(StandardCursorType.Hand)
            : Cursor.Default;
        this.CheckGlobalShowOppositeSideTraces.IsEnabled = isCurrentViewPcb;

        this.GlobalShowCalibrationTracesAndPadsRow.IsVisible = isCalibrationTraceToggleVisible;
        this.GlobalShowCalibrationTracesAndPadsRow.IsEnabled = isCalibrationTraceToggleVisible;
        this.GlobalShowCalibrationTracesAndPadsRow.Opacity = isCalibrationTraceToggleVisible ? 1.0 : 0.55;
        this.GlobalShowCalibrationTracesAndPadsRow.Cursor = isCalibrationTraceToggleVisible
            ? new Cursor(StandardCursorType.Hand)
            : Cursor.Default;
        this.CheckGlobalShowCalibrationTracesAndPads.IsEnabled = isCalibrationTraceToggleVisible;

        this.GlobalShowZonesRow.IsVisible = isCurrentViewPcb;
        this.GlobalShowZonesRow.IsEnabled = isCurrentViewPcb;
        this.GlobalShowZonesRow.Opacity = isCurrentViewPcb ? 1.0 : 0.55;
        this.GlobalShowZonesRow.Cursor = isCurrentViewPcb
            ? new Cursor(StandardCursorType.Hand)
            : Cursor.Default;
        this.CheckGlobalShowZones.IsEnabled = isCurrentViewPcb;

        this.SchematicsInteractiveCadTraceHoverHoldShiftRow.IsVisible = hasKiCadTraces;

        this.BoardShowTracesOnSelectedComponentRow.IsVisible = hasKiCadTraces;
        this.CheckBoardShowTracesOnSelectedComponent.IsEnabled = hasKiCadTraces;

        this.UpdateSchematicsInteractiveCadTraceHoverHoldShiftVisualState(isHoldShiftEnabled);
    }

    // ###########################################################################################
    // Returns the active hovered KiCad net name, honoring the current hover mode settings.
    // ###########################################################################################
    private string? GetActiveHoveredKiCadNetName()
    {
        return this.IsBoardHoverHighlightsTracesEnabled()
            ? this.thisHoveredKiCadNetName
            : null;
    }

    // ###########################################################################################
    // Returns the active hovered KiCad pad number, honoring the current hover mode settings.
    // ###########################################################################################
    private string? GetActiveHoveredKiCadPadNumber()
    {
        return this.IsBoardHoverHighlightsTracesEnabled()
            ? this.thisHoveredKiCadPadNumber
            : null;
    }

    // ###########################################################################################
    // Resets the lightweight throttle state used for KiCad hover hit-testing.
    // ###########################################################################################
    private void ResetKiCadHoverHitTestThrottle()
    {
        this.thisLastKiCadHoverHitTestContainerPoint = new Point(double.NaN, double.NaN);
        this.thisLastKiCadHoverHitTestTimestamp = 0;
    }

    // ###########################################################################################
    // Drops the SCHEMATIC KiCad hover hit-test caches, then resets the hover throttle so the next
    // pointer move rebuilds from scratch.
    //
    // The schematic cache is keyed by schematicName/schematicIndex - never by the active
    // calibration - because it bakes the CURRENT calibration's LOCAL points directly into the
    // cached candidates (GetOrCreateKiCadSchematicHoverHitTestCache maps every one of them through
    // MapKiCadWorldToLocal). That is fine while calibration never changes under a cached entry, but
    // calibrating (dragging the box, then Apply) changes exactly that without touching the key, so
    // a cache built from a hover during the drag - or from the pre-calibration persisted values -
    // keeps being served after Apply, even though RefreshKiCadOverlay draws correctly because it
    // recomputes on every call instead of caching. Reported as: hovering with SHIFT (or the
    // always-on hover mode) missed traces after calibrating a schematic for the first time, until
    // the app was restarted (which clears these dictionaries via
    // LoadKiCadProjectForCurrentBoardAsync anyway).
    //
    // The PCB caches are deliberately NOT cleared here, although an earlier version of this method
    // did clear them. KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, requiredLayer) takes no
    // calibration and stores WORLD-space geometry only - the calibration is applied per hit-test,
    // after the cache is consulted - so nothing in it can go stale when the calibration changes.
    // Clearing it therefore bought nothing and cost a full asynchronous rebuild of a dense board's
    // index three times per calibration session (mode entry, and again on cancel or apply).
    //
    // It was also unsound: clearing thisKiCadPcbHoverHitTestBuildTaskByKey destroys the de-dup
    // guard without being able to cancel the task it was guarding, so a second concurrent build of
    // the same index starts, and whichever task finishes first removes the registry entry the other
    // one now owns.
    //
    // Called when calibration mode starts, is cancelled, and is applied - any of the three can
    // leave a schematic cache behind that was built against a calibration different from what is
    // active afterwards.
    // ###########################################################################################
    private void InvalidateKiCadHoverHitTestCachesForCalibrationChange()
    {
        this.thisKiCadSchematicHoverHitTestCacheByKey.Clear();

        this.ResetKiCadHoverHitTestThrottle();
    }

    // ###########################################################################################
    // Limits how often expensive KiCad hover hit-tests can run while the pointer is moving.
    // This keeps dense PCB overlays responsive during fast pan and pointer motion.
    // ###########################################################################################
    private bool ShouldProcessKiCadHoverHitTest(Point pointerInContainer)
    {
        const double minimumDistance = 3.0;
        const double minimumIntervalMilliseconds = 16.0;

        long now = Stopwatch.GetTimestamp();

        if (double.IsNaN(this.thisLastKiCadHoverHitTestContainerPoint.X) ||
            double.IsNaN(this.thisLastKiCadHoverHitTestContainerPoint.Y))
        {
            this.thisLastKiCadHoverHitTestContainerPoint = pointerInContainer;
            this.thisLastKiCadHoverHitTestTimestamp = now;
            return true;
        }

        double dx = pointerInContainer.X - this.thisLastKiCadHoverHitTestContainerPoint.X;
        double dy = pointerInContainer.Y - this.thisLastKiCadHoverHitTestContainerPoint.Y;
        double distance = Math.Sqrt((dx * dx) + (dy * dy));

        double elapsedMilliseconds =
            this.thisLastKiCadHoverHitTestTimestamp == 0
                ? double.MaxValue
                : ((now - this.thisLastKiCadHoverHitTestTimestamp) * 1000.0) / Stopwatch.Frequency;

        if (distance < minimumDistance && elapsedMilliseconds < minimumIntervalMilliseconds)
        {
            return false;
        }

        this.thisLastKiCadHoverHitTestContainerPoint = pointerInContainer;
        this.thisLastKiCadHoverHitTestTimestamp = now;
        return true;
    }

    // ###########################################################################################
    // Returns the cached PCB hover-hit-test data for the requested board side.
    // The cache is stored both in the current working dictionaries and in the active persistent
    // per-board runtime cache scope so revisiting the same board can reuse hover preparation work.
    // ###########################################################################################
    private KiCadPcbHoverHitTestCache? GetOrCreateKiCadPcbHoverHitTestCache(
        KiCadPcb pcb,
        int pcbIndex,
        string requiredLayer)
    {
        string cacheKey = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCacheKey(pcbIndex, requiredLayer);
        KiCadProjectBundle? expectedProject = this.thisKiCadProject;
        string expectedScopeKey = this.thisCurrentKiCadRuntimeCacheScopeKey;
        var activeScope = this.GetOrCreateCurrentKiCadRuntimeCacheScope();

        lock (this.thisKiCadPcbHoverHitTestCacheSync)
        {
            if (this.thisKiCadPcbHoverHitTestCacheByKey.TryGetValue(cacheKey, out var cache))
            {
                return cache;
            }

            if (activeScope != null &&
                activeScope.HoverHitTestCacheByKey.TryGetValue(cacheKey, out var scopedCache))
            {
                this.thisKiCadPcbHoverHitTestCacheByKey[cacheKey] = scopedCache;
                return scopedCache;
            }

            if (this.thisKiCadPcbHoverHitTestBuildTaskByKey.ContainsKey(cacheKey) ||
                (activeScope != null && activeScope.HoverHitTestBuildTaskByKey.ContainsKey(cacheKey)))
            {
                return null;
            }

            Task buildTask = Task.Run(() =>
            {
                try
                {
                    var builtCache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, requiredLayer);

                    lock (this.thisKiCadPcbHoverHitTestCacheSync)
                    {
                        if (!ReferenceEquals(expectedProject, this.thisKiCadProject) ||
                            !string.Equals(expectedScopeKey, this.thisCurrentKiCadRuntimeCacheScopeKey, StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }

                        this.thisKiCadPcbHoverHitTestCacheByKey[cacheKey] = builtCache;

                        if (activeScope != null)
                        {
                            activeScope.HoverHitTestCacheByKey[cacheKey] = builtCache;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to build KiCad PCB hover cache [{cacheKey}] - [{ex.Message}]");
                }
                finally
                {
                    lock (this.thisKiCadPcbHoverHitTestCacheSync)
                    {
                        this.thisKiCadPcbHoverHitTestBuildTaskByKey.Remove(cacheKey);

                        if (activeScope != null)
                        {
                            activeScope.HoverHitTestBuildTaskByKey.Remove(cacheKey);
                        }
                    }

                    if (ReferenceEquals(expectedProject, this.thisKiCadProject) &&
                        string.Equals(expectedScopeKey, this.thisCurrentKiCadRuntimeCacheScopeKey, StringComparison.OrdinalIgnoreCase))
                    {
                        Dispatcher.UIThread.Post(() =>
                        {
                            if (double.IsNaN(this.thisLastKiCadHoverHitTestContainerPoint.X) ||
                                double.IsNaN(this.thisLastKiCadHoverHitTestContainerPoint.Y))
                            {
                                return;
                            }

                            this.HitTestKiCadOverlayForHover(this.thisLastKiCadHoverHitTestContainerPoint);
                            this.RefreshKiCadHoverPadUi();
                        }, DispatcherPriority.Background);
                    }
                }
            });

            this.thisKiCadPcbHoverHitTestBuildTaskByKey[cacheKey] = buildTask;

            if (activeScope != null)
            {
                activeScope.HoverHitTestBuildTaskByKey[cacheKey] = buildTask;
            }

            return null;
        }
    }

    // ###########################################################################################
    // Persists the copied interactive CAD trace hover option from the schematics global settings panel.
    // ###########################################################################################
    private void OnSchematicsInteractiveCadTraceHoverModeChanged(object? sender, RoutedEventArgs e)
    {
        this.ApplyInteractiveCadTraceHoverModeFromGlobalSettings();
    }

    // ###########################################################################################
    // Handles row clicks for the "Hold SHIFT to highlight traces on hover" option.
    // ###########################################################################################
    private void OnSchematicsInteractiveCadTraceHoverHoldShiftRowClicked(object? sender, PointerPressedEventArgs e)
    {
        if (!this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsEnabled)
        {
            return;
        }

        if (e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked =
                this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked != true;

            e.Handled = true;
        }
    }

    // ###########################################################################################
    // Persists the global hover-trace behavior from the schematics global settings panel.
    // ###########################################################################################
    private void ApplyInteractiveCadTraceHoverModeFromGlobalSettings()
    {
        if (this.thisSuppressGlobalSettingsChanged)
        {
            return;
        }

        bool isEnabled = this.CheckGlobalHoverHighlightsTraces.IsChecked == true;
        bool requiresShift =
            isEnabled &&
            this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsChecked == true;

        this.UpdateSchematicsInteractiveCadTraceHoverHoldShiftVisualState(isEnabled);

        UserSettings.InteractiveCadTraceHoverMode = !isEnabled
            ? "Disabled"
            : requiresShift
                ? "HoldShift"
                : "Always";
    }

    // ###########################################################################################
    // Applies enabled state, cursor, and dimmed appearance to the SHIFT hover option row.
    // ###########################################################################################
    private void UpdateSchematicsInteractiveCadTraceHoverHoldShiftVisualState(bool isEnabled)
    {
        this.SchematicsInteractiveCadTraceHoverHoldShiftCheckBox.IsEnabled = isEnabled;
        this.SchematicsInteractiveCadTraceHoverHoldShiftRow.IsEnabled = isEnabled;
        this.SchematicsInteractiveCadTraceHoverHoldShiftRow.Opacity = isEnabled ? 1.0 : 0.55;
        this.SchematicsInteractiveCadTraceHoverHoldShiftRow.Cursor = isEnabled
            ? new Cursor(StandardCursorType.Hand)
            : Cursor.Default;
    }

    // ###########################################################################################
    // Builds a stable cache key for schematic hover hit-testing on one schematic image.
    // ###########################################################################################
    private static string BuildKiCadSchematicHoverHitTestCacheKey(string schematicName, int schematicIndex)
    {
        return string.Join(
            "\u001F",
            schematicIndex.ToString(CultureInfo.InvariantCulture),
            schematicName?.Trim() ?? string.Empty);
    }

    // ###########################################################################################
    // Returns the cached schematic hover-hit-test data for the current schematic view.
    // ###########################################################################################
    private KiCadSchematicHoverHitTestCache? GetOrCreateKiCadSchematicHoverHitTestCache(
        KiCadProjectView view,
        KiCadSchematic schematic,
        Rect worldBounds,
        Rect contentRect,
        KiCadViewCalibration calibration,
        string schematicName)
    {
        string cacheKey = TabSchematics.BuildKiCadSchematicHoverHitTestCacheKey(schematicName, view.SourceIndex);

        if (this.thisKiCadSchematicHoverHitTestCacheByKey.TryGetValue(cacheKey, out var cache))
        {
            return cache;
        }

        if (!this.thisKiCadProject!.SchematicNetPathIndexBySchematicIndex.TryGetValue(view.SourceIndex, out var indexByNet) ||
            indexByNet.Count == 0)
        {
            return null;
        }

        cache = new KiCadSchematicHoverHitTestCache
        {
            CellSizeLocal = 24.0
        };

        foreach (var label in TabSchematics.EnumerateKiCadSchematicNetLabels(schematic))
        {
            string normalizedNetName = label.NormalizedText?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedNetName) ||
                label.At == null ||
                !indexByNet.ContainsKey(normalizedNetName))
            {
                continue;
            }

            Point localPoint = this.MapKiCadWorldToLocal(
                label.At.X,
                label.At.Y,
                worldBounds,
                contentRect,
                calibration);

            int candidateIndex = cache.LabelCandidates.Count;
            cache.LabelCandidates.Add(new KiCadSchematicHoverLabelCandidate
            {
                NormalizedNetName = normalizedNetName,
                LocalPoint = localPoint
            });

            int cellX = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.X, cache.CellSizeLocal);
            int cellY = KiCadHoverIndex.GetKiCadHoverCellCoord(localPoint.Y, cache.CellSizeLocal);

            KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
                cache.LabelIndicesByCell,
                cellX,
                cellX,
                cellY,
                cellY,
                candidateIndex);
        }

        foreach (var pair in indexByNet)
        {
            string normalizedNetName = pair.Key?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalizedNetName))
            {
                continue;
            }

            foreach (var resolvedPath in pair.Value)
            {
                if (resolvedPath.Points.Count < 2)
                {
                    continue;
                }

                Point previousPoint = this.MapKiCadWorldToLocal(
                    resolvedPath.Points[0].X,
                    resolvedPath.Points[0].Y,
                    worldBounds,
                    contentRect,
                    calibration);

                for (int i = 1; i < resolvedPath.Points.Count; i++)
                {
                    Point currentPoint = this.MapKiCadWorldToLocal(
                        resolvedPath.Points[i].X,
                        resolvedPath.Points[i].Y,
                        worldBounds,
                        contentRect,
                        calibration);

                    int candidateIndex = cache.SegmentCandidates.Count;
                    cache.SegmentCandidates.Add(new KiCadSchematicHoverSegmentCandidate
                    {
                        NormalizedNetName = normalizedNetName,
                        StartLocal = previousPoint,
                        EndLocal = currentPoint
                    });

                    double minX = Math.Min(previousPoint.X, currentPoint.X);
                    double maxX = Math.Max(previousPoint.X, currentPoint.X);
                    double minY = Math.Min(previousPoint.Y, currentPoint.Y);
                    double maxY = Math.Max(previousPoint.Y, currentPoint.Y);

                    KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
                        cache.SegmentIndicesByCell,
                        KiCadHoverIndex.GetKiCadHoverCellCoord(minX, cache.CellSizeLocal),
                        KiCadHoverIndex.GetKiCadHoverCellCoord(maxX, cache.CellSizeLocal),
                        KiCadHoverIndex.GetKiCadHoverCellCoord(minY, cache.CellSizeLocal),
                        KiCadHoverIndex.GetKiCadHoverCellCoord(maxY, cache.CellSizeLocal),
                        candidateIndex);

                    previousPoint = currentPoint;
                }
            }
        }

        this.thisKiCadSchematicHoverHitTestCacheByKey[cacheKey] = cache;
        return cache;
    }
}