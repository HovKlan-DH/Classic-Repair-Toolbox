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
// Draws the KiCad overlay onto the schematic: refresh scheduling, PCB and schematic
// geometry rendering, stroke sizing, and pin-1 marking.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // Per-net primitive cache. A hover change alters the appearance of exactly two nets - the one
    // entered and the one left - but rebuilding the overlay used to recompute all of them, profiled
    // at ~222 ms with everything selected on a dense board. Caching each net's primitives against
    // the state they were built for turns that into recomputing two nets.
    //
    // The decision of when an entry may be reused lives in KiCadOverlayNetCache, where it is unit
    // tested: serving an entry that should have been dropped draws copper for a state that no longer
    // applies, and that looks entirely plausible on screen rather than failing anything.
    private readonly KiCadOverlayNetCache<KiCadOverlayPrimitive> thisKiCadNetPrimitiveCache = new();

    private bool thisIsKiCadOverlayRefreshQueued;

    private int thisKiCadOverlayRefreshRequestVersion;

    private int thisKiCadOverlayLastRenderedVersion;

    // ###########################################################################################
    // Clears the imported KiCad overlay geometry.
    // ###########################################################################################
    private void ClearKiCadOverlay()
    {
        this.SchematicsKiCadOverlayCanvas.ClearGeometry();
    }

    // ###########################################################################################
    // Returns the active KiCad calibration for the current schematic.
    // Uses the temporary interactive box calibration while calibration mode is active; otherwise
    // loads the persisted box calibration from the board JSON file.
    // ###########################################################################################
    private KiCadViewCalibration GetKiCadViewCalibration(string schematicName)
    {
        if (this.thisIsKiCadTraceCalibrationMode &&
            string.Equals(this.GetCurrentSchematicName(), schematicName, StringComparison.OrdinalIgnoreCase) &&
            this.currentFullResBitmap != null &&
            this.currentFullResBitmap.PixelSize.Width > 0 &&
            this.currentFullResBitmap.PixelSize.Height > 0)
        {
            double left = Math.Min(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
            double right = Math.Max(this.thisKiCadCalibrationImageLeft, this.thisKiCadCalibrationImageRight);
            double top = Math.Min(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);
            double bottom = Math.Max(this.thisKiCadCalibrationImageTop, this.thisKiCadCalibrationImageBottom);

            return new KiCadViewCalibration
            {
                ScaleX = (right - left) / this.currentFullResBitmap.PixelSize.Width,
                ScaleY = (bottom - top) / this.currentFullResBitmap.PixelSize.Height,
                OffsetX = left,
                OffsetY = top,
                MirrorX = this.thisKiCadCalibrationImageLeft > this.thisKiCadCalibrationImageRight,
                MirrorY = this.thisKiCadCalibrationImageTop > this.thisKiCadCalibrationImageBottom
            };
        }

        string excelPath = this.MainWindow?.GetCurrentBoardExcelPath() ?? string.Empty;

        if (BoardComponentHighlightStorage.TryLoadKiCadCalibration(
                excelPath,
                schematicName,
                out _,
                out double offsetX,
                out double offsetY,
                out double scaleX,
                out double scaleY,
                out bool mirrorX,
                out bool mirrorY))
        {
            return new KiCadViewCalibration
            {
                ScaleX = scaleX,
                ScaleY = scaleY,
                OffsetX = offsetX,
                OffsetY = offsetY,
                MirrorX = mirrorX,
                MirrorY = mirrorY
            };
        }

        return KiCadViewCalibration.Identity;
    }

    // ###########################################################################################
    // Queues normal KiCad overlay refreshes, but allows blink-driven callers to bypass the queue
    // and render immediately so visual blinking stays synchronized with the main highlight layer.
    // Version tracking prevents stale queued callbacks from redrawing after an immediate refresh.
    // ###########################################################################################
    private void RefreshKiCadOverlay(bool forceImmediate = false)
    {
        this.thisKiCadOverlayRefreshRequestVersion = unchecked(this.thisKiCadOverlayRefreshRequestVersion + 1);

        if (forceImmediate)
        {
            this.thisIsKiCadOverlayRefreshQueued = false;

            int renderVersion = this.thisKiCadOverlayRefreshRequestVersion;
            this.RefreshKiCadOverlayNow();
            this.thisKiCadOverlayLastRenderedVersion = renderVersion;

            if (this.thisKiCadOverlayLastRenderedVersion != this.thisKiCadOverlayRefreshRequestVersion)
            {
                this.RefreshKiCadOverlay();
            }

            return;
        }

        if (this.thisIsKiCadOverlayRefreshQueued)
        {
            return;
        }

        this.thisIsKiCadOverlayRefreshQueued = true;

        Dispatcher.UIThread.Post(() =>
        {
            this.thisIsKiCadOverlayRefreshQueued = false;

            if (this.thisKiCadOverlayLastRenderedVersion == this.thisKiCadOverlayRefreshRequestVersion)
            {
                return;
            }

            int renderVersion = this.thisKiCadOverlayRefreshRequestVersion;
            this.RefreshKiCadOverlayNow();
            this.thisKiCadOverlayLastRenderedVersion = renderVersion;

            if (this.thisKiCadOverlayLastRenderedVersion != this.thisKiCadOverlayRefreshRequestVersion)
            {
                this.RefreshKiCadOverlay();
            }
        }, DispatcherPriority.Background);
    }

    // ###########################################################################################
    // Rebuilds the currently visible KiCad overlay for the selected image view immediately.
    // While trace calibration mode is active, the temporary calibration transform is used and the
    // temporary traces-and-pads checkbox can suppress all overlay geometry except the box itself.
    // ###########################################################################################
    private void RefreshKiCadOverlayNow()
    {
        this.ClearKiCadOverlay();

        var activeTracePreviewReferences = this.BuildActiveKiCadTracePreviewReferences();
        var activeTracePreviewNets = this.BuildActiveKiCadTracePreviewNetNames();

        this.UpdateKiCadNetConnectionsPanel(activeTracePreviewNets);

        bool hasActiveKiCadNets = activeTracePreviewNets.Count > 0;

        if (this.thisKiCadProject == null || this.currentFullResBitmap == null)
        {
            return;
        }

        var currentView = this.ResolveKiCadViewForCurrentSchematic();
        if (currentView == null)
        {
            return;
        }

        if (this.thisIsKiCadTraceCalibrationMode)
        {
            bool thisShouldShowCalibrationTracesAndPads =
                this.CheckGlobalShowCalibrationTracesAndPads.IsChecked != false;

            if (thisShouldShowCalibrationTracesAndPads)
            {
                if (string.Equals(currentView.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(currentView.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
                {
                    var calibrationNetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    var pcb = this.thisKiCadProject.Root.Pcb.ElementAtOrDefault(currentView.SourceIndex);
                    if (pcb != null)
                    {
                        foreach (var net in pcb.Nets.List)
                        {
                            string normalizedName = net.NormalizedName?.Trim() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(normalizedName))
                            {
                                calibrationNetNames.Add(normalizedName);
                            }
                        }
                    }

                    this.RenderKiCadPcbGeometry(
                        currentView,
                        new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                        calibrationNetNames);
                }
                else if (string.Equals(currentView.Type, "schematic", StringComparison.OrdinalIgnoreCase))
                {
                    if (this.thisKiCadProject.SchematicNetPathIndexBySchematicIndex.TryGetValue(currentView.SourceIndex, out var indexByNet))
                    {
                        var calibrationNetNames = new HashSet<string>(
                            indexByNet.Keys.Where(key => !string.IsNullOrWhiteSpace(key)),
                            StringComparer.OrdinalIgnoreCase);

                        this.RenderKiCadSchematicGeometry(currentView, calibrationNetNames);
                    }
                }
            }

            var primitives = this.SchematicsKiCadOverlayCanvas.Primitives.ToList();
            primitives.Add(this.BuildKiCadCalibrationBoxPrimitive());
            this.SchematicsKiCadOverlayCanvas.SetGeometry(primitives);

            return;
        }

        if (string.Equals(currentView.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(currentView.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase))
        {
            if (!hasActiveKiCadNets && !this.HasPin1HighlightTargetReference())
            {
                return;
            }

            this.RenderKiCadPcbGeometry(currentView, activeTracePreviewReferences, activeTracePreviewNets);
            return;
        }

        if (!hasActiveKiCadNets)
        {
            return;
        }

        if (string.Equals(currentView.Type, "schematic", StringComparison.OrdinalIgnoreCase))
        {
            this.RenderKiCadSchematicGeometry(currentView, activeTracePreviewNets);
        }
    }

    // ###########################################################################################
    // Returns true when the given pad is the primary pad that should receive the special marker.
    // Prefers pad "1" when it exists; otherwise falls back to the first visible pad designator.
    // ###########################################################################################
    private static bool IsPrimaryPadForPin1Highlight(
        KiCadPcbFootprint footprint,
        KiCadPcbPad pad,
        string requiredLayer)
    {
        string currentPadDesignator = pad.Number?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(currentPadDesignator))
        {
            return false;
        }

        if (string.Equals(currentPadDesignator, "1", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var visiblePadDesignators = footprint.Pads
            .Where(candidate => candidate.AbsoluteCenter != null)
            .Where(candidate => KiCadLayerGeometry.IsPointVisibleOnSide(candidate.Layers, requiredLayer))
            .Select(candidate => candidate.Number?.Trim() ?? string.Empty)
            .Where(candidate => !string.IsNullOrWhiteSpace(candidate))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (visiblePadDesignators.Count == 0)
        {
            return false;
        }

        if (visiblePadDesignators.Any(candidate =>
                string.Equals(candidate, "1", StringComparison.OrdinalIgnoreCase)))
        {
            return false;
        }

        string primaryDesignator = visiblePadDesignators
            .OrderBy(candidate => candidate, Comparer<string>.Create(KiCadLayerGeometry.ComparePadDesignators))
            .First();

        return string.Equals(currentPadDesignator, primaryDesignator, StringComparison.OrdinalIgnoreCase);
    }

    // ###########################################################################################
    // Returns true when the supplied pad belongs to a selected or hovered component and should
    // receive the special primary-pin highlight.
    // ###########################################################################################
    private bool ShouldUseSelectedComponentPin1Highlight(
        KiCadPcbFootprint footprint,
        KiCadPcbPad pad,
        string requiredLayer)
    {
        if (!this.IsBoardMarkPin1OnSelectedComponentEnabled())
        {
            return false;
        }

        string normalizedReference = footprint.Reference?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalizedReference))
        {
            return false;
        }

        bool isTargetComponent =
            this.thisSelectedKiCadReferences.Contains(normalizedReference) ||
            string.Equals(this.thisHoveredComponentBoardLabel, normalizedReference, StringComparison.OrdinalIgnoreCase);

        if (!isTargetComponent)
        {
            return false;
        }

        return TabSchematics.IsPrimaryPadForPin1Highlight(footprint, pad, requiredLayer);
    }

    // ###########################################################################################
    // Returns true when there is a selected or hovered component reference that can receive a
    // special pin-1 marker on the current PCB KiCad overlay.
    // ###########################################################################################
    private bool HasPin1HighlightTargetReference()
    {
        if (!this.IsBoardMarkPin1OnSelectedComponentEnabled())
        {
            return false;
        }

        return this.thisSelectedKiCadReferences.Count > 0 ||
               !string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel);
    }

    // ###########################################################################################
    // Returns the on-screen KiCad PCB stroke thickness.
    // Keeps most of the mapped width so split trace chains meet cleanly, while trimming only a
    // small amount to avoid the overlay looking too bloated.
    // ###########################################################################################
    private static double GetKiCadOverlayStrokeThickness(double mappedThickness)
    {
        if (mappedThickness <= 1.0)
        {
            return 1.0;
        }

        double trim = Math.Min(0.35, mappedThickness * 0.10);
        return Math.Max(1.0, mappedThickness - trim);
    }

    // ###########################################################################################
    // Renders PCB copper geometry for a precomputed set of active references and net names.
    // Missing KiCad net caches are built in the background so the UI can render immediately and
    // refresh itself when the heavy continuity graph becomes available.
    // Opposite-side traces now share the same opacity behavior as the primary side, so only the
    // opposite-side color remains configurable from board data.
    // ###########################################################################################
    private void RenderKiCadPcbGeometry(
        KiCadProjectView view,
        IReadOnlySet<string> activeReferences,
        IReadOnlySet<string> activeNets)
    {
        var root = this.thisKiCadProject?.Root;
        if (root == null ||
            view.SourceIndex < 0 ||
            view.SourceIndex >= root.Pcb.Count)
        {
            return;
        }

        var pcb = root.Pcb[view.SourceIndex];
        var contentRect = this.GetImageContentRect();
        var worldBounds = this.GetKiCadPcbWorldBounds(pcb);

        if (contentRect.Width <= 0 ||
            contentRect.Height <= 0 ||
            worldBounds.Width <= 0 ||
            worldBounds.Height <= 0)
        {
            return;
        }

        string currentSchematicName = this.GetCurrentSchematicName();
        var calibration = this.GetKiCadViewCalibration(currentSchematicName);

        Color overlayColor = Colors.DeepSkyBlue;
        double baseOpacity = 0.20;
        Color oppositeTraceHighlightColor = Colors.DodgerBlue;

        if (this.schematicByName.TryGetValue(currentSchematicName, out var schematicEntry))
        {
            overlayColor = RectGeometry.ParseColorOrDefault(schematicEntry.SchematicHighlightColor, Colors.DeepSkyBlue);
            baseOpacity = RectGeometry.ParseOpacityOrDefault(schematicEntry.SchematicHighlightOpacity, 0.20);
            oppositeTraceHighlightColor = RectGeometry.ParseColorOrDefault(schematicEntry.OppositeTraceHighlightColor, Colors.DodgerBlue);
        }

        double translatedOpacity = Math.Clamp(baseOpacity + 0.25, 0.0, 1.0);
        string? activeHoveredKiCadNetName = this.GetActiveHoveredKiCadNetName();

        var matchingNetIds = pcb.Nets.List
            .Where(net => !string.IsNullOrWhiteSpace(net.NormalizedName) &&
                          activeNets.Contains(net.NormalizedName.Trim()) &&
                          !string.IsNullOrWhiteSpace(net.Id))
            .Select(net => new { Id = net.Id!.Trim(), Name = net.NormalizedName!.Trim() })
            .Distinct()
            .ToList();

        string primaryLayer = string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase)
            ? "B.Cu"
            : "F.Cu";

        string oppositeLayer = string.Equals(primaryLayer, "F.Cu", StringComparison.OrdinalIgnoreCase)
            ? "B.Cu"
            : "F.Cu";

        bool showOppositeSideTraces = UserSettings.SchematicsShowOppositeSideTraces;
        bool showZones = UserSettings.SchematicsShowZones;

        var primitives = new List<KiCadOverlayPrimitive>();
        var firstPinBrush = this.ResolveThemeBrush("Schematics_FirstPin", new SolidColorBrush(Colors.Orange));

        // Everything a cached net depends on beyond its own appearance. Built by
        // KiCadOverlayCacheKeys so the field list can be tested: a missing field here does not fail a
        // build, it draws stale copper.
        string cacheGeneration = KiCadOverlayCacheKeys.BuildGenerationKey(new KiCadOverlaySharedState
        {
            BoardScopeKey = this.thisCurrentKiCadRuntimeCacheScopeKey,
            SchematicName = currentSchematicName,
            ViewId = view.Id ?? string.Empty,
            ViewSourceIndex = view.SourceIndex,
            PrimaryLayer = primaryLayer,
            ContentRect = contentRect,
            WorldBounds = worldBounds,
            CalibrationScaleX = calibration.ScaleX,
            CalibrationScaleY = calibration.ScaleY,
            CalibrationOffsetX = calibration.OffsetX,
            CalibrationOffsetY = calibration.OffsetY,
            CalibrationMirrorX = calibration.MirrorX,
            CalibrationMirrorY = calibration.MirrorY,
            OverlayColor = overlayColor,
            OppositeTraceColor = oppositeTraceHighlightColor,
            TranslatedOpacity = translatedOpacity,
            ShowOppositeSideTraces = showOppositeSideTraces,
            ShowZones = showZones,
            IsCalibrationMode = this.thisIsKiCadTraceCalibrationMode,
            ActiveReferences = activeReferences.ToList(),
            SelectedReferences = this.thisSelectedKiCadReferences.ToList()
        });

        this.thisKiCadNetPrimitiveCache.BeginRebuild(cacheGeneration);

        // Built once rather than per net. It does not vary by net, and a dense board runs this loop
        // 249 times per rebuild.
        var selectedImportantSignalNetNames = this.BuildSelectedImportantSignalNetNames();

        void AddPadPrimitive(List<KiCadOverlayPrimitive> target, KiCadPcbPad pad, IBrush padBrush)
        {
            if (pad.AbsoluteCenter == null)
            {
                return;
            }

            Point center = this.MapKiCadWorldToLocal(
                pad.AbsoluteCenter.X,
                pad.AbsoluteCenter.Y,
                worldBounds,
                contentRect,
                calibration);

            double width = MapKiCadWorldLengthToLocal(
                pad.Size?.X ?? 1.2,
                worldBounds,
                contentRect,
                calibration);

            double height = MapKiCadWorldLengthToLocal(
                pad.Size?.Y ?? 1.2,
                worldBounds,
                contentRect,
                calibration);

            var rect = new Rect(
                center.X - (Math.Max(2.0, width) / 2.0),
                center.Y - (Math.Max(2.0, height) / 2.0),
                Math.Max(2.0, width),
                Math.Max(2.0, height));

            var pen = new Pen(padBrush, 1.2);

            // The rect above is built axis-aligned from the pad's own width and height; the pad's
            // rotation is what turns it the right way round on the board. Without this a 90-degree
            // rect or oval pad is drawn with its width and height swapped.
            double rotationDegrees = KiCadPadGeometry.ResolveScreenRotationDegrees(
                pad.RotationDegrees,
                calibration.MirrorX,
                calibration.MirrorY);

            target.Add(new KiCadOverlayPrimitive
            {
                Kind = KiCadPadGeometry.IsRectangularShape(pad.Shape)
                    ? KiCadOverlayPrimitiveKind.Rectangle
                    : KiCadOverlayPrimitiveKind.Ellipse,
                Rect = rect,
                RotationDegrees = rotationDegrees,
                Pen = pen,
                Fill = padBrush
            });
        }

        void AddTracePrimitivesForLayer(
            List<KiCadOverlayPrimitive> target,
            KiCadPcbNetRenderCache cache,
            HashSet<string> activeDrawIds,
            IBrush strokeBrush)
        {
            var activeSegmentNodes = cache.SegmentNodes
                .Where(segmentNode => activeDrawIds.Contains(segmentNode.Info.Id))
                .ToList();

            foreach (var segmentGroup in activeSegmentNodes.GroupBy(
                         segmentNode => Math.Round(segmentNode.WidthWorld, 6)))
            {
                var groupedSegments = segmentGroup.ToList();

                if (groupedSegments.Count == 0)
                {
                    continue;
                }

                double thickness = GetKiCadOverlayStrokeThickness(
                    MapKiCadWorldLengthToLocal(
                        groupedSegments[0].WidthWorld,
                        worldBounds,
                        contentRect,
                        calibration));

                var pen = new Pen(strokeBrush, thickness);

                foreach (var chain in KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(groupedSegments))
                {
                    if (chain.Count < 2)
                    {
                        continue;
                    }

                    var localPoints = chain
                        .Select(point => this.MapKiCadWorldToLocal(
                            point.X,
                            point.Y,
                            worldBounds,
                            contentRect,
                            calibration))
                        .ToList();

                    target.Add(new KiCadOverlayPrimitive
                    {
                        Kind = KiCadOverlayPrimitiveKind.Polyline,
                        Points = localPoints,
                        Pen = pen
                    });
                }
            }

            foreach (var arcNode in cache.ArcNodes)
            {
                if (!activeDrawIds.Contains(arcNode.Info.Id))
                {
                    continue;
                }

                Point start = this.MapKiCadWorldToLocal(
                    arcNode.StartWorld.X,
                    arcNode.StartWorld.Y,
                    worldBounds,
                    contentRect,
                    calibration);

                Point mid = this.MapKiCadWorldToLocal(
                    arcNode.MidWorld.X,
                    arcNode.MidWorld.Y,
                    worldBounds,
                    contentRect,
                    calibration);

                Point end = this.MapKiCadWorldToLocal(
                    arcNode.EndWorld.X,
                    arcNode.EndWorld.Y,
                    worldBounds,
                    contentRect,
                    calibration);

                double thickness = GetKiCadOverlayStrokeThickness(
                    MapKiCadWorldLengthToLocal(
                        arcNode.WidthWorld,
                        worldBounds,
                        contentRect,
                        calibration));

                var sampledArcPoints = SampleQuadraticBezier(start, mid, end, 20);

                target.Add(new KiCadOverlayPrimitive
                {
                    Kind = KiCadOverlayPrimitiveKind.Polyline,
                    Points = sampledArcPoints,
                    Pen = new Pen(strokeBrush, thickness)
                });
            }
        }

        foreach (var netInfo in matchingNetIds)
        {
            if (!pcb.HighlightIndex.TryGetValue(netInfo.Id, out var bucket))
            {
                continue;
            }

            bool isHoveredNet = string.Equals(activeHoveredKiCadNetName, netInfo.Name, StringComparison.OrdinalIgnoreCase);
            bool isLockedNet = this.thisLockedKiCadNetNames.Contains(netInfo.Name);

            bool isImportantSignalDerivedNet = selectedImportantSignalNetNames.Contains(netInfo.Name);

            bool isExplicitHighlight = isLockedNet || isHoveredNet || isImportantSignalDerivedNet;

            bool isSelectionDerivedNet = this.thisSelectedKiCadNormalizedNetNames.Contains(netInfo.Name);
            bool shouldBlinkThisNet = isLockedNet || isSelectionDerivedNet || isImportantSignalDerivedNet;

            double blinkFactor = shouldBlinkThisNet ? this.thisCurrentHighlightBlinkFactor : 1.0;
            double effectiveOpacity = Math.Clamp(translatedOpacity * blinkFactor, 0.0, 1.0);

            // Per-net, not global: hovering a component only changes pin-1 marking on that
            // component's own pads. See KiCadOverlayCacheKeys for why that distinction matters.
            string hoveredComponentOnThisNet = KiCadOverlayCacheKeys.ResolveHoveredComponentForNet(
                this.thisHoveredComponentBoardLabel,
                bucket.Pads.Select(padRef => padRef.Reference));

            var appearance = new KiCadNetAppearance(
                isExplicitHighlight,
                isHoveredNet,
                shouldBlinkThisNet,
                effectiveOpacity,
                hoveredComponentOnThisNet);

            if (this.thisKiCadNetPrimitiveCache.TryGet(netInfo.Id, appearance, out var cachedNetPrimitives))
            {
                primitives.AddRange(cachedNetPrimitives);
                continue;
            }

            var netPrimitives = new List<KiCadOverlayPrimitive>();

            SolidColorBrush primaryBrush = isHoveredNet && !shouldBlinkThisNet
                ? new SolidColorBrush(overlayColor, 1.0)
                : new SolidColorBrush(overlayColor, effectiveOpacity);

            IBrush oppositeBrush = isHoveredNet && !shouldBlinkThisNet
                ? new SolidColorBrush(oppositeTraceHighlightColor, 1.0)
                : new SolidColorBrush(oppositeTraceHighlightColor, effectiveOpacity);

            if (showOppositeSideTraces)
            {
                var oppositeCache = this.GetOrCreateKiCadPcbNetRenderCache(
                    pcb,
                    view.SourceIndex,
                    netInfo.Id,
                    bucket,
                    oppositeLayer);

                if (oppositeCache != null)
                {
                    var oppositeActiveDrawIds = KiCadNetGraphBuilder.BuildKiCadPcbActiveDrawIds(
                        oppositeCache,
                        isExplicitHighlight,
                        activeReferences);

                    if (showZones)
                    {
                        foreach (var zoneNode in oppositeCache.ZoneNodes)
                        {
                            if (!oppositeActiveDrawIds.Contains(zoneNode.Info.Id))
                            {
                                continue;
                            }

                            Geometry? zoneGeometry = this.BuildKiCadZoneGeometry(
                                zoneNode.PolygonsWorld,
                                worldBounds,
                                contentRect,
                                calibration);

                            if (zoneGeometry == null)
                            {
                                continue;
                            }

                            double oppositeZoneFillOpacity = isHoveredNet && !shouldBlinkThisNet
                                ? Math.Min(1.0, Math.Clamp(translatedOpacity * 0.65, 0.10, 0.38) + 0.12)
                                : Math.Clamp(effectiveOpacity * 0.65, 0.10, 0.38);

                            netPrimitives.Add(new KiCadOverlayPrimitive
                            {
                                Kind = KiCadOverlayPrimitiveKind.Geometry,
                                Geometry = zoneGeometry,
                                Fill = new SolidColorBrush(oppositeTraceHighlightColor, oppositeZoneFillOpacity),
                                Pen = new Pen(oppositeBrush, 1.0)
                            });
                        }
                    }

                    AddTracePrimitivesForLayer(netPrimitives, oppositeCache, oppositeActiveDrawIds, oppositeBrush);
                }
            }

            var primaryCache = this.GetOrCreateKiCadPcbNetRenderCache(
                pcb,
                view.SourceIndex,
                netInfo.Id,
                bucket,
                primaryLayer);

            if (primaryCache == null)
            {
                continue;
            }

            var primaryActiveDrawIds = KiCadNetGraphBuilder.BuildKiCadPcbActiveDrawIds(
                primaryCache,
                isExplicitHighlight,
                activeReferences);

            if (showZones)
            {
                foreach (var zoneNode in primaryCache.ZoneNodes)
                {
                    if (!primaryActiveDrawIds.Contains(zoneNode.Info.Id))
                    {
                        continue;
                    }

                    Geometry? zoneGeometry = this.BuildKiCadZoneGeometry(
                        zoneNode.PolygonsWorld,
                        worldBounds,
                        contentRect,
                        calibration);

                    if (zoneGeometry == null)
                    {
                        continue;
                    }

                    double zoneFillOpacity = isHoveredNet && !shouldBlinkThisNet
                        ? 0.32
                        : Math.Clamp(effectiveOpacity * 0.65, 0.10, 0.38);

                    netPrimitives.Add(new KiCadOverlayPrimitive
                    {
                        Kind = KiCadOverlayPrimitiveKind.Geometry,
                        Geometry = zoneGeometry,
                        Fill = new SolidColorBrush(overlayColor, zoneFillOpacity),
                        Pen = new Pen(primaryBrush, 1.0)
                    });
                }
            }

            AddTracePrimitivesForLayer(netPrimitives, primaryCache, primaryActiveDrawIds, primaryBrush);

            foreach (var viaNode in primaryCache.ViaNodes)
            {
                if (!primaryActiveDrawIds.Contains(viaNode.Info.Id))
                {
                    continue;
                }

                Point center = this.MapKiCadWorldToLocal(
                    viaNode.CenterWorld.X,
                    viaNode.CenterWorld.Y,
                    worldBounds,
                    contentRect,
                    calibration);

                double diameter = MapKiCadWorldLengthToLocal(
                    viaNode.DiameterWorld,
                    worldBounds,
                    contentRect,
                    calibration);

                netPrimitives.Add(new KiCadOverlayPrimitive
                {
                    Kind = KiCadOverlayPrimitiveKind.Ellipse,
                    Rect = new Rect(
                        center.X - (Math.Max(2.0, diameter) / 2.0),
                        center.Y - (Math.Max(2.0, diameter) / 2.0),
                        Math.Max(2.0, diameter),
                        Math.Max(2.0, diameter)),
                    Pen = new Pen(primaryBrush, 1.2),
                    Fill = primaryBrush
                });
            }

            foreach (var padNode in primaryCache.PadNodes)
            {
                if (!primaryActiveDrawIds.Contains(padNode.Info.Id))
                {
                    continue;
                }

                bool isSelectedComponentPin1 = this.ShouldUseSelectedComponentPin1Highlight(
                    padNode.Footprint,
                    padNode.Pad,
                    primaryLayer);

                IBrush padBrush = isSelectedComponentPin1
                    ? firstPinBrush
                    : primaryBrush;

                AddPadPrimitive(netPrimitives, padNode.Pad, padBrush);
            }

            this.thisKiCadNetPrimitiveCache.Store(netInfo.Id, appearance, netPrimitives);
            primitives.AddRange(netPrimitives);
        }

        if (this.HasPin1HighlightTargetReference())
        {
            foreach (var footprint in pcb.Footprints)
            {
                string reference = footprint.Reference?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(reference))
                {
                    continue;
                }

                if (!this.thisSelectedKiCadReferences.Contains(reference) &&
                    !string.Equals(this.thisHoveredComponentBoardLabel, reference, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                foreach (var pad in footprint.Pads)
                {
                    if (pad.AbsoluteCenter == null ||
                        !KiCadLayerGeometry.IsPointVisibleOnSide(pad.Layers, primaryLayer) ||
                        !TabSchematics.IsPrimaryPadForPin1Highlight(footprint, pad, primaryLayer))
                    {
                        continue;
                    }

                    AddPadPrimitive(primitives, pad, firstPinBrush);
                }
            }
        }


        this.SchematicsKiCadOverlayCanvas.SetGeometry(primitives);
    }

    // ###########################################################################################
    // Renders resolved schematic wire paths for a precomputed set of active normalized net names.
    // This avoids recomputing hover-preview net names multiple times within one refresh cycle.
    // ###########################################################################################
    private void RenderKiCadSchematicGeometry(
        KiCadProjectView view,
        IReadOnlySet<string> activeNets)
    {
        var bundle = this.thisKiCadProject;
        if (bundle == null ||
            view.SourceIndex < 0 ||
            view.SourceIndex >= bundle.Root.Schematics.Count)
        {
            return;
        }

        if (!bundle.SchematicNetPathIndexBySchematicIndex.TryGetValue(view.SourceIndex, out var indexByNet))
        {
            return;
        }

        var schematic = bundle.Root.Schematics[view.SourceIndex];
        var contentRect = this.GetImageContentRect();
        var worldBounds = this.GetKiCadSchematicWorldBounds(schematic);

        if (contentRect.Width <= 0 ||
            contentRect.Height <= 0 ||
            worldBounds.Width <= 0 ||
            worldBounds.Height <= 0)
        {
            return;
        }

        string currentSchematicName = this.GetCurrentSchematicName();
        var calibration = this.GetKiCadViewCalibration(currentSchematicName);

        Color overlayColor = Colors.Orange;
        double baseOpacity = 0.20;
        if (this.schematicByName.TryGetValue(currentSchematicName, out var schematicEntry))
        {
            overlayColor = RectGeometry.ParseColorOrDefault(schematicEntry.SchematicHighlightColor, Colors.Orange);
            baseOpacity = RectGeometry.ParseOpacityOrDefault(schematicEntry.SchematicHighlightOpacity, 0.20);
        }

        double translatedOpacity = Math.Clamp(baseOpacity + 0.25, 0.0, 1.0);
        string? activeHoveredKiCadNetName = this.GetActiveHoveredKiCadNetName();

        var primitives = new List<KiCadOverlayPrimitive>();

        foreach (string normalizedNetName in activeNets)
        {
            if (!indexByNet.TryGetValue(normalizedNetName, out var resolvedPaths))
            {
                continue;
            }

            bool isHoveredNet = string.Equals(activeHoveredKiCadNetName, normalizedNetName, StringComparison.OrdinalIgnoreCase);
            bool isLockedNet = this.thisLockedKiCadNetNames.Contains(normalizedNetName);

            var selectedImportantSignalNetNames = this.BuildSelectedImportantSignalNetNames();
            bool isImportantSignalDerivedNet = selectedImportantSignalNetNames.Contains(normalizedNetName);

            bool isSelectionDerivedNet = this.thisSelectedKiCadNormalizedNetNames.Contains(normalizedNetName);
            bool shouldBlinkThisNet = isLockedNet || isSelectionDerivedNet || isImportantSignalDerivedNet;

            double blinkFactor = shouldBlinkThisNet ? this.thisCurrentHighlightBlinkFactor : 1.0;
            double effectiveOpacity = Math.Clamp(translatedOpacity * blinkFactor, 0.0, 1.0);

            IBrush strokeBrush = isHoveredNet && !shouldBlinkThisNet
                ? new SolidColorBrush(overlayColor, 1.0)
                : new SolidColorBrush(overlayColor, effectiveOpacity);

            var pen = new Pen(strokeBrush, 1.2);

            foreach (var resolvedPath in resolvedPaths)
            {
                if (resolvedPath.Points.Count < 2)
                {
                    continue;
                }

                var localPoints = resolvedPath.Points
                    .Select(point => this.MapKiCadWorldToLocal(
                        point.X,
                        point.Y,
                        worldBounds,
                        contentRect,
                        calibration))
                    .ToList();

                primitives.Add(new KiCadOverlayPrimitive
                {
                    Kind = KiCadOverlayPrimitiveKind.Polyline,
                    Points = localPoints,
                    Pen = pen
                });
            }
        }

        this.SchematicsKiCadOverlayCanvas.SetGeometry(primitives);
    }
}