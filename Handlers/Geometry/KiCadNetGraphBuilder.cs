using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Turns one KiCad net into a connected graph of drawable copper.
    //
    // This is the heaviest and highest-risk pure logic in the Schematics tab: it decides which
    // pads, track segments, vias, arcs and zone pours belong to the same electrical net on one
    // side of the board, and chains loose segments into polylines for drawing. A mistake here
    // does not crash - it silently highlights the wrong copper.
    //
    // Extracted from TabSchematics (where it was ~750 lines of unreachable instance methods that
    // used no instance state at all).
    // ###########################################################################################
    internal static class KiCadNetGraphBuilder
    {

    // ###########################################################################################
    // Builds a stable cache key for one PCB net graph on one board side.
    // ###########################################################################################
    public static string BuildKiCadPcbNetRenderCacheKey(int pcbIndex, string netId, string requiredLayer)
    {
        return string.Join(
            "\u001F",
            pcbIndex.ToString(CultureInfo.InvariantCulture),
            netId?.Trim() ?? string.Empty,
            requiredLayer.Trim());
    }

    // ###########################################################################################
    // Builds a stable rounded key for one PCB world-space point so connected trace endpoints can
    // be grouped into continuous rendered chains without relying on exact floating-point equality.
    // ###########################################################################################
    public static string BuildKiCadWorldPointKey(Point point)
    {
        return string.Join(
            "|",
            Math.Round(point.X, 6).ToString(CultureInfo.InvariantCulture),
            Math.Round(point.Y, 6).ToString(CultureInfo.InvariantCulture));
    }

    // ###########################################################################################
    // Builds one cached PCB net graph containing pads, segments, vias, arcs, zones, and adjacency.
    // Zones participate in connectivity so selected traces can continue into copper pours.
    // Uses a broad-phase zone spatial index so exact zone-touch tests only run for nearby geometry.
    // ###########################################################################################
    public static KiCadPcbNetRenderCache BuildKiCadPcbNetRenderCache(
        KiCadPcb pcb,
        KiCadPcbHighlightBucket bucket,
        string requiredLayer)
    {
        var cache = new KiCadPcbNetRenderCache();

        int idCounter = 0;

        foreach (var padRef in bucket.Pads)
        {
            if (padRef.FootprintIndex < 0 || padRef.FootprintIndex >= pcb.Footprints.Count)
            {
                continue;
            }

            var footprint = pcb.Footprints[padRef.FootprintIndex];
            if (padRef.PadIndex < 0 || padRef.PadIndex >= footprint.Pads.Count)
            {
                continue;
            }

            var pad = footprint.Pads[padRef.PadIndex];
            if (pad.AbsoluteCenter == null ||
                !KiCadLayerGeometry.IsPointVisibleOnSide(pad.Layers, requiredLayer))
            {
                continue;
            }

            var info = new KiCadGraphNode
            {
                Id = $"P{idCounter++}",
                PadRef = padRef
            };

            cache.NodesById[info.Id] = info;
            cache.AllNodeIds.Add(info.Id);
            cache.PadReferenceByNodeId[info.Id] = footprint.Reference?.Trim() ?? string.Empty;

            cache.PadNodes.Add(new KiCadPcbPadRenderNode
            {
                Info = info,
                Footprint = footprint,
                Pad = pad,
                CenterWorld = new Point(pad.AbsoluteCenter.X, pad.AbsoluteCenter.Y),
                RadiusWorld = Math.Max(pad.Size?.X ?? 1.2, pad.Size?.Y ?? 1.2) / 2.0
            });
        }

        foreach (int segmentIndex in bucket.Segments)
        {
            if (segmentIndex < 0 || segmentIndex >= pcb.Routing.Segments.Count)
            {
                continue;
            }

            var segment = pcb.Routing.Segments[segmentIndex];
            if (segment.Start == null ||
                segment.End == null ||
                !string.Equals(segment.Layer?.Trim(), requiredLayer, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var info = new KiCadGraphNode
            {
                Id = $"S{idCounter++}",
                SegmentIndex = segmentIndex
            };

            cache.NodesById[info.Id] = info;
            cache.AllNodeIds.Add(info.Id);

            cache.SegmentNodes.Add(new KiCadPcbSegmentRenderNode
            {
                Info = info,
                StartWorld = new Point(segment.Start.X, segment.Start.Y),
                EndWorld = new Point(segment.End.X, segment.End.Y),
                WidthWorld = segment.Width ?? 0.25
            });
        }

        foreach (int viaIndex in bucket.Vias)
        {
            if (viaIndex < 0 || viaIndex >= pcb.Routing.Vias.Count)
            {
                continue;
            }

            var via = pcb.Routing.Vias[viaIndex];
            if (via.At == null ||
                !KiCadLayerGeometry.IsPointVisibleOnSide(via.Layers, requiredLayer))
            {
                continue;
            }

            var info = new KiCadGraphNode
            {
                Id = $"V{idCounter++}",
                ViaIndex = viaIndex
            };

            cache.NodesById[info.Id] = info;
            cache.AllNodeIds.Add(info.Id);

            cache.ViaNodes.Add(new KiCadPcbViaRenderNode
            {
                Info = info,
                CenterWorld = new Point(via.At.X, via.At.Y),
                DiameterWorld = via.Size ?? 0.8
            });
        }

        foreach (int arcIndex in bucket.Arcs)
        {
            if (arcIndex < 0 || arcIndex >= pcb.Routing.Arcs.Count)
            {
                continue;
            }

            var arc = pcb.Routing.Arcs[arcIndex];
            if (arc.Start == null ||
                arc.Mid == null ||
                arc.End == null ||
                !string.Equals(arc.Layer?.Trim(), requiredLayer, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var info = new KiCadGraphNode
            {
                Id = $"A{idCounter++}",
                ArcIndex = arcIndex
            };

            cache.NodesById[info.Id] = info;
            cache.AllNodeIds.Add(info.Id);

            cache.ArcNodes.Add(new KiCadPcbArcRenderNode
            {
                Info = info,
                StartWorld = new Point(arc.Start.X, arc.Start.Y),
                MidWorld = new Point(arc.Mid.X, arc.Mid.Y),
                EndWorld = new Point(arc.End.X, arc.End.Y),
                WidthWorld = arc.Width ?? 0.25
            });
        }

        foreach (int zoneIndex in bucket.Zones)
        {
            if (zoneIndex < 0 || zoneIndex >= pcb.Routing.Zones.Count)
            {
                continue;
            }

            var zone = pcb.Routing.Zones[zoneIndex];
            if (!KiCadLayerGeometry.IsZoneVisibleOnSide(zone, requiredLayer))
            {
                continue;
            }

            var polygonsWorld = KiCadLayerGeometry.GetZoneWorldPolygons(zone);
            if (polygonsWorld.Count == 0)
            {
                continue;
            }

            var boundsWorld = PolygonGeometry.GetPolygonSetBounds(polygonsWorld);
            if (boundsWorld.Width <= 0 || boundsWorld.Height <= 0)
            {
                continue;
            }

            var info = new KiCadGraphNode
            {
                Id = $"Z{idCounter++}"
            };

            cache.NodesById[info.Id] = info;
            cache.AllNodeIds.Add(info.Id);

            cache.ZoneNodes.Add(new KiCadPcbZoneRenderNode
            {
                Info = info,
                Zone = zone,
                PolygonsWorld = polygonsWorld,
                BoundsWorld = boundsWorld
            });
        }

        void AddEdge(string id1, string id2)
        {
            if (!cache.AdjacencyByNodeId.TryGetValue(id1, out var set1))
            {
                set1 = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                cache.AdjacencyByNodeId[id1] = set1;
            }

            set1.Add(id2);

            if (!cache.AdjacencyByNodeId.TryGetValue(id2, out var set2))
            {
                set2 = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                cache.AdjacencyByNodeId[id2] = set2;
            }

            set2.Add(id1);
        }

        static Rect BuildCircleBounds(Point centerWorld, double radiusWorld)
        {
            double safeRadius = Math.Max(0.05, radiusWorld);

            return new Rect(
                centerWorld.X - safeRadius,
                centerWorld.Y - safeRadius,
                safeRadius * 2.0,
                safeRadius * 2.0);
        }

        static Rect BuildSegmentBounds(Point startWorld, Point endWorld, double radiusWorld)
        {
            double safeRadius = Math.Max(0.05, radiusWorld);
            double minX = Math.Min(startWorld.X, endWorld.X) - safeRadius;
            double minY = Math.Min(startWorld.Y, endWorld.Y) - safeRadius;
            double maxX = Math.Max(startWorld.X, endWorld.X) + safeRadius;
            double maxY = Math.Max(startWorld.Y, endWorld.Y) + safeRadius;

            return new Rect(
                minX,
                minY,
                Math.Max(0.0001, maxX - minX),
                Math.Max(0.0001, maxY - minY));
        }

        static Rect BuildArcBounds(KiCadPcbArcRenderNode arcNode, double radiusWorld)
        {
            double safeRadius = Math.Max(0.05, radiusWorld);
            double minX = Math.Min(arcNode.StartWorld.X, Math.Min(arcNode.MidWorld.X, arcNode.EndWorld.X)) - safeRadius;
            double minY = Math.Min(arcNode.StartWorld.Y, Math.Min(arcNode.MidWorld.Y, arcNode.EndWorld.Y)) - safeRadius;
            double maxX = Math.Max(arcNode.StartWorld.X, Math.Max(arcNode.MidWorld.X, arcNode.EndWorld.X)) + safeRadius;
            double maxY = Math.Max(arcNode.StartWorld.Y, Math.Max(arcNode.MidWorld.Y, arcNode.EndWorld.Y)) + safeRadius;

            return new Rect(
                minX,
                minY,
                Math.Max(0.0001, maxX - minX),
                Math.Max(0.0001, maxY - minY));
        }

        static bool RectsIntersect(Rect left, Rect right)
        {
            return left.Left <= right.Right &&
                   left.Right >= right.Left &&
                   left.Top <= right.Bottom &&
                   left.Bottom >= right.Top;
        }

        for (int i = 0; i < cache.SegmentNodes.Count; i++)
        {
            for (int j = i + 1; j < cache.SegmentNodes.Count; j++)
            {
                var s1 = cache.SegmentNodes[i];
                var s2 = cache.SegmentNodes[j];

                double dist = Math.Min(
                    Math.Min(
                        PolygonGeometry.DistanceToSegment(s1.StartWorld, s2.StartWorld.X, s2.StartWorld.Y, s2.EndWorld.X, s2.EndWorld.Y),
                        PolygonGeometry.DistanceToSegment(s1.EndWorld, s2.StartWorld.X, s2.StartWorld.Y, s2.EndWorld.X, s2.EndWorld.Y)),
                    Math.Min(
                        PolygonGeometry.DistanceToSegment(s2.StartWorld, s1.StartWorld.X, s1.StartWorld.Y, s1.EndWorld.X, s1.EndWorld.Y),
                        PolygonGeometry.DistanceToSegment(s2.EndWorld, s1.StartWorld.X, s1.StartWorld.Y, s1.EndWorld.X, s1.EndWorld.Y)));

                if (dist <= (s1.WidthWorld / 2.0) + (s2.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(s1.Info.Id, s2.Info.Id);
                }
            }
        }

        foreach (var padNode in cache.PadNodes)
        {
            foreach (var segmentNode in cache.SegmentNodes)
            {
                if (PolygonGeometry.DistanceToSegment(
                        padNode.CenterWorld,
                        segmentNode.StartWorld.X,
                        segmentNode.StartWorld.Y,
                        segmentNode.EndWorld.X,
                        segmentNode.EndWorld.Y) <= padNode.RadiusWorld + (segmentNode.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(padNode.Info.Id, segmentNode.Info.Id);
                }
            }
        }

        foreach (var padNode in cache.PadNodes)
        {
            foreach (var viaNode in cache.ViaNodes)
            {
                double dx = padNode.CenterWorld.X - viaNode.CenterWorld.X;
                double dy = padNode.CenterWorld.Y - viaNode.CenterWorld.Y;
                double dist = Math.Sqrt((dx * dx) + (dy * dy));

                if (dist <= padNode.RadiusWorld + (viaNode.DiameterWorld / 2.0) + 0.05)
                {
                    AddEdge(padNode.Info.Id, viaNode.Info.Id);
                }
            }
        }

        foreach (var viaNode in cache.ViaNodes)
        {
            foreach (var segmentNode in cache.SegmentNodes)
            {
                if (PolygonGeometry.DistanceToSegment(
                        viaNode.CenterWorld,
                        segmentNode.StartWorld.X,
                        segmentNode.StartWorld.Y,
                        segmentNode.EndWorld.X,
                        segmentNode.EndWorld.Y) <= (viaNode.DiameterWorld / 2.0) + (segmentNode.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(viaNode.Info.Id, segmentNode.Info.Id);
                }
            }
        }

        foreach (var arcNode in cache.ArcNodes)
        {
            foreach (var segmentNode in cache.SegmentNodes)
            {
                double dist = Math.Min(
                    PolygonGeometry.DistanceToSegment(
                        arcNode.StartWorld,
                        segmentNode.StartWorld.X,
                        segmentNode.StartWorld.Y,
                        segmentNode.EndWorld.X,
                        segmentNode.EndWorld.Y),
                    Math.Min(
                        PolygonGeometry.DistanceToSegment(
                            arcNode.MidWorld,
                            segmentNode.StartWorld.X,
                            segmentNode.StartWorld.Y,
                            segmentNode.EndWorld.X,
                            segmentNode.EndWorld.Y),
                        PolygonGeometry.DistanceToSegment(
                            arcNode.EndWorld,
                            segmentNode.StartWorld.X,
                            segmentNode.StartWorld.Y,
                            segmentNode.EndWorld.X,
                            segmentNode.EndWorld.Y)));

                if (dist <= (arcNode.WidthWorld / 2.0) + (segmentNode.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(arcNode.Info.Id, segmentNode.Info.Id);
                }
            }

            foreach (var padNode in cache.PadNodes)
            {
                double dist = Math.Min(
                    Math.Sqrt(Math.Pow(padNode.CenterWorld.X - arcNode.StartWorld.X, 2) + Math.Pow(padNode.CenterWorld.Y - arcNode.StartWorld.Y, 2)),
                    Math.Min(
                        Math.Sqrt(Math.Pow(padNode.CenterWorld.X - arcNode.MidWorld.X, 2) + Math.Pow(padNode.CenterWorld.Y - arcNode.MidWorld.Y, 2)),
                        Math.Sqrt(Math.Pow(padNode.CenterWorld.X - arcNode.EndWorld.X, 2) + Math.Pow(padNode.CenterWorld.Y - arcNode.EndWorld.Y, 2))));

                if (dist <= padNode.RadiusWorld + (arcNode.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(padNode.Info.Id, arcNode.Info.Id);
                }
            }

            foreach (var viaNode in cache.ViaNodes)
            {
                double dist = Math.Min(
                    Math.Sqrt(Math.Pow(viaNode.CenterWorld.X - arcNode.StartWorld.X, 2) + Math.Pow(viaNode.CenterWorld.Y - arcNode.StartWorld.Y, 2)),
                    Math.Min(
                        Math.Sqrt(Math.Pow(viaNode.CenterWorld.X - arcNode.MidWorld.X, 2) + Math.Pow(viaNode.CenterWorld.Y - arcNode.MidWorld.Y, 2)),
                        Math.Sqrt(Math.Pow(viaNode.CenterWorld.X - arcNode.EndWorld.X, 2) + Math.Pow(viaNode.CenterWorld.Y - arcNode.EndWorld.Y, 2))));

                if (dist <= (viaNode.DiameterWorld / 2.0) + (arcNode.WidthWorld / 2.0) + 0.05)
                {
                    AddEdge(viaNode.Info.Id, arcNode.Info.Id);
                }
            }
        }

        const double zoneGridCellSizeWorld = 8.0;
        var zoneIndicesByCell = new Dictionary<long, List<int>>();
        var zoneBoundsByIndex = new List<Rect>(cache.ZoneNodes.Count);

        for (int i = 0; i < cache.ZoneNodes.Count; i++)
        {
            Rect boundsWorld = cache.ZoneNodes[i].BoundsWorld;
            zoneBoundsByIndex.Add(boundsWorld);

            int minCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(boundsWorld.Left, zoneGridCellSizeWorld);
            int maxCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(boundsWorld.Right, zoneGridCellSizeWorld);
            int minCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(boundsWorld.Top, zoneGridCellSizeWorld);
            int maxCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(boundsWorld.Bottom, zoneGridCellSizeWorld);

            KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
                zoneIndicesByCell,
                minCellX,
                maxCellX,
                minCellY,
                maxCellY,
                i);
        }

        List<int> GetCandidateZoneIndices(Rect candidateBounds)
        {
            if (zoneBoundsByIndex.Count == 0)
            {
                return new List<int>();
            }

            int minCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(candidateBounds.Left, zoneGridCellSizeWorld);
            int maxCellX = KiCadHoverIndex.GetKiCadHoverCellCoord(candidateBounds.Right, zoneGridCellSizeWorld);
            int minCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(candidateBounds.Top, zoneGridCellSizeWorld);
            int maxCellY = KiCadHoverIndex.GetKiCadHoverCellCoord(candidateBounds.Bottom, zoneGridCellSizeWorld);

            var result = new List<int>();
            var seen = new HashSet<int>();

            for (int cellY = minCellY; cellY <= maxCellY; cellY++)
            {
                for (int cellX = minCellX; cellX <= maxCellX; cellX++)
                {
                    long cellKey = KiCadHoverIndex.BuildKiCadHoverCellKey(cellX, cellY);

                    if (!zoneIndicesByCell.TryGetValue(cellKey, out var zoneIndices))
                    {
                        continue;
                    }

                    foreach (int zoneIndex in zoneIndices)
                    {
                        if (!seen.Add(zoneIndex))
                        {
                            continue;
                        }

                        if (!RectsIntersect(candidateBounds, zoneBoundsByIndex[zoneIndex]))
                        {
                            continue;
                        }

                        result.Add(zoneIndex);
                    }
                }
            }

            return result;
        }

        foreach (var padNode in cache.PadNodes)
        {
            var candidateZoneIndices = GetCandidateZoneIndices(
                BuildCircleBounds(
                    padNode.CenterWorld,
                    padNode.RadiusWorld + 0.05));

            foreach (int zoneIndex in candidateZoneIndices)
            {
                var zoneNode = cache.ZoneNodes[zoneIndex];

                if (PolygonGeometry.DoesCircleTouchZone(
                        padNode.CenterWorld,
                        padNode.RadiusWorld + 0.05,
                        zoneNode.PolygonsWorld))
                {
                    AddEdge(zoneNode.Info.Id, padNode.Info.Id);
                }
            }
        }

        foreach (var viaNode in cache.ViaNodes)
        {
            var candidateZoneIndices = GetCandidateZoneIndices(
                BuildCircleBounds(
                    viaNode.CenterWorld,
                    (viaNode.DiameterWorld / 2.0) + 0.05));

            foreach (int zoneIndex in candidateZoneIndices)
            {
                var zoneNode = cache.ZoneNodes[zoneIndex];

                if (PolygonGeometry.DoesCircleTouchZone(
                        viaNode.CenterWorld,
                        (viaNode.DiameterWorld / 2.0) + 0.05,
                        zoneNode.PolygonsWorld))
                {
                    AddEdge(zoneNode.Info.Id, viaNode.Info.Id);
                }
            }
        }

        foreach (var segmentNode in cache.SegmentNodes)
        {
            var candidateZoneIndices = GetCandidateZoneIndices(
                BuildSegmentBounds(
                    segmentNode.StartWorld,
                    segmentNode.EndWorld,
                    (segmentNode.WidthWorld / 2.0) + 0.05));

            foreach (int zoneIndex in candidateZoneIndices)
            {
                var zoneNode = cache.ZoneNodes[zoneIndex];

                if (PolygonGeometry.DoesSegmentTouchZone(
                        segmentNode.StartWorld,
                        segmentNode.EndWorld,
                        (segmentNode.WidthWorld / 2.0) + 0.05,
                        zoneNode.PolygonsWorld))
                {
                    AddEdge(zoneNode.Info.Id, segmentNode.Info.Id);
                }
            }
        }

        foreach (var arcNode in cache.ArcNodes)
        {
            var candidateZoneIndices = GetCandidateZoneIndices(
                BuildArcBounds(
                    arcNode,
                    (arcNode.WidthWorld / 2.0) + 0.05));

            foreach (int zoneIndex in candidateZoneIndices)
            {
                var zoneNode = cache.ZoneNodes[zoneIndex];

                if (PolygonGeometry.DoesArcTouchZone(
                        arcNode.StartWorld,
                        arcNode.MidWorld,
                        arcNode.EndWorld,
                        (arcNode.WidthWorld / 2.0) + 0.05,
                        zoneNode.PolygonsWorld))
                {
                    AddEdge(zoneNode.Info.Id, arcNode.Info.Id);
                }
            }
        }

        return cache;
    }

    // ###########################################################################################
    // Groups connected PCB segments into continuous point chains so the overlay can render one
    // smoothed polyline per trace run instead of many separate line primitives with visible seams.
    // ###########################################################################################
    public static List<List<Point>> BuildConnectedKiCadPcbSegmentPointChains(IReadOnlyList<KiCadPcbSegmentRenderNode> segmentNodes)
    {
        var chains = new List<List<Point>>();

        if (segmentNodes.Count == 0)
        {
            return chains;
        }

        var segmentIndicesByPointKey = new Dictionary<string, List<int>>(StringComparer.Ordinal);

        void AddSegmentIndex(string pointKey, int segmentIndex)
        {
            if (!segmentIndicesByPointKey.TryGetValue(pointKey, out var indices))
            {
                indices = new List<int>();
                segmentIndicesByPointKey[pointKey] = indices;
            }

            indices.Add(segmentIndex);
        }

        for (int i = 0; i < segmentNodes.Count; i++)
        {
            var segmentNode = segmentNodes[i];

            AddSegmentIndex(BuildKiCadWorldPointKey(segmentNode.StartWorld), i);
            AddSegmentIndex(BuildKiCadWorldPointKey(segmentNode.EndWorld), i);
        }

        var remainingSegmentIndices = new HashSet<int>(Enumerable.Range(0, segmentNodes.Count));

        int GetRemainingDegree(string pointKey)
        {
            if (!segmentIndicesByPointKey.TryGetValue(pointKey, out var indices))
            {
                return 0;
            }

            int degree = 0;

            for (int i = 0; i < indices.Count; i++)
            {
                if (remainingSegmentIndices.Contains(indices[i]))
                {
                    degree++;
                }
            }

            return degree;
        }

        Point GetOtherEndpoint(KiCadPcbSegmentRenderNode segmentNode, string currentPointKey)
        {
            string startKey = BuildKiCadWorldPointKey(segmentNode.StartWorld);
            return string.Equals(startKey, currentPointKey, StringComparison.Ordinal)
                ? segmentNode.EndWorld
                : segmentNode.StartWorld;
        }

        while (remainingSegmentIndices.Count > 0)
        {
            int seedSegmentIndex = remainingSegmentIndices.First();
            var seedSegment = segmentNodes[seedSegmentIndex];

            string seedStartKey = BuildKiCadWorldPointKey(seedSegment.StartWorld);
            string seedEndKey = BuildKiCadWorldPointKey(seedSegment.EndWorld);

            int seedStartDegree = GetRemainingDegree(seedStartKey);
            int seedEndDegree = GetRemainingDegree(seedEndKey);

            Point currentPoint;
            Point nextPoint;

            if (seedEndDegree != 2 && seedStartDegree == 2)
            {
                currentPoint = seedSegment.EndWorld;
                nextPoint = seedSegment.StartWorld;
            }
            else
            {
                currentPoint = seedSegment.StartWorld;
                nextPoint = seedSegment.EndWorld;
            }

            var chain = new List<Point> { currentPoint };
            int currentSegmentIndex = seedSegmentIndex;

            while (true)
            {
                remainingSegmentIndices.Remove(currentSegmentIndex);
                chain.Add(nextPoint);

                string nextPointKey = BuildKiCadWorldPointKey(nextPoint);

                if (!segmentIndicesByPointKey.TryGetValue(nextPointKey, out var connectedIndices))
                {
                    break;
                }

                int nextSegmentIndex = -1;

                for (int i = 0; i < connectedIndices.Count; i++)
                {
                    int candidateIndex = connectedIndices[i];

                    if (remainingSegmentIndices.Contains(candidateIndex))
                    {
                        if (nextSegmentIndex >= 0)
                        {
                            nextSegmentIndex = -1;
                            break;
                        }

                        nextSegmentIndex = candidateIndex;
                    }
                }

                if (nextSegmentIndex < 0)
                {
                    break;
                }

                var nextSegmentNode = segmentNodes[nextSegmentIndex];
                currentSegmentIndex = nextSegmentIndex;
                nextPoint = GetOtherEndpoint(nextSegmentNode, nextPointKey);
            }

            if (chain.Count >= 2)
            {
                chains.Add(chain);
            }
        }

        return chains;
    }

    // ###########################################################################################
    // Resolves the currently drawable node ids from a cached PCB net graph.
    // Explicit hover/lock draws the whole net, while selection-derived rendering starts from the
    // selected or hovered component pads and stops traversal at foreign pads.
    // ###########################################################################################
    public static HashSet<string> BuildKiCadPcbActiveDrawIds(
        KiCadPcbNetRenderCache cache,
        bool isExplicitHighlight,
        IReadOnlySet<string> activeReferences)
    {
        if (isExplicitHighlight)
        {
            return new HashSet<string>(cache.AllNodeIds, StringComparer.OrdinalIgnoreCase);
        }

        var activeDrawIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<string>();

        foreach (var padNode in cache.PadNodes)
        {
            string reference = padNode.Footprint.Reference?.Trim() ?? string.Empty;
            bool isTargetPad = activeReferences.Count == 0 ||
                               activeReferences.Contains(reference);

            if (!isTargetPad)
            {
                continue;
            }

            if (activeDrawIds.Add(padNode.Info.Id))
            {
                queue.Enqueue(padNode.Info.Id);
            }
        }

        while (queue.Count > 0)
        {
            string currentId = queue.Dequeue();

            if (!cache.AdjacencyByNodeId.TryGetValue(currentId, out var neighbors))
            {
                continue;
            }

            foreach (string neighborId in neighbors)
            {
                if (!activeDrawIds.Add(neighborId))
                {
                    continue;
                }

                bool isForeignPad =
                    cache.PadReferenceByNodeId.TryGetValue(neighborId, out string? reference) &&
                    activeReferences.Count > 0 &&
                    !activeReferences.Contains(reference);

                if (!isForeignPad)
                {
                    queue.Enqueue(neighborId);
                }
            }
        }

        return activeDrawIds;
    }
    }
}