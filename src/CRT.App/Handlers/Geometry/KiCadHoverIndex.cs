using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Builds the spatial index used to answer "what copper is under the pointer?" without
    // walking every item on the board.
    //
    // Candidates are bucketed into a uniform world-space grid; a hover test then only examines
    // the cells the pointer actually falls in. Extracted from TabSchematics, where it used no
    // instance state but could not be reached by a test.
    // ###########################################################################################
    internal static class KiCadHoverIndex
    {

    // ###########################################################################################
    // Builds a stable cache key for one PCB-side hover lookup cache.
    // ###########################################################################################
    public static string BuildKiCadPcbHoverHitTestCacheKey(int pcbIndex, string requiredLayer)
    {
        return string.Join(
            "\u001F",
            pcbIndex.ToString(CultureInfo.InvariantCulture),
            requiredLayer.Trim());
    }

    // ###########################################################################################
    // Packs one grid-cell coordinate pair into a stable dictionary key.
    // ###########################################################################################
    public static long BuildKiCadHoverCellKey(int cellX, int cellY)
    {
        return ((long)cellX << 32) ^ (uint)cellY;
    }

    // ###########################################################################################
    // Converts one world coordinate into the hover-grid cell coordinate for spatial lookup.
    // ###########################################################################################
    public static int GetKiCadHoverCellCoord(double worldCoord, double cellSizeWorld)
    {
        return (int)Math.Floor(worldCoord / Math.Max(0.0001, cellSizeWorld));
    }

    // ###########################################################################################
    // Adds one candidate index to every spatial cell touched by its expanded hit area.
    // ###########################################################################################
    public static void AddKiCadHoverIndexToCellRange(
        Dictionary<long, List<int>> cellMap,
        int minCellX,
        int maxCellX,
        int minCellY,
        int maxCellY,
        int candidateIndex)
    {
        for (int cellY = minCellY; cellY <= maxCellY; cellY++)
        {
            for (int cellX = minCellX; cellX <= maxCellX; cellX++)
            {
                long key = BuildKiCadHoverCellKey(cellX, cellY);

                if (!cellMap.TryGetValue(key, out var indices))
                {
                    indices = new List<int>();
                    cellMap[key] = indices;
                }

                indices.Add(candidateIndex);
            }
        }
    }

    // ###########################################################################################
    // Builds one spatial hover cache for a PCB side so pointer hover no longer scans every pad,
    // segment, via, and zone in the board on every move event.
    // ###########################################################################################
    public static KiCadPcbHoverHitTestCache BuildKiCadPcbHoverHitTestCache(KiCadPcb pcb, string requiredLayer)
    {
        var cache = new KiCadPcbHoverHitTestCache
        {
            CellSizeWorld = 2.0,
            MaxHitRadiusWorld = 0.8
        };

        foreach (var footprint in pcb.Footprints)
        {
            foreach (var pad in footprint.Pads)
            {
                if (pad.Net == null ||
                    string.IsNullOrWhiteSpace(pad.Net.NormalizedName) ||
                    pad.AbsoluteCenter == null ||
                    !KiCadLayerGeometry.IsPointVisibleOnSide(pad.Layers, requiredLayer))
                {
                    continue;
                }

                double hitRadiusWorld = Math.Max(pad.Size?.X ?? 0.5, pad.Size?.Y ?? 0.5) / 2.0 + 0.3;
                cache.MaxHitRadiusWorld = Math.Max(cache.MaxHitRadiusWorld, hitRadiusWorld);

                int candidateIndex = cache.PadCandidates.Count;

                cache.PadCandidates.Add(new KiCadPcbHoverPadCandidate
                {
                    Net = pad.Net,
                    PadNumber = pad.Number?.Trim() ?? string.Empty,
                    CenterWorld = new Point(pad.AbsoluteCenter.X, pad.AbsoluteCenter.Y),
                    HitRadiusWorld = hitRadiusWorld
                });

                int minCellX = GetKiCadHoverCellCoord(pad.AbsoluteCenter.X - hitRadiusWorld, cache.CellSizeWorld);
                int maxCellX = GetKiCadHoverCellCoord(pad.AbsoluteCenter.X + hitRadiusWorld, cache.CellSizeWorld);
                int minCellY = GetKiCadHoverCellCoord(pad.AbsoluteCenter.Y - hitRadiusWorld, cache.CellSizeWorld);
                int maxCellY = GetKiCadHoverCellCoord(pad.AbsoluteCenter.Y + hitRadiusWorld, cache.CellSizeWorld);

                AddKiCadHoverIndexToCellRange(
                    cache.PadIndicesByCell,
                    minCellX,
                    maxCellX,
                    minCellY,
                    maxCellY,
                    candidateIndex);
            }
        }

        foreach (var segment in pcb.Routing.Segments)
        {
            if (segment.Net == null ||
                string.IsNullOrWhiteSpace(segment.Net.NormalizedName) ||
                segment.Start == null ||
                segment.End == null ||
                !string.Equals(segment.Layer?.Trim(), requiredLayer, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            double hitRadiusWorld = (segment.Width ?? 0.25) / 2.0 + 0.3;
            cache.MaxHitRadiusWorld = Math.Max(cache.MaxHitRadiusWorld, hitRadiusWorld);

            int candidateIndex = cache.SegmentCandidates.Count;

            cache.SegmentCandidates.Add(new KiCadPcbHoverSegmentCandidate
            {
                Net = segment.Net,
                StartWorld = new Point(segment.Start.X, segment.Start.Y),
                EndWorld = new Point(segment.End.X, segment.End.Y),
                HitRadiusWorld = hitRadiusWorld
            });

            double minX = Math.Min(segment.Start.X, segment.End.X) - hitRadiusWorld;
            double maxX = Math.Max(segment.Start.X, segment.End.X) + hitRadiusWorld;
            double minY = Math.Min(segment.Start.Y, segment.End.Y) - hitRadiusWorld;
            double maxY = Math.Max(segment.Start.Y, segment.End.Y) + hitRadiusWorld;

            AddKiCadHoverIndexToCellRange(
                cache.SegmentIndicesByCell,
                GetKiCadHoverCellCoord(minX, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(maxX, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(minY, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(maxY, cache.CellSizeWorld),
                candidateIndex);
        }

        foreach (var via in pcb.Routing.Vias)
        {
            if (via.Net == null ||
                string.IsNullOrWhiteSpace(via.Net.NormalizedName) ||
                via.At == null ||
                !KiCadLayerGeometry.IsPointVisibleOnSide(via.Layers, requiredLayer))
            {
                continue;
            }

            double hitRadiusWorld = (via.Size ?? 0.4) / 2.0 + 0.3;
            cache.MaxHitRadiusWorld = Math.Max(cache.MaxHitRadiusWorld, hitRadiusWorld);

            int candidateIndex = cache.ViaCandidates.Count;

            cache.ViaCandidates.Add(new KiCadPcbHoverViaCandidate
            {
                Net = via.Net,
                CenterWorld = new Point(via.At.X, via.At.Y),
                HitRadiusWorld = hitRadiusWorld
            });

            int minCellX = GetKiCadHoverCellCoord(via.At.X - hitRadiusWorld, cache.CellSizeWorld);
            int maxCellX = GetKiCadHoverCellCoord(via.At.X + hitRadiusWorld, cache.CellSizeWorld);
            int minCellY = GetKiCadHoverCellCoord(via.At.Y - hitRadiusWorld, cache.CellSizeWorld);
            int maxCellY = GetKiCadHoverCellCoord(via.At.Y + hitRadiusWorld, cache.CellSizeWorld);

            AddKiCadHoverIndexToCellRange(
                cache.ViaIndicesByCell,
                minCellX,
                maxCellX,
                minCellY,
                maxCellY,
                candidateIndex);
        }

        const double zoneHoverToleranceWorld = 0.4;

        foreach (var zone in pcb.Routing.Zones)
        {
            if (zone.Net == null ||
                string.IsNullOrWhiteSpace(zone.Net.NormalizedName) ||
                !KiCadLayerGeometry.IsZoneVisibleOnSide(zone, requiredLayer))
            {
                continue;
            }

            var polygonsWorld = KiCadLayerGeometry.GetZoneWorldPolygons(zone);
            if (polygonsWorld.Count == 0)
            {
                continue;
            }

            Rect boundsWorld = PolygonGeometry.GetPolygonSetBounds(polygonsWorld);
            if (boundsWorld.Width <= 0 || boundsWorld.Height <= 0)
            {
                continue;
            }

            int candidateIndex = cache.ZoneCandidates.Count;

            cache.ZoneCandidates.Add(new KiCadPcbHoverZoneCandidate
            {
                Net = zone.Net,
                PolygonsWorld = polygonsWorld,
                BoundsWorld = boundsWorld
            });

            cache.MaxHitRadiusWorld = Math.Max(cache.MaxHitRadiusWorld, zoneHoverToleranceWorld);

            double minX = boundsWorld.Left - zoneHoverToleranceWorld;
            double maxX = boundsWorld.Right + zoneHoverToleranceWorld;
            double minY = boundsWorld.Top - zoneHoverToleranceWorld;
            double maxY = boundsWorld.Bottom + zoneHoverToleranceWorld;

            AddKiCadHoverIndexToCellRange(
                cache.ZoneIndicesByCell,
                GetKiCadHoverCellCoord(minX, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(maxX, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(minY, cache.CellSizeWorld),
                GetKiCadHoverCellCoord(maxY, cache.CellSizeWorld),
                candidateIndex);
        }

        return cache;
    }
    }
}