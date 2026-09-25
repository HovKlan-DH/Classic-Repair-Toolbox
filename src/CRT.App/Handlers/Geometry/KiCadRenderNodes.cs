using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // Data types describing one KiCad net as a graph of drawable copper: pads, track segments,
    // vias, arcs and zones, plus the spatial hover index built over them.
    //
    // These were private nested types inside TabSchematics, which meant the ~750 lines of logic
    // that build them could not be reached by a test. They are plain DTOs - no behaviour, no
    // control references - and live here so KiCadNetGraphBuilder and KiCadHoverIndex can be
    // tested directly.
    // ###########################################################################################

    internal class KiCadGraphNode
    {
        public string Id { get; set; } = string.Empty;
        public bool IsTargetPad { get; set; }
        public bool IsForeignPad { get; set; }
        public int SegmentIndex { get; set; } = -1;
        public int ViaIndex { get; set; } = -1;
        public int ArcIndex { get; set; } = -1;
        public KiCadPcbHighlightPadRef? PadRef { get; set; }
    }    

    internal sealed class KiCadPcbPadRenderNode
    {
        public KiCadGraphNode Info { get; init; } = new();
        public KiCadPcbFootprint Footprint { get; init; } = null!;
        public KiCadPcbPad Pad { get; init; } = null!;
        public Point CenterWorld { get; init; }
        public double RadiusWorld { get; init; }
    }

    internal sealed class KiCadPcbSegmentRenderNode
    {
        public KiCadGraphNode Info { get; init; } = new();
        public Point StartWorld { get; init; }
        public Point EndWorld { get; init; }
        public double WidthWorld { get; init; }
    }

    internal sealed class KiCadPcbViaRenderNode
    {
        public KiCadGraphNode Info { get; init; } = new();
        public Point CenterWorld { get; init; }
        public double DiameterWorld { get; init; }
    }

    internal sealed class KiCadPcbArcRenderNode
    {
        public KiCadGraphNode Info { get; init; } = new();
        public Point StartWorld { get; init; }
        public Point MidWorld { get; init; }
        public Point EndWorld { get; init; }
        public double WidthWorld { get; init; }
    }

    internal sealed class KiCadPcbZoneRenderNode
    {
        public KiCadGraphNode Info { get; init; } = new();
        public KiCadPcbZone Zone { get; init; } = null!;
        public IReadOnlyList<IReadOnlyList<Point>> PolygonsWorld { get; init; } = Array.Empty<IReadOnlyList<Point>>();
        public Rect BoundsWorld { get; init; }
    }

    internal sealed class KiCadPcbNetRenderCache
    {
        public Dictionary<string, KiCadGraphNode> NodesById { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, HashSet<string>> AdjacencyByNodeId { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> PadReferenceByNodeId { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public HashSet<string> AllNodeIds { get; init; } = new(StringComparer.OrdinalIgnoreCase);
        public List<KiCadPcbPadRenderNode> PadNodes { get; init; } = new();
        public List<KiCadPcbSegmentRenderNode> SegmentNodes { get; init; } = new();
        public List<KiCadPcbViaRenderNode> ViaNodes { get; init; } = new();
        public List<KiCadPcbArcRenderNode> ArcNodes { get; init; } = new();
        public List<KiCadPcbZoneRenderNode> ZoneNodes { get; init; } = new();
    }

    internal sealed class KiCadPcbHoverPadCandidate
    {
        public KiCadNetRef Net { get; init; } = null!;
        public string PadNumber { get; init; } = string.Empty;
        public Point CenterWorld { get; init; }
        public double HitRadiusWorld { get; init; }
    }

    internal sealed class KiCadPcbHoverSegmentCandidate
    {
        public KiCadNetRef Net { get; init; } = null!;
        public Point StartWorld { get; init; }
        public Point EndWorld { get; init; }
        public double HitRadiusWorld { get; init; }
    }

    internal sealed class KiCadPcbHoverViaCandidate
    {
        public KiCadNetRef Net { get; init; } = null!;
        public Point CenterWorld { get; init; }
        public double HitRadiusWorld { get; init; }
    }

    internal sealed class KiCadPcbHoverZoneCandidate
    {
        public KiCadNetRef Net { get; init; } = null!;
        public IReadOnlyList<IReadOnlyList<Point>> PolygonsWorld { get; init; } = Array.Empty<IReadOnlyList<Point>>();
        public Rect BoundsWorld { get; init; }
    }

    internal sealed class KiCadPcbHoverHitTestCache
    {
        public double CellSizeWorld { get; init; } = 2.0;
        public double MaxHitRadiusWorld { get; set; } = 0.8;

        public List<KiCadPcbHoverPadCandidate> PadCandidates { get; init; } = new();
        public List<KiCadPcbHoverSegmentCandidate> SegmentCandidates { get; init; } = new();
        public List<KiCadPcbHoverViaCandidate> ViaCandidates { get; init; } = new();
        public List<KiCadPcbHoverZoneCandidate> ZoneCandidates { get; init; } = new();

        public Dictionary<long, List<int>> PadIndicesByCell { get; init; } = new();
        public Dictionary<long, List<int>> SegmentIndicesByCell { get; init; } = new();
        public Dictionary<long, List<int>> ViaIndicesByCell { get; init; } = new();
        public Dictionary<long, List<int>> ZoneIndicesByCell { get; init; } = new();
    }
}
