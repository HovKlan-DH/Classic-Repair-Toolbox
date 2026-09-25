using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadNetGraphBuilder - the heaviest pure logic in the app, and the reason Tier B
// was worth doing. It decides which pads, tracks, vias, arcs and zone pours belong to the same
// electrical net on the side of the board you are looking at.
//
// It used to be a 546-line private instance method on TabSchematics that used no instance state
// at all, so nothing could reach it. A bug in here does not crash: it highlights the wrong
// copper, which is exactly the failure that wastes someone's afternoon at the bench.
public sealed class KiCadNetGraphBuilderTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    // A board where net 1 (GND) has two tracks meeting end-to-end, a via, and a pour;
    // net 2 (CLK) is a separate island so cross-net leakage is detectable.
    private const string Board = """
    (kicad_pcb (version 20221018)
      (net 0 "")
      (net 1 "GND")
      (net 2 "CLK")

      (footprint "DIP-14" (layer "F.Cu")
        (at 0 0)
        (fp_text reference "U1" (at 0 0) (layer "F.SilkS"))
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND"))
        (pad "7" thru_hole oval (at 20 0) (size 1.6 1.6) (layers "*.Cu") (net 2 "CLK"))
      )

      (segment (start 0 0) (end 10 0) (width 0.25) (layer "F.Cu") (net 1))
      (segment (start 10 0) (end 10 10) (width 0.25) (layer "F.Cu") (net 1))
      (segment (start 40 40) (end 50 40) (width 0.25) (layer "B.Cu") (net 1))
      (segment (start 20 0) (end 30 0) (width 0.25) (layer "F.Cu") (net 2))

      (via (at 10 10) (size 0.8) (drill 0.4) (layers "F.Cu" "B.Cu") (net 1))
      (arc (start 10 10) (mid 12 12) (end 14 10) (width 0.25) (layer "F.Cu") (net 1))

      (zone (net 1) (net_name "GND") (layer "F.Cu")
        (filled_polygon (pts (xy 8 8) (xy 20 8) (xy 20 20) (xy 8 20)))
      )
    )
    """;

    private async Task<KiCadPcb> LoadBoardAsync(string content = null!)
    {
        string path = this.thisWorkspace.WriteFile("board.kicad_pcb", content ?? Board);
        KiCadProjectRoot? root = await KiCadRawProjectLoader.LoadAsync(new[] { path });
        Assert.NotNull(root);
        return root!.Pcb[0];
    }

    private static KiCadPcbHighlightBucket Bucket(KiCadPcb pcb, string netId) =>
        pcb.HighlightIndex.TryGetValue(netId, out var bucket) ? bucket : new KiCadPcbHighlightBucket();

    // ------------------------------------------------------------------- cache keys

    [Fact]
    public void A_net_cache_key_separates_board_side_net_and_pcb()
    {
        string front = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(0, "1", "F.Cu");
        string back = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(0, "1", "B.Cu");
        string otherNet = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(0, "2", "F.Cu");
        string otherPcb = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(1, "1", "F.Cu");

        Assert.Equal(4, new[] { front, back, otherNet, otherPcb }.Distinct().Count());
    }

    [Fact]
    public void A_net_cache_key_ignores_surrounding_whitespace()
    {
        Assert.Equal(
            KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(0, "1", "F.Cu"),
            KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(0, "  1  ", "  F.Cu  "));
    }

    [Fact]
    public void A_world_point_key_is_stable_and_position_specific()
    {
        Assert.Equal(
            KiCadNetGraphBuilder.BuildKiCadWorldPointKey(new Point(1.5, 2.5)),
            KiCadNetGraphBuilder.BuildKiCadWorldPointKey(new Point(1.5, 2.5)));

        Assert.NotEqual(
            KiCadNetGraphBuilder.BuildKiCadWorldPointKey(new Point(1.5, 2.5)),
            KiCadNetGraphBuilder.BuildKiCadWorldPointKey(new Point(2.5, 1.5)));
    }

    [Fact]
    public void A_world_point_key_uses_an_invariant_decimal_point()
    {
        // A comma separator here would make two different points collide on a Danish machine.
        Assert.DoesNotContain(",", KiCadNetGraphBuilder.BuildKiCadWorldPointKey(new Point(1.5, 2.5)));
    }

    // ------------------------------------------------------------------ node building

    [Fact]
    public async Task Every_kind_of_copper_on_the_net_becomes_a_node()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        Assert.NotEmpty(cache.PadNodes);
        Assert.NotEmpty(cache.SegmentNodes);
        Assert.NotEmpty(cache.ViaNodes);
        Assert.NotEmpty(cache.ArcNodes);
        Assert.NotEmpty(cache.ZoneNodes);
    }

    [Fact]
    public async Task Only_copper_on_the_inspected_side_is_included()
    {
        // The third GND segment is on B.Cu and must not appear in an F.Cu view.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var front = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");
        var back = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "B.Cu");

        Assert.Equal(2, front.SegmentNodes.Count);
        Assert.Equal(40, Assert.Single(back.SegmentNodes).StartWorld.X);
    }

    [Fact]
    public async Task A_through_hole_pad_appears_on_both_sides()
    {
        // The pads are on "*.Cu", so they belong to every copper view.
        KiCadPcb pcb = await this.LoadBoardAsync();

        Assert.NotEmpty(KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu").PadNodes);
        Assert.NotEmpty(KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "B.Cu").PadNodes);
    }

    [Fact]
    public async Task Copper_from_a_different_net_never_leaks_in()
    {
        // The CLK track runs from (20,0) to (30,0); no GND node may sit there.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var gnd = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        Assert.DoesNotContain(gnd.SegmentNodes, s => s.StartWorld.X == 20 && s.EndWorld.X == 30);
    }

    [Fact]
    public async Task Each_node_gets_a_unique_id_and_is_registered()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        var ids = cache.PadNodes.Select(n => n.Info.Id)
            .Concat(cache.SegmentNodes.Select(n => n.Info.Id))
            .Concat(cache.ViaNodes.Select(n => n.Info.Id))
            .Concat(cache.ArcNodes.Select(n => n.Info.Id))
            .Concat(cache.ZoneNodes.Select(n => n.Info.Id))
            .ToList();

        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.All(ids, id => Assert.True(cache.NodesById.ContainsKey(id)));
        Assert.All(ids, id => Assert.Contains(id, cache.AllNodeIds));
    }

    [Fact]
    public async Task An_empty_bucket_produces_an_empty_graph()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(
            pcb, new KiCadPcbHighlightBucket(), "F.Cu");

        Assert.Empty(cache.AllNodeIds);
        Assert.Empty(cache.SegmentNodes);
    }

    // -------------------------------------------------------------------- adjacency

    [Fact]
    public async Task Two_tracks_meeting_end_to_end_are_connected()
    {
        // (0,0)-(10,0) and (10,0)-(10,10) share a point, so selecting one must reach the other.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        var first = cache.SegmentNodes.Single(s => s.StartWorld.X == 0);
        var second = cache.SegmentNodes.Single(s => s.EndWorld.Y == 10);

        Assert.True(cache.AdjacencyByNodeId.TryGetValue(first.Info.Id, out var neighbours));
        Assert.Contains(second.Info.Id, neighbours!);
    }

    [Fact]
    public async Task Adjacency_is_symmetric()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        foreach (var (nodeId, neighbours) in cache.AdjacencyByNodeId)
        {
            foreach (string neighbour in neighbours)
            {
                Assert.True(
                    cache.AdjacencyByNodeId.TryGetValue(neighbour, out var back) && back!.Contains(nodeId),
                    $"adjacency {nodeId} -> {neighbour} is not mirrored");
            }
        }
    }

    [Fact]
    public async Task A_node_is_never_adjacent_to_itself()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        foreach (var (nodeId, neighbours) in cache.AdjacencyByNodeId)
        {
            Assert.DoesNotContain(nodeId, neighbours);
        }
    }

    [Fact]
    public async Task Every_adjacency_endpoint_refers_to_a_real_node()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        foreach (var (nodeId, neighbours) in cache.AdjacencyByNodeId)
        {
            Assert.Contains(nodeId, cache.AllNodeIds);
            foreach (string neighbour in neighbours)
            {
                Assert.Contains(neighbour, cache.AllNodeIds);
            }
        }
    }

    [Fact]
    public async Task A_pour_participates_in_connectivity_so_a_track_can_continue_into_it()
    {
        // The zone covers (8,8)-(20,20); the second track ends at (10,10) inside it.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        var zone = Assert.Single(cache.ZoneNodes);

        Assert.True(cache.AdjacencyByNodeId.TryGetValue(zone.Info.Id, out var neighbours));
        Assert.NotEmpty(neighbours!);
    }

    [Fact]
    public async Task Distant_copper_on_the_same_net_is_not_falsely_connected()
    {
        // Two GND tracks far apart with nothing between them must stay separate islands.
        const string islands = """
        (kicad_pcb (version 20221018)
          (net 1 "GND")
          (segment (start 0 0) (end 10 0) (width 0.25) (layer "F.Cu") (net 1))
          (segment (start 900 900) (end 910 900) (width 0.25) (layer "F.Cu") (net 1))
        )
        """;

        KiCadPcb pcb = await this.LoadBoardAsync(islands);

        var cache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, Bucket(pcb, "1"), "F.Cu");

        var a = cache.SegmentNodes.Single(s => s.StartWorld.X == 0);
        var b = cache.SegmentNodes.Single(s => s.StartWorld.X == 900);

        cache.AdjacencyByNodeId.TryGetValue(a.Info.Id, out var neighbours);
        Assert.DoesNotContain(b.Info.Id, neighbours ?? new HashSet<string>());
    }

    // ------------------------------------------------- segment chaining for drawing

    private static KiCadPcbSegmentRenderNode Seg(double x1, double y1, double x2, double y2, string id) =>
        new()
        {
            Info = new KiCadGraphNode { Id = id },
            StartWorld = new Point(x1, y1),
            EndWorld = new Point(x2, y2),
            WidthWorld = 0.25
        };

    [Fact]
    public void Segments_meeting_end_to_end_chain_into_one_polyline()
    {
        var chains = KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[]
        {
            Seg(0, 0, 10, 0, "a"),
            Seg(10, 0, 10, 10, "b"),
            Seg(10, 10, 20, 10, "c")
        });

        List<Point> chain = Assert.Single(chains);
        Assert.Equal(4, chain.Count);
        Assert.Equal(new Point(0, 0), chain[0]);
        Assert.Equal(new Point(20, 10), chain[^1]);
    }

    [Fact]
    public void Chaining_does_not_depend_on_the_order_segments_are_supplied_in()
    {
        var shuffled = KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[]
        {
            Seg(10, 10, 20, 10, "c"),
            Seg(0, 0, 10, 0, "a"),
            Seg(10, 0, 10, 10, "b")
        });

        Assert.Single(shuffled);
        Assert.Equal(4, shuffled[0].Count);
    }

    [Fact]
    public void Chaining_follows_a_segment_written_back_to_front()
    {
        // KiCad does not guarantee start/end orientation between adjacent tracks.
        var chains = KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[]
        {
            Seg(0, 0, 10, 0, "a"),
            Seg(10, 10, 10, 0, "b")   // reversed relative to 'a'
        });

        List<Point> chain = Assert.Single(chains);
        Assert.Equal(3, chain.Count);
    }

    [Fact]
    public void Disconnected_runs_produce_separate_chains()
    {
        var chains = KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[]
        {
            Seg(0, 0, 10, 0, "a"),
            Seg(50, 50, 60, 50, "b")
        });

        Assert.Equal(2, chains.Count);
        Assert.All(chains, c => Assert.Equal(2, c.Count));
    }

    [Fact]
    public void A_single_segment_becomes_a_two_point_chain()
    {
        List<Point> chain = Assert.Single(
            KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[] { Seg(0, 0, 10, 0, "a") }));

        Assert.Equal(2, chain.Count);
    }

    [Fact]
    public void No_segments_produce_no_chains()
    {
        Assert.Empty(KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(
            Array.Empty<KiCadPcbSegmentRenderNode>()));
    }

    [Fact]
    public void A_branching_junction_still_covers_every_segment()
    {
        // Three tracks meeting at one point cannot be a single polyline, but nothing may be lost.
        var chains = KiCadNetGraphBuilder.BuildConnectedKiCadPcbSegmentPointChains(new[]
        {
            Seg(0, 0, 10, 0, "a"),
            Seg(10, 0, 20, 0, "b"),
            Seg(10, 0, 10, 10, "c")
        });

        int totalPoints = chains.Sum(c => c.Count);

        Assert.True(chains.Count >= 1);
        Assert.True(totalPoints >= 4, "every segment endpoint should still be drawn");
    }
}
