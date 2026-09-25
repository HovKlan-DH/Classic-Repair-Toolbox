using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for KiCadHoverIndex - the uniform spatial grid that answers "what copper is under the
// pointer?" without walking every item on the board.
//
// Previously a private instance method on TabSchematics that used no instance state. A bug here
// shows up as a net that will not highlight on hover, or one that highlights from too far away.
public sealed class KiCadHoverIndexTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private const string Board = """
    (kicad_pcb (version 20221018)
      (net 1 "GND")
      (net 2 "CLK")

      (footprint "DIP-14" (layer "F.Cu")
        (at 0 0)
        (fp_text reference "U1" (at 0 0) (layer "F.SilkS"))
        (pad "1" thru_hole rect (at 0 0) (size 1.6 1.6) (layers "*.Cu") (net 1 "GND"))
      )

      (segment (start 0 0) (end 10 0) (width 0.25) (layer "F.Cu") (net 1))
      (segment (start 40 40) (end 50 40) (width 0.25) (layer "B.Cu") (net 2))
      (via (at 10 10) (size 0.8) (drill 0.4) (layers "F.Cu" "B.Cu") (net 1))

      (zone (net 1) (net_name "GND") (layer "F.Cu")
        (filled_polygon (pts (xy 20 20) (xy 30 20) (xy 30 30) (xy 20 30)))
      )
    )
    """;

    private async Task<KiCadPcb> LoadBoardAsync()
    {
        string path = this.thisWorkspace.WriteFile("board.kicad_pcb", Board);
        KiCadProjectRoot? root = await KiCadRawProjectLoader.LoadAsync(new[] { path });
        Assert.NotNull(root);
        return root!.Pcb[0];
    }

    // ------------------------------------------------------------------ cell maths

    [Fact]
    public void A_cell_key_is_unique_per_cell()
    {
        var keys = new[]
        {
            KiCadHoverIndex.BuildKiCadHoverCellKey(0, 0),
            KiCadHoverIndex.BuildKiCadHoverCellKey(1, 0),
            KiCadHoverIndex.BuildKiCadHoverCellKey(0, 1),
            KiCadHoverIndex.BuildKiCadHoverCellKey(-1, 0),
            KiCadHoverIndex.BuildKiCadHoverCellKey(0, -1),
        };

        Assert.Equal(keys.Length, keys.Distinct().Count());
    }

    [Fact]
    public void A_cell_key_is_stable_for_the_same_cell()
    {
        Assert.Equal(
            KiCadHoverIndex.BuildKiCadHoverCellKey(7, -3),
            KiCadHoverIndex.BuildKiCadHoverCellKey(7, -3));
    }

    [Theory]
    [InlineData(0.0, 2.0, 0)]
    [InlineData(1.9, 2.0, 0)]
    [InlineData(2.0, 2.0, 1)]
    [InlineData(5.5, 2.0, 2)]
    [InlineData(-0.5, 2.0, -1)]
    [InlineData(-2.5, 2.0, -2)]
    public void A_world_coordinate_maps_to_its_cell(double world, double cellSize, int expected)
    {
        Assert.Equal(expected, KiCadHoverIndex.GetKiCadHoverCellCoord(world, cellSize));
    }

    [Fact]
    public void Negative_coordinates_floor_rather_than_truncate_toward_zero()
    {
        // Truncation would put -0.5 and +0.5 in the same cell and break hover near the origin.
        Assert.NotEqual(
            KiCadHoverIndex.GetKiCadHoverCellCoord(-0.5, 2.0),
            KiCadHoverIndex.GetKiCadHoverCellCoord(0.5, 2.0));
    }

    [Fact]
    public void A_hover_cache_key_separates_board_sides_and_boards()
    {
        var keys = new[]
        {
            KiCadHoverIndex.BuildKiCadPcbHoverHitTestCacheKey(0, "F.Cu"),
            KiCadHoverIndex.BuildKiCadPcbHoverHitTestCacheKey(0, "B.Cu"),
            KiCadHoverIndex.BuildKiCadPcbHoverHitTestCacheKey(1, "F.Cu"),
        };

        Assert.Equal(3, keys.Distinct().Count());
    }

    // ------------------------------------------------- AddKiCadHoverIndexToCellRange

    [Fact]
    public void An_item_is_registered_in_every_cell_its_extent_covers()
    {
        var index = new Dictionary<long, List<int>>();

        // Cells 0..2 on each axis => 9 cells.
        KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
            index, minCellX: 0, maxCellX: 2, minCellY: 0, maxCellY: 2, candidateIndex: 7);

        Assert.Equal(9, index.Count);
        Assert.All(index.Values, v => Assert.Contains(7, v));
    }

    [Fact]
    public void An_item_inside_a_single_cell_is_registered_once()
    {
        var index = new Dictionary<long, List<int>>();

        KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
            index, minCellX: 4, maxCellX: 4, minCellY: 4, maxCellY: 4, candidateIndex: 3);

        Assert.Single(index);
        Assert.Equal(new[] { 3 }, Assert.Single(index.Values));
    }

    [Fact]
    public void Several_items_can_share_a_cell()
    {
        var index = new Dictionary<long, List<int>>();

        KiCadHoverIndex.AddKiCadHoverIndexToCellRange(index, 0, 0, 0, 0, 1);
        KiCadHoverIndex.AddKiCadHoverIndexToCellRange(index, 0, 0, 0, 0, 2);

        Assert.Equal(new[] { 1, 2 }, Assert.Single(index.Values));
    }

    [Fact]
    public void An_item_spanning_negative_cells_is_registered_across_them()
    {
        var index = new Dictionary<long, List<int>>();

        KiCadHoverIndex.AddKiCadHoverIndexToCellRange(
            index, minCellX: -1, maxCellX: 1, minCellY: -1, maxCellY: 1, candidateIndex: 5);

        Assert.Equal(9, index.Count);
    }

    // ---------------------------------------------------------------- cache building

    [Fact]
    public async Task Pads_segments_vias_and_zones_all_become_hover_candidates()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        Assert.NotEmpty(cache.PadCandidates);
        Assert.NotEmpty(cache.SegmentCandidates);
        Assert.NotEmpty(cache.ViaCandidates);
        Assert.NotEmpty(cache.ZoneCandidates);
    }

    [Fact]
    public async Task Only_copper_on_the_inspected_side_becomes_a_candidate()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var front = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");
        var back = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "B.Cu");

        Assert.Single(front.SegmentCandidates);
        Assert.Equal(40, Assert.Single(back.SegmentCandidates).StartWorld.X);
    }

    [Fact]
    public async Task Every_candidate_carries_the_net_it_belongs_to()
    {
        // Without this the hover would light up copper but not know which signal it is.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        Assert.All(cache.SegmentCandidates, c => Assert.NotNull(c.Net));
        Assert.All(cache.PadCandidates, c => Assert.NotNull(c.Net));
        Assert.All(cache.ViaCandidates, c => Assert.NotNull(c.Net));
        Assert.All(cache.ZoneCandidates, c => Assert.NotNull(c.Net));
    }

    [Fact]
    public async Task Every_candidate_is_reachable_through_the_cell_index()
    {
        // The index is the whole point: an item missing from it can never be hovered.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        AssertAllIndexed(cache.SegmentCandidates.Count, cache.SegmentIndicesByCell, "segment");
        AssertAllIndexed(cache.PadCandidates.Count, cache.PadIndicesByCell, "pad");
        AssertAllIndexed(cache.ViaCandidates.Count, cache.ViaIndicesByCell, "via");
        AssertAllIndexed(cache.ZoneCandidates.Count, cache.ZoneIndicesByCell, "zone");

        static void AssertAllIndexed(int count, Dictionary<long, List<int>> index, string kind)
        {
            var indexed = index.Values.SelectMany(v => v).Distinct().ToHashSet();
            for (int i = 0; i < count; i++)
            {
                Assert.True(indexed.Contains(i), $"{kind} candidate {i} is not in the cell index");
            }
        }
    }

    [Fact]
    public async Task Every_indexed_position_refers_to_a_real_candidate()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        Assert.All(cache.SegmentIndicesByCell.Values.SelectMany(v => v),
            i => Assert.InRange(i, 0, cache.SegmentCandidates.Count - 1));
        Assert.All(cache.ZoneIndicesByCell.Values.SelectMany(v => v),
            i => Assert.InRange(i, 0, cache.ZoneCandidates.Count - 1));
    }

    [Fact]
    public async Task The_hit_radius_is_positive_so_thin_tracks_stay_clickable()
    {
        // A 0.25mm track is far thinner than the pointer; the radius is what makes it hoverable.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        Assert.All(cache.SegmentCandidates, c => Assert.True(c.HitRadiusWorld > 0));
        Assert.True(cache.MaxHitRadiusWorld > 0);
        Assert.True(cache.CellSizeWorld > 0);
    }

    [Fact]
    public async Task The_max_hit_radius_covers_every_candidate()
    {
        // The hover search widens by this much; anything larger could never be found.
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");

        foreach (var candidate in cache.SegmentCandidates)
        {
            Assert.True(candidate.HitRadiusWorld <= cache.MaxHitRadiusWorld);
        }

        foreach (var candidate in cache.PadCandidates)
        {
            Assert.True(candidate.HitRadiusWorld <= cache.MaxHitRadiusWorld);
        }
    }

    [Fact]
    public async Task A_zone_candidate_carries_its_polygons_and_bounds()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "F.Cu");
        var zone = Assert.Single(cache.ZoneCandidates);

        Assert.NotEmpty(zone.PolygonsWorld);
        Assert.True(zone.BoundsWorld.Width > 0);
        Assert.True(zone.BoundsWorld.Height > 0);
    }

    [Fact]
    public async Task A_board_with_no_copper_on_a_side_yields_an_empty_index()
    {
        KiCadPcb pcb = await this.LoadBoardAsync();

        var cache = KiCadHoverIndex.BuildKiCadPcbHoverHitTestCache(pcb, "In1.Cu");

        Assert.Empty(cache.SegmentCandidates);
        Assert.Empty(cache.ZoneCandidates);
    }
}
