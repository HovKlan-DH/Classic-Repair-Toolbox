using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// RowDragSlots - where a dragged row goes, for the worklog editor's Photos and Files lists and the
// Drafts tab's "Schematic images" window (2026-09-27).
//
// The rule these pin is the fix for rows that "rapidly switch position and cannot settle": the
// slots are the row MIDPOINTS AS THEY WERE WHEN THE DRAG STARTED, so the answer depends on the
// pointer alone. A test that re-measured after each move would be testing the bug.
// ###########################################################################################
public sealed class RowDragSlotsTests
{
    // Three rows of different heights, stacked from 0 with a 4px gap: 0-60, 64-164, 168-208.
    private static readonly (double? Top, double Height)[] ThreeRows = [(0, 60), (64, 100), (168, 40)];

    [Fact]
    public void The_slots_are_the_middle_of_each_row()
    {
        Assert.Equal([30.0, 114.0, 188.0], RowDragSlots.BuildMidpoints(ThreeRows));
    }

    // ###########################################################################################
    // A row whose container is not realised has no top to read. It is placed straight after the
    // previous row rather than skipped - skipping would shift every later slot by one, and the drop
    // would land one row away from where it was shown.
    // ###########################################################################################
    [Fact]
    public void A_row_with_no_measured_top_is_placed_after_the_previous_one_and_keeps_its_index()
    {
        List<double> midpoints = RowDragSlots.BuildMidpoints([(0, 60), (null, 100), (170, 40)]);

        Assert.Equal(3, midpoints.Count);
        Assert.Equal(110.0, midpoints[1]);
        Assert.Equal(190.0, midpoints[2]);
    }

    // Dragging the TOP row (index 0) down. Its own middle (30) is never counted.
    [Theory]
    [InlineData(-500, 0)]   // dragged above the list entirely
    [InlineData(10, 0)]
    [InlineData(35, 0)]     // past its OWN middle, still over itself - it stays put
    [InlineData(113.9, 0)]
    [InlineData(114, 1)]    // past the second row's middle - below it
    [InlineData(187.9, 1)]
    [InlineData(188, 2)]    // past the last middle is the last index...
    [InlineData(5000, 2)]   // ...however far below the list the pointer goes
    public void Dragging_down_counts_the_other_rows_whose_middle_the_pointer_is_past(double pointerY, int expected)
    {
        Assert.Equal(expected, RowDragSlots.ResolveDropIndex(RowDragSlots.BuildMidpoints(ThreeRows), pointerY, draggedIndex: 0));
    }

    // Dragging the BOTTOM row (index 2) up.
    [Theory]
    [InlineData(5000, 2)]
    [InlineData(150, 2)]    // below the second row's middle - still last
    [InlineData(113.9, 1)]  // above it - before it
    [InlineData(29.9, 0)]
    [InlineData(-500, 0)]
    public void Dragging_up_counts_the_other_rows_whose_middle_the_pointer_is_past(double pointerY, int expected)
    {
        Assert.Equal(expected, RowDragSlots.ResolveDropIndex(RowDragSlots.BuildMidpoints(ThreeRows), pointerY, draggedIndex: 2));
    }

    // ###########################################################################################
    // *** THE OFF-BY-ONE, PINNED (2026-09-27). *** The worklog's first version counted the dragged
    // row's own middle, so a row dragged a few pixels down past its own middle - still over itself
    // - was already put BELOW the next row, and the placeholder ran a row ahead of the pointer all
    // the way down. The answer is an index in the list without the dragged row, which is what
    // ObservableCollection.Move expects.
    // ###########################################################################################
    [Fact]
    public void A_row_nudged_down_past_its_own_middle_does_not_jump_below_the_next_one()
    {
        List<double> midpoints = RowDragSlots.BuildMidpoints(ThreeRows);

        Assert.Equal(0, RowDragSlots.ResolveDropIndex(midpoints, 40, draggedIndex: 0));
    }

    // ###########################################################################################
    // *** THE OSCILLATION, PINNED. *** Against frozen slots, a pointer held still - jiggling by a
    // pixel, as a hand does - gives one answer every time. The live version answered differently
    // after each move because each move changed what it measured.
    // ###########################################################################################
    [Fact]
    public void A_pointer_held_still_gives_the_same_slot_every_time()
    {
        List<double> midpoints = RowDragSlots.BuildMidpoints(ThreeRows);

        int[] answers = Enumerable.Range(0, 20)
            .Select(i => RowDragSlots.ResolveDropIndex(midpoints, 150 + (i % 3), draggedIndex: 0))
            .ToArray();

        Assert.All(answers, answer => Assert.Equal(1, answer));
    }

    [Fact]
    public void No_rows_means_no_slot()
    {
        Assert.Equal(-1, RowDragSlots.ResolveDropIndex([], 10, draggedIndex: 0));
        Assert.Empty(RowDragSlots.BuildMidpoints([]));
    }

    // ------------------------------------------------------------------ auto-scroll

    [Theory]
    [InlineData(200)]
    [InlineData(40)]    // exactly on the band's inner edge
    [InlineData(360)]
    public void A_pointer_away_from_both_edges_does_not_scroll(double pointerY)
    {
        Assert.Equal(0, RowDragSlots.EdgeScrollDelta(pointerY, viewportHeight: 400, band: 40, maxStep: 18));
    }

    [Fact]
    public void Near_the_top_it_scrolls_up_and_near_the_bottom_down()
    {
        Assert.True(RowDragSlots.EdgeScrollDelta(10, 400, 40, 18) < 0);
        Assert.True(RowDragSlots.EdgeScrollDelta(390, 400, 40, 18) > 0);
    }

    // Resting just inside the band creeps; at or past the edge it runs at full speed, and never
    // faster - so pushing far outside the window cannot fling the list.
    [Fact]
    public void Deeper_into_the_edge_scrolls_faster_up_to_the_most_one_step_may()
    {
        double creep = RowDragSlots.EdgeScrollDelta(35, 400, 40, 18);
        double edge = RowDragSlots.EdgeScrollDelta(0, 400, 40, 18);
        double farOutside = RowDragSlots.EdgeScrollDelta(-900, 400, 40, 18);

        Assert.True(Math.Abs(creep) < Math.Abs(edge));
        Assert.Equal(-18, edge);
        Assert.Equal(-18, farOutside);
        Assert.Equal(18, RowDragSlots.EdgeScrollDelta(1300, 400, 40, 18));
    }

    // Even the gentlest scroll moves at least a pixel, or resting in the band would do nothing.
    [Fact]
    public void Inside_the_band_it_always_moves_at_least_a_pixel()
    {
        Assert.Equal(-1, RowDragSlots.EdgeScrollDelta(39.99, 400, 40, 18));
    }

    // On a short viewport a 40px band at each end could meet in the middle; capped at a third of
    // the height, the two can never both claim the same pointer.
    [Fact]
    public void On_a_short_viewport_the_bands_never_overlap()
    {
        Assert.Equal(0, RowDragSlots.EdgeScrollDelta(30, viewportHeight: 60, band: 40, maxStep: 18));
        Assert.True(RowDragSlots.EdgeScrollDelta(10, 60, 40, 18) < 0);
        Assert.True(RowDragSlots.EdgeScrollDelta(50, 60, 40, 18) > 0);
    }

    [Fact]
    public void Nothing_to_scroll_without_a_viewport()
    {
        Assert.Equal(0, RowDragSlots.EdgeScrollDelta(0, 0, 40, 18));
    }
}
