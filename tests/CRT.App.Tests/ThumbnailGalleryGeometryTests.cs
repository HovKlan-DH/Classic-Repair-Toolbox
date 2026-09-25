using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// The auto-fit maths behind the detached schematics thumbnails window: given a thumbnail count
// and the window's available client area, pick the rows x columns grid and uniform cell size
// that shows every thumbnail as large as possible with no scrolling.
//
// Every tile gets the SAME cell size regardless of its own image's aspect ratio - a deliberate
// simplification over true aspect-aware bin packing (see ThumbnailGalleryGeometry's own header).
public class ThumbnailGalleryGeometryTests
{
    private const double Spacing = 8;
    private const double Margin = 8;

    [Fact]
    public void Zero_items_yields_an_empty_grid()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(0, new Size(800, 600), Spacing, Margin);

        Assert.Equal(0, shape.Rows);
        Assert.Equal(0, shape.Columns);
    }

    [Fact]
    public void Negative_item_count_yields_an_empty_grid()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(-3, new Size(800, 600), Spacing, Margin);

        Assert.Equal(0, shape.Rows);
        Assert.Equal(0, shape.Columns);
    }

    // A single thumbnail should fill the available area (minus margins), not be shrunk into a
    // fraction of a grid meant for many.
    [Fact]
    public void One_item_fills_the_available_area()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(1, new Size(800, 600), Spacing, Margin);
        Assert.Equal(1, shape.Rows);
        Assert.Equal(1, shape.Columns);

        var cell = ThumbnailGalleryGeometry.ComputeCellSize(shape, new Size(800, 600), Spacing, Margin);
        Assert.Equal(800 - (Margin * 2), cell.Width, 3);
        Assert.Equal(600 - (Margin * 2), cell.Height, 3);
    }

    // "Small window still shows all, big window shows bigger" - pinned against actual chosen grid
    // shapes for two contrasting window aspect ratios, not just that a shape exists.
    [Fact]
    public void A_wide_window_chooses_more_columns_than_rows_for_many_items()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(12, new Size(2000, 400), Spacing, Margin);

        Assert.True(shape.Columns > shape.Rows,
            $"Expected a wider-than-tall grid in a wide window, got {shape.Rows}x{shape.Columns}");
        Assert.True(shape.Rows * shape.Columns >= 12);
    }

    [Fact]
    public void A_tall_window_chooses_more_rows_than_columns_for_many_items()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(12, new Size(400, 2000), Spacing, Margin);

        Assert.True(shape.Rows > shape.Columns,
            $"Expected a taller-than-wide grid in a tall window, got {shape.Rows}x{shape.Columns}");
        Assert.True(shape.Rows * shape.Columns >= 12);
    }

    [Theory]
    [InlineData(20000, 50)]
    [InlineData(50, 20000)]
    public void An_extreme_aspect_ratio_window_still_covers_every_item(double width, double height)
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(15, new Size(width, height), Spacing, Margin);

        Assert.True(shape.Rows > 0);
        Assert.True(shape.Columns > 0);
        Assert.True(shape.Rows * shape.Columns >= 15);
    }

    // A tiny window must not blow up the maths - every thumbnail just becomes very small.
    [Fact]
    public void A_tiny_window_still_produces_valid_positive_cell_sizes()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(6, new Size(10, 10), Spacing, Margin);
        var cell = ThumbnailGalleryGeometry.ComputeCellSize(shape, new Size(10, 10), Spacing, Margin);

        Assert.True(cell.Width > 0);
        Assert.True(cell.Height > 0);
        Assert.False(double.IsNaN(cell.Width));
        Assert.False(double.IsNaN(cell.Height));
        Assert.False(double.IsInfinity(cell.Width));
        Assert.False(double.IsInfinity(cell.Height));
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(-100, -100)]
    public void A_degenerate_available_size_never_produces_nan_or_infinite_cells(double width, double height)
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(4, new Size(Math.Max(0, width), Math.Max(0, height)), Spacing, Margin);
        var cell = ThumbnailGalleryGeometry.ComputeCellSize(shape, new Size(width, height), Spacing, Margin);

        Assert.False(double.IsNaN(cell.Width));
        Assert.False(double.IsNaN(cell.Height));
        Assert.False(double.IsInfinity(cell.Width));
        Assert.False(double.IsInfinity(cell.Height));
    }

    // Row-first, left-to-right, top row first - the same reading order the existing vertical list
    // already presents thumbnails in, so switching layouts reflows but never reorders them.
    [Fact]
    public void ArrangeGrid_fills_row_first_left_to_right()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(2, 3);
        var cellSize = new Size(100, 80);

        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, cellSize, Spacing, Margin, 5);

        Assert.Equal(5, positions.Count);

        // Position 0 is top-left; position 1 is immediately to its right (same Y, greater X).
        Assert.Equal(positions[0].Y, positions[1].Y, 3);
        Assert.True(positions[1].X > positions[0].X);

        // Position 3 starts the second row: same X as position 0, greater Y.
        Assert.Equal(positions[0].X, positions[3].X, 3);
        Assert.True(positions[3].Y > positions[0].Y);
    }

    [Fact]
    public void ArrangeGrid_returns_exactly_itemCount_positions()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(3, 4);
        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, new Size(50, 50), Spacing, Margin, 10);

        Assert.Equal(10, positions.Count);
    }

    // The three individually-callable parts must agree with the one-call convenience method - a
    // regression guard against the three drifting the way ParkedBadgeGeometry's rows/columns once did.
    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    [InlineData(13)]
    public void BuildLayout_agrees_with_the_individual_calls(int itemCount)
    {
        var availableSize = new Size(900, 500);

        var shape = ThumbnailGalleryGeometry.ChooseGridShape(itemCount, availableSize, Spacing, Margin);
        var cellSize = ThumbnailGalleryGeometry.ComputeCellSize(shape, availableSize, Spacing, Margin);
        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, cellSize, Spacing, Margin, itemCount);

        var built = ThumbnailGalleryGeometry.BuildLayout(itemCount, availableSize, Spacing, Margin);

        Assert.Equal(shape, built.Shape);
        Assert.Equal(cellSize, built.CellSize);
        Assert.Equal(positions, built.Positions);
    }

    [Fact]
    public void FindNearestCellIndex_returns_the_cell_the_point_is_inside()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(2, 2);
        var cellSize = new Size(100, 100);
        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, cellSize, Spacing, Margin, 4);

        // Center of the second cell (index 1, top row right column).
        var pointInsideSecondCell = new Point(positions[1].X + 50, positions[1].Y + 50);

        int index = ThumbnailGalleryGeometry.FindNearestCellIndex(positions, cellSize, pointInsideSecondCell, 4);

        Assert.Equal(1, index);
    }

    [Fact]
    public void FindNearestCellIndex_picks_the_nearer_cell_when_the_point_is_in_a_gap()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(1, 2);
        var cellSize = new Size(100, 100);
        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, cellSize, Spacing, Margin, 2);

        // Just past the midpoint between the two cells' centers, leaning toward the second cell.
        double firstCenterX = positions[0].X + (cellSize.Width / 2);
        double secondCenterX = positions[1].X + (cellSize.Width / 2);
        var pointInGapNearSecond = new Point((firstCenterX + secondCenterX) / 2 + 1, positions[0].Y + 50);

        int index = ThumbnailGalleryGeometry.FindNearestCellIndex(positions, cellSize, pointInGapNearSecond, 2);

        Assert.Equal(1, index);
    }

    [Fact]
    public void FindNearestCellIndex_clamps_to_the_last_valid_index_past_the_grid()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(1, 3);
        var cellSize = new Size(100, 100);
        var positions = ThumbnailGalleryGeometry.ArrangeGrid(shape, cellSize, Spacing, Margin, 3);

        var pointFarPastLastCell = new Point(100000, 50);

        int index = ThumbnailGalleryGeometry.FindNearestCellIndex(positions, cellSize, pointFarPastLastCell, 3);

        Assert.Equal(2, index);
    }

    [Fact]
    public void FindNearestCellIndex_returns_negative_one_for_zero_items()
    {
        int index = ThumbnailGalleryGeometry.FindNearestCellIndex(Array.Empty<Point>(), new Size(100, 100), new Point(10, 10), 0);

        Assert.Equal(-1, index);
    }

    // ---------------------------------------------------------------------------------------
    // Unconstrained (infinite) available size
    // ---------------------------------------------------------------------------------------

    // An ancestor that measures without a constraint - a ScrollViewer with scrolling enabled, a
    // StackPanel, an auto-sized Window - hands an infinite axis down. Math.Max only floors the
    // SMALL side, so the division produced an Infinity cell, every child was measured against it,
    // and the panel returned an Infinity desired size, which Avalonia treats as a layout error.
    [Theory]
    [InlineData(double.PositiveInfinity, 600)]
    [InlineData(800, double.PositiveInfinity)]
    [InlineData(double.PositiveInfinity, double.PositiveInfinity)]
    public void An_infinite_available_axis_never_produces_an_infinite_cell(double width, double height)
    {
        var available = new Size(width, height);

        var shape = ThumbnailGalleryGeometry.ChooseGridShape(9, available, Spacing, Margin);
        var cell = ThumbnailGalleryGeometry.ComputeCellSize(shape, available, Spacing, Margin);

        Assert.False(double.IsInfinity(cell.Width));
        Assert.False(double.IsInfinity(cell.Height));
        Assert.False(double.IsNaN(cell.Width));
        Assert.False(double.IsNaN(cell.Height));
        Assert.True(cell.Width > 0);
        Assert.True(cell.Height > 0);
    }

    // A CONSTRAINED axis must still auto-fit normally when the OTHER one is infinite - clamping
    // both to the fallback would throw away the one real measurement available.
    [Fact]
    public void A_finite_axis_still_auto_fits_when_the_other_is_infinite()
    {
        var shape = new ThumbnailGalleryGeometry.GridShape(2, 2);
        var cell = ThumbnailGalleryGeometry.ComputeCellSize(
            shape, new Size(808, double.PositiveInfinity), Spacing, Margin);

        // (808 - 16 margin - 8 spacing) / 2 columns
        Assert.Equal(392, cell.Width, 3);
        Assert.False(double.IsInfinity(cell.Height));
    }

    // With both axes unconstrained every shape scores identically, and a "strictly better" scan
    // then keeps whatever it saw first - which is always one column, the single shape an auto-fit
    // gallery should never fall back to. The tie must break towards the squarest grid instead.
    [Fact]
    public void With_no_constraint_at_all_the_grid_is_square_ish_rather_than_a_single_column()
    {
        var shape = ThumbnailGalleryGeometry.ChooseGridShape(
            9, new Size(double.PositiveInfinity, double.PositiveInfinity), Spacing, Margin);

        Assert.Equal(3, shape.Columns);
        Assert.Equal(3, shape.Rows);
    }
}
