using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Where the worklog pills go when their entry has no marked area to sit on.
//
// An entry with "Show marked area" unticked draws no rectangle, so its "#N" pill has nothing to
// anchor to. It parks in the schematic panel's top-right corner instead of being hidden - hiding
// it too would make the entry invisible on the board and unreachable without opening the worklog
// list.
//
// They are arranged as a compact BLOCK, not one long column. A single column was the first
// implementation and it ran down over the middle of the board, hiding more of it the more entries
// there were.
//
// The block fills ROW-FIRST: five pills sit 3-over-2, not 3-down-then-2. That is both what reading
// order expects and the flatter result - a part-filled last ROW leaves the block as short as it can
// be, where a part-filled last COLUMN would leave it as tall as a full one. Measured: five pills
// are 50px tall this way, against 78px column-first and 134px as a single column.
public class ParkedBadgeGeometryTests
{
    private static readonly Size Viewport = new(800, 600);

    private const double Margin = 10;
    private const double Spacing = 6;

    private static List<Size> Sizes(params double[] widths)
    {
        var sizes = new List<Size>();
        foreach (double w in widths)
        {
            sizes.Add(new Size(w, 20));
        }
        return sizes;
    }

    private static List<Size> UniformSizes(int count, double width = 50, double height = 20)
    {
        var sizes = new List<Size>();
        for (int i = 0; i < count; i++)
        {
            sizes.Add(new Size(width, height));
        }
        return sizes;
    }

    private static List<Point> Arrange(IReadOnlyList<Size> sizes, double reservedRight = 0) =>
        ParkedBadgeGeometry.ArrangeInTopRightBlock(sizes, Viewport, Margin, Spacing, reservedRight);

    // ------------------------------------------------------------- the grid progression

    // The shape: one column before a second appears, then the column count grows square-ish. Two
    // pills side by side read as two unrelated things; one above the other reads as a list. Past
    // that, square-ish keeps the block compact - a tall column covers the board and a wide row runs
    // into the "Netlist names" panel and the thumbnails.
    //
    // Rows are DERIVED from the count and the columns, not the square's side, because the block
    // fills row-first and a part-filled last row needs no extra row of space: 5 pills in 3 columns
    // occupy 2 rows, not 3. This table previously asserted the square's side (3 for 5 pills) while
    // the layout produced 2 - GetGridShape's Rows was discarded by the caller and nothing noticed
    // the two disagreed.
    [Theory]
    [InlineData(1, 1, 1)]
    [InlineData(2, 2, 1)]
    [InlineData(3, 2, 2)]
    [InlineData(4, 2, 2)]
    [InlineData(5, 2, 3)]
    [InlineData(9, 3, 3)]
    [InlineData(10, 3, 4)]
    [InlineData(16, 4, 4)]
    [InlineData(17, 4, 5)]
    public void The_grid_grows_one_column_then_square(int badgeCount, int expectedRows, int expectedColumns)
    {
        var (rows, columns) = ParkedBadgeGeometry.GetGridShape(badgeCount);

        Assert.Equal(expectedRows, rows);
        Assert.Equal(expectedColumns, columns);
    }

    // The reported shape must be the shape actually laid out - the divergence above was invisible
    // because nothing compared the two.
    [Theory]
    [InlineData(3)]
    [InlineData(5)]
    [InlineData(10)]
    [InlineData(17)]
    public void The_reported_grid_matches_the_arrangement(int badgeCount)
    {
        var (rows, columns) = ParkedBadgeGeometry.GetGridShape(badgeCount);
        var positions = Arrange(UniformSizes(badgeCount));

        Assert.Equal(rows, positions.Select(p => Math.Round(p.Y, 3)).Distinct().Count());
        Assert.Equal(columns, positions.Select(p => Math.Round(p.X, 3)).Distinct().Count());
    }

    [Fact]
    public void No_pills_produces_no_grid()
    {
        Assert.Equal((0, 0), ParkedBadgeGeometry.GetGridShape(0));
    }

    // The grid must always be big enough to hold every pill - the property the table above is
    // really asserting, checked across the whole range rather than at the sampled points.
    [Fact]
    public void The_grid_always_has_room_for_every_pill()
    {
        for (int count = 1; count <= 40; count++)
        {
            var (rows, columns) = ParkedBadgeGeometry.GetGridShape(count);

            Assert.True(rows * columns >= count, $"{count} pills do not fit in {rows}x{columns}");
        }
    }

    // ------------------------------------------------------------- placement

    [Fact]
    public void The_first_pill_sits_in_the_top_right_corner_inside_the_margin()
    {
        var positions = Arrange(Sizes(50));

        var only = Assert.Single(positions);
        Assert.Equal(Viewport.Width - Margin - 50, only.X, 6);
        Assert.Equal(Margin, only.Y, 6);
    }

    // Row-first: the next pill goes BESIDE the first (to its left), not below it. Two pills is the
    // one case that is still a single column by design, so this uses three.
    [Fact]
    public void The_next_pill_goes_beside_the_first_not_below_it()
    {
        var positions = Arrange(UniformSizes(3));

        Assert.Equal(positions[0].Y, positions[1].Y, 6);
        Assert.True(positions[1].X < positions[0].X, $"the second pill at {positions[1].X} is not left of the first at {positions[0].X}");
    }

    // Only once a row is full does the next start beneath it.
    [Fact]
    public void A_full_row_starts_the_next_one_below()
    {
        // 3 pills -> 2 columns, so the third wraps to a second row.
        var positions = Arrange(UniformSizes(3));

        Assert.Equal(positions[0].Y + 20 + Spacing, positions[2].Y, 6);
        Assert.Equal(positions[0].X, positions[2].X, 6);
    }

    // THE case from the sketch: five pills sit three-over-two, not three-down-then-two.
    [Fact]
    public void Five_pills_sit_three_over_two()
    {
        var positions = Arrange(UniformSizes(5));

        double topRow = positions[0].Y;
        double secondRow = positions[4].Y;

        Assert.Equal(3, positions.Count(p => Math.Abs(p.Y - topRow) < 0.001));
        Assert.Equal(2, positions.Count(p => Math.Abs(p.Y - secondRow) < 0.001));
        Assert.True(secondRow > topRow, "the second row is not below the first");
    }

    // Within a row the pills run RIGHT to left, so #1 stays in the corner the block is anchored to
    // and the block grows towards the board rather than off the edge.
    [Fact]
    public void Pills_in_a_row_run_right_to_left_from_the_corner()
    {
        var positions = Arrange(UniformSizes(3));

        Assert.True(positions[0].X > positions[1].X, "#1 is not the rightmost pill in its row");
    }

    // Only the rows actually needed. Five pills in a 3x3 grid occupy TWO rows, so the block must
    // not be padded out with an empty third one.
    [Fact]
    public void A_part_filled_grid_uses_only_the_rows_it_needs()
    {
        var positions = Arrange(UniformSizes(5));

        int distinctRows = positions.Select(p => Math.Round(p.Y, 3)).Distinct().Count();

        Assert.Equal(2, distinctRows);
    }

    // No pill may overlap another - the point of the arrangement, asserted directly rather than
    // inferred from the coordinates above.
    [Theory]
    [InlineData(2)]
    [InlineData(5)]
    [InlineData(9)]
    [InlineData(16)]
    public void No_two_pills_overlap(int badgeCount)
    {
        var sizes = UniformSizes(badgeCount);
        var positions = Arrange(sizes);

        for (int i = 0; i < positions.Count; i++)
        {
            for (int j = i + 1; j < positions.Count; j++)
            {
                var a = new Rect(positions[i], sizes[i]);
                var b = new Rect(positions[j], sizes[j]);

                Assert.False(a.Intersects(b), $"pill {i} at {a} overlaps pill {j} at {b}");
            }
        }
    }

    // The block is what makes this worth doing: many pills must be shorter than a single column of
    // the same pills would have been, or they still run down over the board.
    [Theory]
    [InlineData(5)]
    [InlineData(9)]
    [InlineData(16)]
    public void A_block_is_shorter_than_a_column_of_the_same_pills(int badgeCount)
    {
        var positions = Arrange(UniformSizes(badgeCount));

        double blockBottom = positions.Max(p => p.Y) + 20;
        double columnBottom = Margin + (badgeCount * 20) + ((badgeCount - 1) * Spacing);

        Assert.True(blockBottom < columnBottom, $"the block reaches {blockBottom}, no better than a column's {columnBottom}");
    }

    // Each column is only as wide as its OWN widest pill, so a narrow column does not reserve the
    // width of a wide pill sitting in a different one.
    [Fact]
    public void A_column_is_only_as_wide_as_its_own_widest_pill()
    {
        // 2 columns, filled row-first: indices 0 and 2 land in the right column, 1 and 3 in the
        // left. Wide pills on the right, narrow on the left.
        var sizes = Sizes(80, 30, 80, 30);
        var positions = Arrange(sizes);

        double rightColumnLeft = positions[0].X;
        double narrowColumnRight = positions[1].X + 30;

        // The narrow column ends exactly one spacing left of the wide column's left edge.
        Assert.Equal(rightColumnLeft - Spacing, narrowColumnRight, 6);
    }

    // Right-aligned within a column, not left-aligned: a narrow "#7" and a wide "#12" in the same
    // column must end at the same x, so the block has a straight right edge.
    [Fact]
    public void Pills_in_a_column_share_a_right_edge()
    {
        // 2 columns: indices 0 and 2 share the right column.
        var sizes = Sizes(40, 50, 70, 50);
        var positions = Arrange(sizes);

        Assert.Equal(positions[0].X + 40, positions[2].X + 70, 6);
    }

    // ------------------------------------------------------------- the "Netlist names" panel

    [Fact]
    public void An_open_panel_pushes_the_whole_block_left_by_its_reserved_width()
    {
        var without = Arrange(UniformSizes(5));
        var with = Arrange(UniformSizes(5), reservedRight: 180);

        for (int i = 0; i < without.Count; i++)
        {
            Assert.Equal(without[i].X - 180, with[i].X, 6);
        }
    }

    [Fact]
    public void Closing_the_panel_returns_the_block_to_the_corner()
    {
        var moved = Arrange(UniformSizes(3), reservedRight: 180);
        var returned = Arrange(UniformSizes(3));

        Assert.NotEqual(moved[0].X, returned[0].X);
        Assert.Equal(Viewport.Width - Margin - 50, returned[0].X, 6);
    }

    // The reservation only moves the block sideways - a panel opening must not shunt it down.
    [Fact]
    public void The_reservation_does_not_change_the_vertical_arrangement()
    {
        var without = Arrange(UniformSizes(5));
        var with = Arrange(UniformSizes(5), reservedRight: 180);

        for (int i = 0; i < without.Count; i++)
        {
            Assert.Equal(without[i].Y, with[i].Y, 6);
        }
    }

    [Fact]
    public void A_negative_reservation_is_ignored()
    {
        var positions = Arrange(Sizes(50), reservedRight: -100);

        Assert.Equal(Viewport.Width - Margin - 50, positions[0].X, 6);
    }

    // ------------------------------------------------------------- edges

    // A pill too wide for what is left is pinned to the left margin, not pushed off it: unreadable
    // pills at a sensible position beat correctly-sized ones off-screen.
    [Fact]
    public void A_pill_wider_than_the_space_left_is_pinned_to_the_left_margin()
    {
        var positions = Arrange(Sizes(600), reservedRight: 300);

        Assert.Equal(Margin, positions[0].X, 6);
    }

    [Fact]
    public void No_pills_produces_no_positions()
    {
        Assert.Empty(Arrange(new List<Size>()));
    }

    [Fact]
    public void A_null_list_produces_no_positions()
    {
        Assert.Empty(ParkedBadgeGeometry.ArrangeInTopRightBlock(null!, Viewport, Margin, Spacing, 0));
    }

    // One position per pill, in the order given - the caller matches them up by index.
    [Fact]
    public void Every_pill_gets_exactly_one_position()
    {
        var positions = Arrange(UniformSizes(7));

        Assert.Equal(7, positions.Count);
    }

    // ------------------------------------------------------------- the bottom edge

    // The parked canvas lives inside the clipped schematic container, so a pill positioned below
    // the viewport is not merely awkward - it is clipped away entirely and becomes invisible AND
    // unclickable, the exact failure parking exists to prevent. Measured: 13 pills overflow a
    // 120px-tall panel, which a dragged splitter reaches easily.
    [Theory]
    [InlineData(13, 120.0)]
    [InlineData(25, 120.0)]
    [InlineData(43, 200.0)]
    public void No_pill_is_placed_below_a_short_viewport(int badgeCount, double viewportHeight)
    {
        var viewport = new Size(800, viewportHeight);
        var sizes = UniformSizes(badgeCount);

        var positions = ParkedBadgeGeometry.ArrangeInTopRightBlock(sizes, viewport, Margin, Spacing, 0);

        for (int i = 0; i < positions.Count; i++)
        {
            double bottom = positions[i].Y + sizes[i].Height;

            Assert.True(bottom <= viewportHeight + 0.001, $"pill {i} reaches {bottom}, past the {viewportHeight}px viewport");
        }
    }

    // A viewport too short for even one pill still places it at the margin rather than at a
    // negative coordinate - pinned, not pushed, exactly as the left edge behaves.
    [Fact]
    public void A_viewport_shorter_than_a_pill_still_pins_it_to_the_top()
    {
        var positions = ParkedBadgeGeometry.ArrangeInTopRightBlock(UniformSizes(1), new Size(800, 15), Margin, Spacing, 0);

        Assert.Equal(Margin, positions[0].Y, 6);
    }

    // An unmeasured viewport gives no basis for clamping, so the block simply lays out.
    [Fact]
    public void An_unmeasured_viewport_height_does_not_clamp()
    {
        var positions = ParkedBadgeGeometry.ArrangeInTopRightBlock(UniformSizes(6), new Size(800, 0), Margin, Spacing, 0);

        Assert.Equal(6, positions.Count);
        Assert.True(positions.Select(pos => Math.Round(pos.Y, 3)).Distinct().Count() > 1, "the block collapsed to one row");
    }
}
