using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ColumnAutoFitGeometry - double-clicking a board table column heading's edge fits the column to its
// text, as in Excel (owner request, 2026-10-04). The table measures (BoardTableEditorTextWrapTests
// drives it with a real pointer); these pin what it decides from the measurements.
// ###########################################################################################
public sealed class ColumnAutoFitGeometryTests
{
    // The right edge is the heading's own column, the left edge the column before it - where the
    // grid shows its resize cursor, 5px either side. (By name: ColumnEdge is internal, and a public
    // test method cannot take it.)
    [Theory]
    [InlineData(100, "ThisColumn")]
    [InlineData(95, "ThisColumn")]
    [InlineData(94, "None")]
    [InlineData(50, "None")]
    [InlineData(6, "None")]
    [InlineData(5, "PreviousColumn")]
    [InlineData(0, "PreviousColumn")]
    public void A_press_near_a_headings_edge_belongs_to_the_column_that_edge_resizes(double x, string expected)
    {
        Assert.Equal(expected, ColumnAutoFitGeometry.EdgeAt(x, headerWidth: 100).ToString());
    }

    // A heading narrower than both grips together answers for itself, as the grid decides it - its
    // right edge is checked first. 4px into an 8px heading is within both.
    [Fact]
    public void On_a_heading_narrower_than_both_grips_the_right_edge_wins()
    {
        Assert.Equal(ColumnEdge.ThisColumn, ColumnAutoFitGeometry.EdgeAt(4, headerWidth: 8));
    }

    [Fact]
    public void A_heading_with_no_width_has_no_edge()
    {
        Assert.Equal(ColumnEdge.None, ColumnAutoFitGeometry.EdgeAt(0, headerWidth: 0));
        Assert.Equal(ColumnEdge.None, ColumnAutoFitGeometry.EdgeAt(double.NaN, headerWidth: 100));
    }

    // ###########################################################################################
    // The widest text plus the cell's own margins, so it fits on one line - with a pixel to spare,
    // since a column exactly as wide as its text can still wrap the last word once the layout
    // rounds - and in whole pixels.
    // ###########################################################################################
    [Fact]
    public void The_column_fits_its_widest_text_with_the_cells_own_margins()
    {
        double width = ColumnAutoFitGeometry.FitWidth(headerWidth: 50, [40, 180.2, 90], cellChrome: 24, minWidth: 60, maxWidth: double.PositiveInfinity);

        Assert.Equal(206, width);
    }

    // A heading wider than every text keeps its name readable.
    [Fact]
    public void A_heading_wider_than_every_text_sets_the_width()
    {
        double width = ColumnAutoFitGeometry.FitWidth(headerWidth: 150, [10, 20], cellChrome: 24, minWidth: 60, maxWidth: double.PositiveInfinity);

        Assert.Equal(150, width);
    }

    // ###########################################################################################
    // NEVER WIDER THAN THE TABLE CAN SHOW: the cells wrap, so a column as wide as the table shows all
    // of a long text - one wider would push the end of every line past the right edge.
    // ###########################################################################################
    [Fact]
    public void A_text_wider_than_the_table_fits_the_table_and_wraps_the_rest()
    {
        double width = ColumnAutoFitGeometry.FitWidth(headerWidth: 50, [2400], cellChrome: 24, minWidth: 60, maxWidth: 812);

        Assert.Equal(812, width);
    }

    // Empty cells and a column of nothing keep the column's minimum - it never shrinks to a sliver.
    [Fact]
    public void A_column_of_empty_cells_keeps_its_minimum_width()
    {
        Assert.Equal(60, ColumnAutoFitGeometry.FitWidth(headerWidth: 20, [0, 0], cellChrome: 24, minWidth: 60, maxWidth: double.PositiveInfinity));
        Assert.Equal(60, ColumnAutoFitGeometry.FitWidth(headerWidth: 20, [], cellChrome: 24, minWidth: 60, maxWidth: double.PositiveInfinity));
    }

    // A table squeezed narrower than a column's minimum still gives the minimum - the column's own
    // floor wins over the table's room.
    [Fact]
    public void A_table_narrower_than_the_columns_minimum_still_gives_the_minimum()
    {
        Assert.Equal(60, ColumnAutoFitGeometry.FitWidth(headerWidth: 50, [500], cellChrome: 24, minWidth: 60, maxWidth: 30));
    }
}
