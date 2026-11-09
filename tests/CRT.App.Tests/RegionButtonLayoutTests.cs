using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where a board's region buttons stand (RegionButtonLayout; 2026-11-09: a board may come in regions of its own). PAL and NTSC
// are the two halves they always were. The grids are fixed - six equal columns 6px apart under the component list, two Auto
// columns on the popup - as Avalonia's Grid threw on a column it had lost (owner report, 2026-11-09).
// ###########################################################################################
public sealed class RegionButtonLayoutTests
{
    // Two buttons span three of the six columns each: with 6px between every column, each is (width - 6) / 2 - exactly the
    // halves "*,6,*" gave.
    [Fact]
    public void Pal_and_ntsc_are_two_halves_as_they_always_were()
    {
        Assert.Equal((2, 1), RegionButtonLayout.Of(2));
        Assert.Equal((0, 0, 3), RegionButtonLayout.PlaceOf(0, 2));
        Assert.Equal((0, 3, 3), RegionButtonLayout.PlaceOf(1, 2));

        const double width = 300;
        double column = (width - ((RegionButtonLayout.GridColumns - 1) * RegionButtonLayout.Gap)) / RegionButtonLayout.GridColumns;
        double half = (3 * column) + (2 * RegionButtonLayout.Gap);
        Assert.Equal((width - RegionButtonLayout.Gap) / 2, half, 6);
    }

    // Three to a row at most, spanning two columns each, the next under them - four two by two, never one alone a third as
    // wide.
    [Fact]
    public void More_than_three_go_on_to_a_row_of_their_own()
    {
        Assert.Equal((2, 2), RegionButtonLayout.Of(4));
        Assert.Equal((1, 0, 3), RegionButtonLayout.PlaceOf(2, 4));
        Assert.Equal((1, 3, 3), RegionButtonLayout.PlaceOf(3, 4));

        Assert.Equal((3, 2), RegionButtonLayout.Of(5));
        Assert.Equal((0, 4, 2), RegionButtonLayout.PlaceOf(2, 5));
        Assert.Equal((1, 0, 2), RegionButtonLayout.PlaceOf(3, 5));
        Assert.Equal((1, 2, 2), RegionButtonLayout.PlaceOf(4, 5));

        Assert.Equal((1, 4, 2), RegionButtonLayout.PlaceOf(5, 6));
    }

    // Every place stands inside the fixed grid: six columns, two rows - whatever the count up to the most a board has.
    [Fact]
    public void Every_place_is_inside_the_fixed_grid()
    {
        for (int count = 1; count <= 6; count++)
        {
            for (int index = 0; index < count; index++)
            {
                (int row, int column, int span) = RegionButtonLayout.PlaceOf(index, count);

                Assert.InRange(row, 0, 1);
                Assert.True(column >= 0 && column + span <= RegionButtonLayout.GridColumns, $"{index} of {count}: {column}+{span}");
            }
        }
    }

    // The popup stacks them in its footer's three rows: PAL and NTSC in the second and third, as they always stood; a fourth
    // starts its second column.
    [Fact]
    public void The_popup_stacks_them_from_the_bottom_with_pal_and_ntsc_where_they_always_stood()
    {
        Assert.Equal((2, 0), RegionButtonLayout.PopupCellOf(0, 2));
        Assert.Equal((4, 0), RegionButtonLayout.PopupCellOf(1, 2));
        Assert.Equal((4, 0), RegionButtonLayout.PopupCellOf(0, 1));
        Assert.Equal((0, 0), RegionButtonLayout.PopupCellOf(0, 3));
        Assert.Equal((4, 1), RegionButtonLayout.PopupCellOf(5, 6));

        // The second column from the bottom too (code review, 2026-11-09: a fourth region stood at the top, beside the
        // first, an empty space under it).
        Assert.Equal((4, 1), RegionButtonLayout.PopupCellOf(3, 4));
        Assert.Equal((2, 1), RegionButtonLayout.PopupCellOf(3, 5));
        Assert.Equal((4, 1), RegionButtonLayout.PopupCellOf(4, 5));
        Assert.Equal((0, 1), RegionButtonLayout.PopupCellOf(3, 6));
    }

    [Fact]
    public void One_region_is_one_button_across_and_none_is_nothing()
    {
        Assert.Equal((1, 1), RegionButtonLayout.Of(1));
        Assert.Equal((0, 0, 6), RegionButtonLayout.PlaceOf(0, 1));
        Assert.Equal((0, 0), RegionButtonLayout.Of(0));
    }
}
