using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// SideListsGeometry - the category list on the far left leaves the component list its half of the room (2026-11-25:
// on CRT's smallest window the C64 250407's seven categories took it all, the component list 9 high).
// ###########################################################################################
public sealed class SideListsGeometryTests
{
    // The categories may have half of what the two lists share - on a tall window far more than their rows, so nothing
    // changes there.
    [Theory]
    [InlineData(600, 300)]
    [InlineData(100, 50)]
    public void The_category_list_may_have_half_the_room(double room, double expected)
    {
        Assert.Equal(expected, SideListsGeometry.CategoryListMaxHeight(room));
    }

    // However little room, two of its rows - a list that shows less is no list.
    [Fact]
    public void The_category_list_keeps_two_rows_however_little_room()
    {
        Assert.Equal(SideListsGeometry.CategoryListLeast, SideListsGeometry.CategoryListMaxHeight(30));
    }

    // A room not laid out yet limits nothing - the category list is not shut by a layout pass still to come.
    [Theory]
    [InlineData(0)]
    [InlineData(-10)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void A_room_not_laid_out_yet_limits_nothing(double room)
    {
        Assert.True(double.IsPositiveInfinity(SideListsGeometry.CategoryListMaxHeight(room)));
    }
}
