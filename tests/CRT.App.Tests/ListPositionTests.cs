using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where a list stands, as the other tabs' list and the Contribute tab's should both stand (ListPosition; owner request,
// 2026-11-09: "it should not change position in the lists").
// ###########################################################################################
public sealed class ListPositionTests
{
    // 10 rows of 20: an offset of 50 is row 2, half way into it.
    [Fact]
    public void The_top_row_and_how_far_into_it_are_read_off_the_offset()
    {
        Assert.Equal((2, 0.5), ListPosition.TopOf(50, 200, 10));
        Assert.Equal((0, 0.0), ListPosition.TopOf(0, 200, 10));
    }

    [Fact]
    public void A_list_with_nothing_to_measure_stands_nowhere()
    {
        Assert.Null(ListPosition.TopOf(10, 0, 10));
        Assert.Null(ListPosition.TopOf(10, 200, 0));
    }

    // An offset past the end - a list longer a moment ago - is its last row, never one past it.
    [Fact]
    public void An_offset_past_the_end_is_the_last_row()
    {
        Assert.Equal(9, ListPosition.TopOf(500, 200, 10)!.Value.Row);
    }

    [Fact]
    public void The_same_row_of_the_other_list_is_found_by_its_key_whatever_its_place()
    {
        Assert.Equal(1, ListPosition.MatchingRow(["U1", "U2", "U3"], 2, ["U2", "U3"]));
        Assert.Equal(2, ListPosition.MatchingRow(["U1", "U2", "U3"], 2, ["u1", "u2", "u3"]));
    }

    // A row the other list has not - narrowed there - stands for the first after it that it has; none after it, its last.
    [Fact]
    public void A_row_the_other_list_lacks_stands_for_the_next_it_has()
    {
        Assert.Equal(1, ListPosition.MatchingRow(["U1", "U2", "U3", "U4"], 1, ["U1", "U3"]));
        Assert.Equal(1, ListPosition.MatchingRow(["U1", "U2", "U3", "U4"], 3, ["U1", "U2"]));
        Assert.Equal(-1, ListPosition.MatchingRow(["U1"], 0, []));
    }

    [Fact]
    public void The_offset_puts_the_row_at_the_top_as_far_into_it_and_never_past_what_scrolls()
    {
        Assert.Equal(50, ListPosition.OffsetOf(2, 0.5, 200, 10, 100), 6);
        Assert.Equal(100, ListPosition.OffsetOf(9, 0, 200, 10, 100), 6);
        Assert.Equal(0, ListPosition.OffsetOf(3, 0, 200, 10, 300), 6);
    }
}
