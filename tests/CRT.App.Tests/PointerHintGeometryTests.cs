using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// PointerHintGeometry - a hint following the pointer (owner request, 2026-11-20: "have the hover/helper text follow the mouse"):
// below and right of it, never under it, and inside the room it is drawn in.
// ###########################################################################################
public sealed class PointerHintGeometryTests
{
    private static readonly Size Hint = new(130, 22);
    private static readonly Size Room = new(1000, 700);

    [Fact]
    public void The_hint_stands_below_and_right_of_the_pointer()
    {
        Assert.Equal(new Point(100 + PointerHintGeometry.Right, 200 + PointerHintGeometry.Below), PointerHintGeometry.Place(new Point(100, 200), Hint, Room));
    }

    // Near the right edge it moves left, near the bottom it goes above the pointer - never over the pointer itself.
    [Fact]
    public void Near_an_edge_the_hint_stays_inside_the_room_and_clear_of_the_pointer()
    {
        Point right = PointerHintGeometry.Place(new Point(950, 200), Hint, Room);
        Assert.Equal(Room.Width - PointerHintGeometry.Edge - Hint.Width, right.X);
        Assert.Equal(200 + PointerHintGeometry.Below, right.Y);

        Point bottom = PointerHintGeometry.Place(new Point(100, 690), Hint, Room);
        Assert.Equal(690 - PointerHintGeometry.Edge - Hint.Height, bottom.Y);
        Assert.True(bottom.Y + Hint.Height < 690, "the hint covers the pointer");
    }

    // A room smaller than the hint: from its top left corner, never off it.
    [Fact]
    public void In_a_room_smaller_than_the_hint_it_starts_at_the_corner()
    {
        Assert.Equal(new Point(0, 0), PointerHintGeometry.Place(new Point(10, 5), Hint, new Size(60, 10)));
    }
}
