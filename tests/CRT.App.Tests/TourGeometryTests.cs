using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TourGeometry - where an introduction tour's frame, card and pointer stand (owner request, 2026-11-11). What must hold:
// the frame is seen whole, the card never covers what its step frames when there is room beside it, and nothing is ever
// past the window's edge.
// ###########################################################################################
public sealed class TourGeometryTests
{
    private static readonly Size Window = new(1500, 920);

    private static readonly Size Card = new(340, 180);

    private static void AssertInside(Rect inner, Size area)
    {
        Assert.True(inner.X >= 0 && inner.Y >= 0 && inner.Right <= area.Width && inner.Bottom <= area.Height, $"{inner} is not inside {area}");
    }

    [Fact]
    public void The_hole_is_the_target_a_little_larger()
    {
        Rect? hole = TourGeometry.HoleFor(new Rect(100, 200, 50, 40), Window);

        Assert.Equal(new Rect(94, 194, 62, 52), hole);
    }

    // The thumbnails stand against the window's right edge and foot: their frame would be drawn half outside it.
    [Fact]
    public void The_hole_never_goes_past_the_windows_edge()
    {
        Rect hole = TourGeometry.HoleFor(new Rect(1200, 100, 300, 820), Window)!.Value;

        Assert.Equal(Window.Width - TourGeometry.WindowEdge, hole.Right);
        Assert.Equal(Window.Height - TourGeometry.WindowEdge, hole.Bottom);
    }

    [Fact]
    public void No_target_or_one_wholly_outside_the_window_is_no_hole()
    {
        Assert.Null(TourGeometry.HoleFor(null, Window));
        Assert.Null(TourGeometry.HoleFor(new Rect(2000, 100, 50, 50), Window));
        Assert.Null(TourGeometry.HoleFor(new Rect(10, 10, 50, 50), new Size(0, 0)));
    }

    [Fact]
    public void Without_a_hole_the_card_stands_in_the_middle()
    {
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(null, Card, Window);

        Assert.Equal(TourCardSide.Centre, side);
        Assert.Equal((Window.Width - Card.Width) / 2, card.X);
        AssertInside(card, Window);
    }

    // The "Hardware" and "Board" drop-downs in the left panel: room on the right.
    [Fact]
    public void A_small_target_on_the_left_gets_its_card_on_the_right_level_with_it()
    {
        var hole = new Rect(10, 17, 190, 130);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(TourCardSide.Right, side);
        Assert.Equal(hole.Right + TourGeometry.Gap, card.X);
        Assert.False(card.Intersects(hole));
        AssertInside(card, Window);
    }

    // A tall target - the component list - has its card level with its top, not half way down it.
    [Fact]
    public void Beside_a_tall_target_the_card_stands_level_with_its_top()
    {
        var hole = new Rect(10, 140, 190, 680);
        (Rect card, _) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(hole.Y, card.Y);
    }

    // The thumbnails against the right edge: no room on the right, so on the left.
    [Fact]
    public void A_target_against_the_right_edge_gets_its_card_on_the_left()
    {
        var hole = new Rect(1190, 90, 307, 827);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(TourCardSide.Left, side);
        Assert.Equal(hole.X - TourGeometry.Gap, card.Right);
        Assert.False(card.Intersects(hole));
    }

    // A tab's header at the top, as wide as the window leaves no room on either side: below it.
    [Fact]
    public void A_wide_target_at_the_top_gets_its_card_below()
    {
        var hole = new Rect(20, 10, 1460, 60);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(TourCardSide.Below, side);
        Assert.Equal(hole.Bottom + TourGeometry.Gap, card.Y);
        Assert.False(card.Intersects(hole));
    }

    [Fact]
    public void A_wide_target_at_the_foot_gets_its_card_above()
    {
        var hole = new Rect(20, 800, 1460, 100);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(TourCardSide.Above, side);
        Assert.Equal(hole.Y - TourGeometry.Gap, card.Bottom);
    }

    // The picture fills the window: the card goes inside it, lower right, clear of the buttons along its foot.
    [Fact]
    public void A_target_with_no_room_around_it_gets_its_card_inside_clear_of_its_foot()
    {
        var hole = new Rect(200, 90, 1000, 820);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        Assert.Equal(TourCardSide.Inside, side);
        Assert.Equal(hole.Right - TourGeometry.InsideRight, card.Right);
        Assert.Equal(hole.Bottom - TourGeometry.InsideFoot, card.Bottom);
    }

    // However small the window, the card's top left stays in sight - its buttons are what ends the tour.
    [Fact]
    public void In_a_window_smaller_than_the_card_its_top_left_stays_in_sight()
    {
        var tiny = new Size(200, 100);
        (Rect card, _) = TourGeometry.PlaceCard(new Rect(5, 5, 50, 50), Card, tiny);

        Assert.Equal(TourGeometry.EdgeMargin, card.X);
        Assert.Equal(TourGeometry.EdgeMargin, card.Y);
    }

    [Fact]
    public void The_pointer_reaches_from_the_cards_edge_towards_the_hole()
    {
        var hole = new Rect(10, 17, 190, 130);
        (Rect card, TourCardSide side) = TourGeometry.PlaceCard(hole, Card, Window);

        (Point a, Point tip, Point b) = TourGeometry.PointerFor(hole, card, side)!.Value;

        Assert.Equal(card.X - TourGeometry.PointerDepth, tip.X);
        Assert.True(tip.X > hole.Right, "the pointer must stop short of the hole");
        Assert.Equal(a.X, b.X);
        Assert.Equal(2 * TourGeometry.PointerHalf, b.Y - a.Y);
    }

    // Level with the hole's middle when the card allows it, else as near as the card's corners let it be.
    [Fact]
    public void The_pointer_stays_within_the_cards_edge()
    {
        var hole = new Rect(10, 140, 190, 680);
        var card = new Rect(218, 140, 340, 180);

        (_, Point tip, _) = TourGeometry.PointerFor(hole, card, TourCardSide.Right)!.Value;

        Assert.Equal(card.Bottom - TourGeometry.PointerInset, tip.Y);
    }

    [Fact]
    public void A_card_in_the_middle_or_inside_the_hole_has_no_pointer()
    {
        Assert.Null(TourGeometry.PointerFor(new Rect(10, 10, 100, 100), new Rect(20, 20, 50, 50), TourCardSide.Centre));
        Assert.Null(TourGeometry.PointerFor(new Rect(10, 10, 100, 100), new Rect(20, 20, 50, 50), TourCardSide.Inside));
    }
}
