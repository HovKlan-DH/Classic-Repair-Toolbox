using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BadgeCloudGeometry - where a BadgeCloud's badges stand, as the WrapPanel they replaced laid them out (2026-11-13):
// "KiCad: Files" draws its hundreds of badges as one control, and must wrap them exactly as before.
// ###########################################################################################
public sealed class BadgeCloudGeometryTests
{
    private static Size[] Badges(params double[] widths) => widths.Select(width => new Size(width, 20)).ToArray();

    [Fact]
    public void Badges_go_left_to_right_and_a_new_line_starts_where_the_next_does_not_fit()
    {
        // 30 + 4 + 40 = 74 fits in 80; another 4 + 30 does not. On the second line 30 + 4 + 46 = 80 fills it exactly.
        (List<Rect> places, Size size) = BadgeCloudGeometry.Place(Badges(30, 40, 30, 46), 80, 4, 4);

        Assert.Equal(new Rect(0, 0, 30, 20), places[0]);
        Assert.Equal(new Rect(34, 0, 40, 20), places[1]);
        Assert.Equal(new Rect(0, 24, 30, 20), places[2]);
        Assert.Equal(new Rect(34, 24, 46, 20), places[3]);

        // Two lines, as wide as the widest of them.
        Assert.Equal(new Size(80, 44), size);
    }

    // Exactly the width still fits; a badge wider than the width stands alone on its line, never on an empty one.
    [Fact]
    public void A_badge_that_fills_the_width_fits_and_one_wider_stands_alone()
    {
        (List<Rect> places, _) = BadgeCloudGeometry.Place(Badges(38, 38, 200, 10), 80, 4, 4);

        Assert.Equal(0, places[0].Y);
        Assert.Equal(0, places[1].Y);
        Assert.Equal(new Rect(0, 24, 200, 20), places[2]);
        Assert.Equal(new Rect(0, 48, 10, 20), places[3]);
    }

    // Each line as tall as its tallest badge, every badge of it that tall - as a WrapPanel stretches them.
    [Fact]
    public void A_line_is_as_tall_as_its_tallest_badge_and_so_is_every_badge_on_it()
    {
        (List<Rect> places, Size size) = BadgeCloudGeometry.Place([new Size(20, 18), new Size(20, 22), new Size(100, 18)], 50, 4, 4);

        Assert.Equal(22, places[0].Height);
        Assert.Equal(22, places[1].Height);
        Assert.Equal(26, places[2].Y);
        Assert.Equal(new Size(100, 44), size);
    }

    [Fact]
    public void No_width_is_one_line_and_no_badges_take_no_room()
    {
        (List<Rect> places, Size size) = BadgeCloudGeometry.Place(Badges(30, 40, 50), double.PositiveInfinity, 4, 4);

        Assert.All(places, place => Assert.Equal(0, place.Y));
        Assert.Equal(new Size(128, 20), size);

        (List<Rect> none, Size nothing) = BadgeCloudGeometry.Place([], 80, 4, 4);

        Assert.Empty(none);
        Assert.Equal(default, nothing);
    }
}
