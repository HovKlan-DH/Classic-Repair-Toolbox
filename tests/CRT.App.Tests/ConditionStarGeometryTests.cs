using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ConditionStarGeometry - the shape of a star of a device's condition, drawn as a shape so it can take a gradient, an
// edge and a shadow (owner request, 2026-11-06: "more 'depth' and coloring"): ten corners, the first straight up,
// alternating outer and inner, inside its square with room for the edge, standing in its middle.
// ###########################################################################################
public sealed class ConditionStarGeometryTests
{
    [Fact]
    public void A_star_has_ten_corners_from_the_top_point_alternating_out_and_in()
    {
        IReadOnlyList<Point> corners = ConditionStarGeometry.Corners(20, inset: 1);

        Assert.Equal(10, corners.Count);

        // The first point straight up, at the top of the room inside the square.
        Assert.Equal(10, corners[0].X, 6);
        Point middle = new(10, corners[0].Y + 9);

        double Distance(Point point) => Math.Sqrt(Math.Pow(point.X - middle.X, 2) + Math.Pow(point.Y - middle.Y, 2));

        for (int corner = 0; corner < 10; corner++)
        {
            Assert.Equal(corner % 2 == 0 ? 9 : 9 * ConditionStarGeometry.InnerRatio, Distance(corners[corner]), 6);
        }

        // Mirrored left and right: corner n and corner 10 - n.
        for (int corner = 1; corner < 5; corner++)
        {
            Assert.Equal(20 - corners[corner].X, corners[10 - corner].X, 6);
            Assert.Equal(corners[corner].Y, corners[10 - corner].Y, 6);
        }
    }

    [Fact]
    public void A_star_stands_in_the_middle_of_its_square_inside_the_room_for_its_edge()
    {
        IReadOnlyList<Point> corners = ConditionStarGeometry.Corners(24, inset: 0.7);

        double top = corners.Min(point => point.Y);
        double bottom = corners.Max(point => point.Y);

        Assert.True(corners.All(point => point.X >= 0.7 - 1e-9 && point.X <= 24 - 0.7 + 1e-9));
        Assert.Equal(24 - bottom, top, 6);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(-4, 0)]
    [InlineData(10, 5)]
    [InlineData(10, -1)]
    [InlineData(double.NaN, 0)]
    public void No_room_is_no_star(double size, double inset)
    {
        Assert.Empty(ConditionStarGeometry.Corners(size, inset));
    }
}
