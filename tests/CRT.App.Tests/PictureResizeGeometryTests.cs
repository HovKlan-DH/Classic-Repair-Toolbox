using Handlers.DataHandling;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// PictureResizeGeometry - a picture in the Feedback tab's text resized by its corners (owner
// request, 2026-10-15): the width becomes the picture's width in the mail, the height follows the
// picture's proportions, and it is never made wider than itself or the mail's column.
public sealed class PictureResizeGeometryTests
{
    [Theory]
    [InlineData(3840, 32, 980)]
    [InlineData(640, 32, 640)]
    [InlineData(20, 20, 20)]   // a picture narrower than the minimum is never made wider than itself
    public void A_picture_is_dragged_between_the_minimum_and_its_own_width_or_the_mails_column(int natural, int min, int max)
    {
        Assert.Equal((min, max), PictureResizeGeometry.WidthRange(natural));
    }

    [Theory]
    [InlineData(980, 3840, 2160, 551)]
    [InlineData(640, 640, 480, 480)]
    [InlineData(1, 1000, 10, 1)]      // never below 1
    [InlineData(300, 0, 0, 300)]      // no proportions known: square
    public void The_height_follows_the_pictures_proportions(int width, int naturalWidth, int naturalHeight, int expected)
    {
        Assert.Equal(expected, PictureResizeGeometry.HeightFor(width, naturalWidth, naturalHeight));
    }

    // Each corner grows the picture as it is dragged away from it, and shrinks it as it is dragged in.
    [Theory]
    [InlineData(PictureCorner.BottomRight, -100, 0, 500)]
    [InlineData(PictureCorner.BottomRight, 50, 0, 640)]   // stops at its own width
    [InlineData(PictureCorner.TopRight, -100, 0, 500)]
    [InlineData(PictureCorner.BottomLeft, 100, 0, 500)]
    [InlineData(PictureCorner.TopLeft, 100, 0, 500)]
    [InlineData(PictureCorner.BottomRight, 0, -75, 500)]  // 75 down of a 4:3 picture is 100 across
    [InlineData(PictureCorner.TopRight, 0, 75, 500)]
    [InlineData(PictureCorner.BottomLeft, 0, -75, 500)]
    [InlineData(PictureCorner.TopLeft, 0, 75, 500)]
    public void A_corner_dragged_inwards_shrinks_the_picture(PictureCorner corner, double dx, double dy, int expected)
    {
        Assert.Equal(expected, PictureResizeGeometry.DraggedWidth(600, 640, 480, corner, dx, dy));
    }

    // A diagonal drag follows whichever direction moved the picture more.
    [Fact]
    public void Of_the_two_directions_the_larger_movement_wins()
    {
        Assert.Equal(400, PictureResizeGeometry.DraggedWidth(600, 640, 480, PictureCorner.BottomRight, -200, -30));
        Assert.Equal(520, PictureResizeGeometry.DraggedWidth(600, 640, 480, PictureCorner.BottomRight, -10, -60));
    }

    [Fact]
    public void It_never_goes_below_the_minimum_or_past_the_mails_column()
    {
        Assert.Equal(FeedbackComposition.MinPictureWidth, PictureResizeGeometry.DraggedWidth(600, 3840, 2160, PictureCorner.BottomRight, -5000, 0));
        Assert.Equal(FeedbackComposition.MaxPictureWidth, PictureResizeGeometry.DraggedWidth(600, 3840, 2160, PictureCorner.BottomRight, 5000, 0));
    }
}
