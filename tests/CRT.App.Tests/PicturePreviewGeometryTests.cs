using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// PicturePreviewGeometry - the pane's preview of a picture and its marked components (CRT 4.0 step
// 2, slice 4b): the picture fitted whole and centred, a highlight moved and scaled with it.
public sealed class PicturePreviewGeometryTests
{
    [Fact]
    public void A_wide_picture_fills_the_width_and_is_centred_top_to_bottom()
    {
        Assert.Equal(new Rect(0, 25, 200, 50), PicturePreviewGeometry.Fit(new Size(4000, 1000), new Size(200, 100)));
    }

    [Fact]
    public void A_tall_picture_fills_the_height_and_is_centred_side_to_side()
    {
        Assert.Equal(new Rect(75, 0, 50, 100), PicturePreviewGeometry.Fit(new Size(500, 1000), new Size(200, 100)));
    }

    [Fact]
    public void Nothing_of_no_size_is_drawn()
    {
        Assert.Equal(default, PicturePreviewGeometry.Fit(new Size(0, 100), new Size(200, 100)));
        Assert.Equal(default, PicturePreviewGeometry.Fit(new Size(100, 100), new Size(0, 100)));
        Assert.Equal(default, PicturePreviewGeometry.Place(new Rect(1, 1, 1, 1), new Size(0, 0), new Rect(0, 0, 10, 10)));
    }

    // A highlight in the picture's pixels lands on the same part of the drawn picture.
    [Fact]
    public void A_highlight_is_scaled_and_moved_with_the_picture()
    {
        Rect target = PicturePreviewGeometry.Fit(new Size(4000, 1000), new Size(200, 100));

        Assert.Equal(new Rect(100, 50, 10, 5), PicturePreviewGeometry.Place(new Rect(2000, 500, 200, 100), new Size(4000, 1000), target));
    }
}
