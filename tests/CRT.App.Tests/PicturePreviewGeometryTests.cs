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

    // A trace's ends - points in the picture's pixels - land where the rectangles do.
    [Fact]
    public void A_point_is_scaled_and_moved_with_the_picture()
    {
        Rect target = PicturePreviewGeometry.Fit(new Size(4000, 1000), new Size(200, 100));

        Assert.Equal(new Point(100, 50), PicturePreviewGeometry.Place(new Point(2000, 500), new Size(4000, 1000), target));
        Assert.Equal(default, PicturePreviewGeometry.Place(new Point(1, 1), new Size(0, 0), target));
    }

    // ---- zoomed (2026-10-29) ----------------------------------------------------------------------

    private static readonly Size Image = new(400, 200);
    private static readonly Size Area = new(200, 100);

    // Not zoomed, the preview is as it always was - the picture whole, centred.
    [Fact]
    public void Unzoomed_the_view_is_the_picture_fitted_whole()
    {
        Assert.Equal(PicturePreviewGeometry.Fit(Image, Area), PicturePreviewGeometry.View(Image, Area, 1, PicturePreviewGeometry.Middle));

        // Less than whole is whole, and a centre asked for elsewhere stays in the middle then.
        Assert.Equal(PicturePreviewGeometry.Fit(Image, Area), PicturePreviewGeometry.View(Image, Area, 0.25, new Point(0.1, 0.9)));
    }

    // Zooming keeps the point of the picture under the pointer under it - what makes it feel right.
    [Fact]
    public void Zooming_keeps_the_point_under_the_pointer_where_it_is()
    {
        var pointer = new Point(150, 30);
        Rect before = PicturePreviewGeometry.View(Image, Area, 1, PicturePreviewGeometry.Middle);
        Point under = new((pointer.X - before.X) / before.Width, (pointer.Y - before.Y) / before.Height);

        (double zoom, Point centre) = PicturePreviewGeometry.ZoomAt(Image, Area, 1, PicturePreviewGeometry.Middle, pointer, 2);
        Rect after = PicturePreviewGeometry.View(Image, Area, zoom, centre);

        Assert.Equal(2, zoom);
        Assert.Equal(pointer.X, after.X + (under.X * after.Width), 6);
        Assert.Equal(pointer.Y, after.Y + (under.Y * after.Height), 6);
    }

    // However far it zooms, it stays between whole and MaxZoom.
    [Fact]
    public void Zoom_stays_between_whole_and_the_most()
    {
        Assert.Equal(1, PicturePreviewGeometry.ZoomAt(Image, Area, 1.5, PicturePreviewGeometry.Middle, new Point(10, 10), 0.1).Zoom);
        Assert.Equal(PicturePreviewGeometry.MaxZoom, PicturePreviewGeometry.ZoomAt(Image, Area, 10, PicturePreviewGeometry.Middle, new Point(10, 10), 100).Zoom);
        Assert.Equal(1, PicturePreviewGeometry.ZoomAt(Image, Area, 1, PicturePreviewGeometry.Middle, new Point(10, 10), double.NaN).Zoom);
    }

    // Zoomed in, the picture always covers the area - never dragged off to show nothing.
    [Fact]
    public void Zoomed_in_the_picture_always_covers_the_area()
    {
        Point centre = PicturePreviewGeometry.Pan(Image, Area, 4, PicturePreviewGeometry.Middle, new Vector(100000, -100000));
        Rect view = PicturePreviewGeometry.View(Image, Area, 4, centre);

        Assert.Equal(0, view.X, 6);
        Assert.Equal(Area.Height, view.Bottom, 6);
    }

    // A drag moves the picture with the pointer.
    [Fact]
    public void A_drag_moves_the_picture_with_the_pointer()
    {
        Rect before = PicturePreviewGeometry.View(Image, Area, 4, PicturePreviewGeometry.Middle);
        Point centre = PicturePreviewGeometry.Pan(Image, Area, 4, PicturePreviewGeometry.Middle, new Vector(30, -10));
        Rect after = PicturePreviewGeometry.View(Image, Area, 4, centre);

        Assert.Equal(before.X + 30, after.X, 6);
        Assert.Equal(before.Y - 10, after.Y, 6);
    }

    // A picture that fits the area across stays centred across, however it is dragged.
    [Fact]
    public void Along_a_side_the_picture_is_no_larger_than_the_area_it_stays_centred()
    {
        var tall = new Size(100, 1000);
        Point centre = PicturePreviewGeometry.ClampCentre(tall, Area, 2, new Point(0.9, 0.9));

        Assert.Equal(0.5, centre.X);
        Assert.True(centre.Y < 0.9);
    }

    // How sharp the picture must be decoded: its width on screen, in device pixels - never more
    // than the picture has.
    [Fact]
    public void The_pixels_wanted_are_the_zoomed_width_on_screen_up_to_the_pictures_own()
    {
        Assert.Equal(200, PicturePreviewGeometry.PixelsWanted(Image, Area, 1, 1));
        Assert.Equal(400, PicturePreviewGeometry.PixelsWanted(Image, Area, 4, 1));
        Assert.Equal(300, PicturePreviewGeometry.PixelsWanted(Image, Area, 1, 1.5));
        Assert.Equal(0, PicturePreviewGeometry.PixelsWanted(new Size(0, 0), Area, 2, 1));
    }

    // The sharper image is decoded in steps, the sharpest the most - never wider than the picture.
    [Theory]
    [InlineData(700, 6000, 1280)]
    [InlineData(1280, 6000, 1280)]
    [InlineData(1281, 6000, 2048)]
    [InlineData(2500, 6000, 3072)]
    [InlineData(9000, 6000, 3072)]
    [InlineData(2500, 1800, 1800)]
    [InlineData(2500, 0, 3072)]
    public void A_sharper_image_is_decoded_in_steps_up_to_the_sharpest(int wanted, double imageWidth, int expected)
    {
        Assert.Equal(expected, PicturePreviewGeometry.SharpWidth(wanted, imageWidth));
    }

    // ###########################################################################################
    // The cards' pictures are read for the previews as wide as they are dragged, on the screen's
    // scaling, in steps (code review, 2026-10-30: read at 640 whatever their width). The default
    // previews read as before; the widest a drag allows on a 150 % screen still fits.
    // ###########################################################################################
    [Theory]
    [InlineData(320, 1.0, 640)]
    [InlineData(320, 2.0, 640)]
    [InlineData(640, 1.0, 640)]
    [InlineData(641, 1.0, 1024)]
    [InlineData(900, 1.5, 1536)]
    [InlineData(1000, 1.5, 1536)]
    [InlineData(1000, 3.0, 1536)]
    [InlineData(500, 0, 640)]
    [InlineData(double.NaN, 1.0, 640)]
    public void The_cards_pictures_are_read_for_the_previews_width_in_steps(double previewWidth, double scaling, int expected)
    {
        Assert.Equal(expected, PicturePreviewGeometry.WholeWidth(previewWidth, scaling));
    }
}
