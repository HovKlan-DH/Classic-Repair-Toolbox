using Avalonia;
using Handlers.Geometry;
using Xunit;

namespace Classic_Repair_Toolbox.Tests;

// ###########################################################################################
// OverlaySnapshotGeometry: where a picture of the KiCad overlay is taken, and where it is drawn back
// while the calibration box is dragged or nudged (KiCadOverlayRenderControl.DrawFromSnapshot).
//
// The rule these defend: a pixel of the picture lands on screen EXACTLY where the vector drawing
// would put the trace it shows - with the box moved, resized or flipped since, at any zoom. If it
// did not, the traces seen while aligning would sit somewhere else than where they are saved, which
// is the one thing an alignment must never do, and it would look entirely plausible on screen.
// ###########################################################################################
public class OverlaySnapshotGeometryTests
{
    // A point built in the overlay, as the screen shows it: overlay transform, then the zoom and pan.
    private static Point OnScreen(Point built, Matrix overlay, Matrix visual) => built.Transform(overlay * visual);

    // The same point through the picture: drawn into it under (overlayThen, visual), drawn back under
    // overlayNow in the control's local space, then the zoom and pan the compositor applies.
    private static Point ThroughThePicture(Point built, Rect area, Matrix overlayThen, Matrix visual, Matrix overlayNow)
    {
        Assert.True(OverlaySnapshotGeometry.TryPictureToLocal(area, overlayThen, visual, overlayNow, out Matrix pictureToLocal));

        Point inPicture = built.Transform(OverlaySnapshotGeometry.BuiltToPicture(area, overlayThen, visual));
        return inPicture.Transform(pictureToLocal * visual);
    }

    public static TheoryData<string> Cases() => ["unzoomed, moved", "zoomed, resized", "zoomed, mirrored", "taken mid-drag, moved on"];

    private static (Matrix OverlayThen, Matrix Visual, Matrix OverlayNow) Case(string name) => name switch
    {
        "unzoomed, moved" => (Matrix.Identity, Matrix.Identity, Matrix.CreateTranslation(12.5, -7.25)),
        "zoomed, resized" => (Matrix.Identity, new Matrix(3.0, 0, 0, 3.0, -410.0, -95.5), Matrix.CreateScale(1.07, 0.94) * Matrix.CreateTranslation(-30.0, 18.0)),
        "zoomed, mirrored" => (Matrix.Identity, new Matrix(2.0, 0, 0, 2.0, -120.0, -60.0), Matrix.CreateScale(-1.0, 1.0) * Matrix.CreateTranslation(1300.0, 0.0)),
        "taken mid-drag, moved on" => (Matrix.CreateScale(1.02, 1.0) * Matrix.CreateTranslation(4.0, 2.0), new Matrix(1.5, 0, 0, 1.5, -20.0, -10.0), Matrix.CreateScale(0.97, 1.01) * Matrix.CreateTranslation(-6.0, 9.0)),
        _ => throw new ArgumentException(name),
    };

    [Theory]
    [MemberData(nameof(Cases))]
    public void A_point_drawn_from_the_picture_lands_where_the_vector_drawing_puts_it(string name)
    {
        (Matrix overlayThen, Matrix visual, Matrix overlayNow) = Case(name);
        var area = new Rect(-140.0, -97.0, 1690.0, 1181.0);

        foreach (Point built in new[] { new Point(0, 0), new Point(1410, 987), new Point(333.3, 612.9), new Point(1200, 15) })
        {
            Point expected = OnScreen(built, overlayNow, visual);
            Point actual = ThroughThePicture(built, area, overlayThen, visual, overlayNow);

            Assert.Equal(expected.X, actual.X, 6);
            Assert.Equal(expected.Y, actual.Y, 6);
        }
    }

    // Nothing moved since the picture was taken: it goes back pixel for pixel, a plain move by its corner.
    [Fact]
    public void A_picture_drawn_back_unmoved_is_only_moved_by_its_corner()
    {
        var visual = new Matrix(2.0, 0, 0, 2.0, -50.0, -30.0);
        var area = new Rect(-10.0, 20.0, 500.0, 300.0);

        Assert.True(OverlaySnapshotGeometry.TryPictureToLocal(area, Matrix.Identity, visual, Matrix.Identity, out Matrix pictureToLocal));

        Matrix pictureToScreen = pictureToLocal * visual;
        Assert.Equal(1.0, pictureToScreen.M11, 9);
        Assert.Equal(1.0, pictureToScreen.M22, 9);
        Assert.Equal(-10.0, pictureToScreen.M31, 9);
        Assert.Equal(20.0, pictureToScreen.M32, 9);
    }

    // ###########################################################################################
    // *** ON A SCALED DISPLAY THE WHOLE PICTURE IS DRAWN BACK, ONE OF ITS PIXELS TO ONE OF THE SCREEN'S
    // (owner report, 2026-11-06, a display at 125%: "it looks like it is moving or zooming the entire
    // trace block slightly ... as soon as I grab a marker"). *** DrawImage's source is counted in the
    // image's pixels; the picture's size in DIPs, given in its place, drew 80% of it over the whole area -
    // a 25% zoom as the picture took over. The tests ran at 100%, where the two are the same number.
    // ###########################################################################################
    [Theory]
    [InlineData(1.0)]
    [InlineData(1.25)]
    [InlineData(1.5)]
    [InlineData(2.0)]
    public void On_a_scaled_display_the_whole_picture_is_drawn_back_pixel_for_pixel(double scaling)
    {
        Rect area = OverlaySnapshotGeometry.AreaToCapture(new Size(1229.6, 768.8), new Rect(-400, -300, 3000, 2000), scaling)!.Value;
        PixelSize picture = OverlaySnapshotGeometry.PixelSizeOf(area, scaling)!.Value;

        Rect source = OverlaySnapshotGeometry.SourceRectOf(picture);

        Assert.Equal(new Rect(0, 0, picture.Width, picture.Height), source);
        Assert.Equal(area.Width * scaling, source.Width, 6);
        Assert.Equal(area.Height * scaling, source.Height, 6);
    }

    [Fact]
    public void A_picture_taken_under_a_matrix_that_cannot_be_undone_is_not_drawn_back()
    {
        var flat = new Matrix(0, 0, 0, 0, 0, 0);

        Assert.False(OverlaySnapshotGeometry.TryPictureToLocal(new Rect(0, 0, 10, 10), flat, Matrix.Identity, Matrix.Identity, out _));
        Assert.False(OverlaySnapshotGeometry.TryPictureToLocal(new Rect(0, 0, 10, 10), Matrix.Identity, flat, Matrix.Identity, out _));
    }

    // The view with a quarter of it around every side, cut to where the overlay draws.
    [Fact]
    public void The_picture_covers_the_view_and_its_margin_where_the_overlay_draws()
    {
        var view = new Size(1000, 600);

        Rect? all = OverlaySnapshotGeometry.AreaToCapture(view, new Rect(-5000, -5000, 20000, 20000), 1.0);
        Assert.Equal(new Rect(-250, -150, 1500, 900), all);

        Rect? part = OverlaySnapshotGeometry.AreaToCapture(view, new Rect(100, 50, 400, 300), 1.0);
        Assert.Equal(new Rect(100, 50, 400, 300), part);
    }

    // Edges on whole device pixels: a picture drawn back unmoved lands pixel for pixel.
    [Fact]
    public void The_pictures_edges_are_whole_device_pixels()
    {
        Rect? area = OverlaySnapshotGeometry.AreaToCapture(new Size(1000, 600), new Rect(10.2, 20.9, 100.3, 50.05), 1.5);

        // 10.2 .. 110.5 across is device pixels 15.3 .. 165.75 - so 15 .. 166; 20.9 .. 70.95 down is
        // 31.35 .. 106.425 - so 31 .. 107.
        Assert.NotNull(area);
        Assert.Equal(15.0 / 1.5, area.Value.Left, 9);
        Assert.Equal(166.0 / 1.5, area.Value.Right, 9);
        Assert.Equal(31.0 / 1.5, area.Value.Top, 9);
        Assert.Equal(107.0 / 1.5, area.Value.Bottom, 9);
        Assert.Equal(new PixelSize(151, 76), OverlaySnapshotGeometry.PixelSizeOf(area.Value, 1.5));
    }

    // Nothing near the view: a picture of nothing, which is still a picture.
    [Fact]
    public void An_overlay_away_from_the_view_is_a_picture_of_nothing()
    {
        Rect? area = OverlaySnapshotGeometry.AreaToCapture(new Size(1000, 600), new Rect(5000, 5000, 10, 10), 1.0);

        Assert.NotNull(area);
        Assert.Equal(0.0, area.Value.Width);
        Assert.Null(OverlaySnapshotGeometry.PixelSizeOf(area.Value, 1.0));
    }

    // ###########################################################################################
    // *** A PICTURE IS DRAWN BACK ONLY WHILE IT STILL HOLDS ALL THE VIEW SHOWS (code review, 2026-11-07). *** It reaches a
    // quarter of the view past every side, and was kept from move to move: a box moved further, a little at a time, drew a
    // band with no traces, and a picture of nothing drew nothing once the overlay came into view.
    // ###########################################################################################
    [Theory]
    [MemberData(nameof(Cases))]
    public void A_picture_covers_the_view_only_until_the_box_moved_past_its_margin(string name)
    {
        (Matrix overlayThen, Matrix visual, _) = Case(name);
        var view = new Size(1000, 600);
        var board = new Rect(-5000, -5000, 20000, 20000);

        Rect area = OverlaySnapshotGeometry.AreaToCapture(view, board.TransformToAABB(overlayThen * visual), 1.0)!.Value;

        Assert.True(OverlaySnapshotGeometry.StillCovers(view, board, area, overlayThen, overlayThen, visual));

        // Moved on screen by a fifth of the view: within the quarter around it.
        Assert.True(OverlaySnapshotGeometry.StillCovers(view, board, area, overlayThen, OnScreenMove(overlayThen, visual, 200, 0), visual));
        Assert.True(OverlaySnapshotGeometry.StillCovers(view, board, area, overlayThen, OnScreenMove(overlayThen, visual, 0, -120), visual));

        // Moved further: a band of the view lies beyond the picture.
        Assert.False(OverlaySnapshotGeometry.StillCovers(view, board, area, overlayThen, OnScreenMove(overlayThen, visual, 300, 0), visual));
        Assert.False(OverlaySnapshotGeometry.StillCovers(view, board, area, overlayThen, OnScreenMove(overlayThen, visual, 0, 200), visual));
    }

    // The overlay transform that moves what `overlay` shows by (dx, dy) on screen.
    private static Matrix OnScreenMove(Matrix overlay, Matrix visual, double dx, double dy)
    {
        Assert.True(visual.TryInvert(out Matrix visualInverse));
        return overlay * visual * Matrix.CreateTranslation(dx, dy) * visualInverse;
    }

    // A picture of nothing holds all while nothing is in view - and nothing once the overlay is.
    [Fact]
    public void A_picture_of_nothing_holds_only_while_the_overlay_is_away_from_the_view()
    {
        var view = new Size(1000, 600);
        var board = new Rect(5000, 5000, 100, 100);

        Rect area = OverlaySnapshotGeometry.AreaToCapture(view, board, 1.0)!.Value;

        Assert.True(OverlaySnapshotGeometry.StillCovers(view, board, area, Matrix.Identity, Matrix.Identity, Matrix.Identity));
        Assert.True(OverlaySnapshotGeometry.StillCovers(view, board, area, Matrix.Identity, Matrix.CreateTranslation(-100, -100), Matrix.Identity));
        Assert.False(OverlaySnapshotGeometry.StillCovers(view, board, area, Matrix.Identity, Matrix.CreateTranslation(-4800, -4800), Matrix.Identity));
    }

    // Too many pixels with the margin: the view alone. Too many for that too, or nothing to show it
    // in: no picture, and the overlay is drawn as vectors.
    [Fact]
    public void A_picture_too_large_drops_its_margin_and_then_is_not_taken()
    {
        var everywhere = new Rect(-100000, -100000, 200000, 200000);

        // 3400 x 3400 device pixels is 11.6 million; with the margin it is 26 million.
        Rect? noMargin = OverlaySnapshotGeometry.AreaToCapture(new Size(1700, 1700), everywhere, 2.0);
        Assert.Equal(new Rect(0, 0, 1700, 1700), noMargin);

        Assert.Null(OverlaySnapshotGeometry.AreaToCapture(new Size(3000, 3000), everywhere, 2.0));
        Assert.Null(OverlaySnapshotGeometry.AreaToCapture(new Size(0, 600), everywhere, 1.0));
        Assert.Null(OverlaySnapshotGeometry.AreaToCapture(new Size(1000, 600), everywhere, 0.0));
    }

    // A sheet's few hundred wires are drawn as vectors; a side of a board's thousands from a picture.
    [Fact]
    public void A_picture_is_worth_taking_only_of_enough_primitives()
    {
        Assert.False(OverlaySnapshotGeometry.IsWorthAPicture(222));
        Assert.False(OverlaySnapshotGeometry.IsWorthAPicture(OverlaySnapshotGeometry.PrimitivesWorthAPicture - 1));
        Assert.True(OverlaySnapshotGeometry.IsWorthAPicture(OverlaySnapshotGeometry.PrimitivesWorthAPicture));
        Assert.True(OverlaySnapshotGeometry.IsWorthAPicture(3752));
        Assert.True(OverlaySnapshotGeometry.IsWorthAPicture(2, worthAPicture: 1));
    }

    [Fact]
    public void Whole_pixels_grow_a_rect_outward_and_leave_an_empty_one_empty()
    {
        Assert.Equal(new Rect(1, 2, 4, 5), OverlaySnapshotGeometry.OnWholePixels(new Rect(1.4, 2.9, 3.2, 3.2), 1.0));
        Assert.Equal(default, OverlaySnapshotGeometry.OnWholePixels(new Rect(1, 1, 0, 5), 1.0));
    }
}
