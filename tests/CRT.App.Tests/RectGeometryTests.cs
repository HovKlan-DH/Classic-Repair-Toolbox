using Avalonia;
using Avalonia.Media;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Tests for RectGeometry - the rectangle, transform and value-parsing helpers shared by the
// schematic viewer and its overlays. Previously private statics inside TabSchematics and
// SchematicHighlightsOverlay.
//
// The parsing helpers matter beyond geometry: they read contributed board data, so they must be
// locale-proof and must never throw on bad input.
public class RectGeometryTests
{
    // ------------------------------------------------------------------- TryInvert

    [Fact]
    public void TryInvert_inverts_a_translation()
    {
        Assert.True(RectGeometry.TryInvert(Matrix.CreateTranslation(10, 20), out Matrix inv));

        Point back = new Point(15, 25).Transform(inv);

        Assert.Equal(5, back.X, precision: 9);
        Assert.Equal(5, back.Y, precision: 9);
    }

    [Fact]
    public void TryInvert_inverts_a_scale()
    {
        Assert.True(RectGeometry.TryInvert(Matrix.CreateScale(2, 4), out Matrix inv));

        Point back = new Point(10, 20).Transform(inv);

        Assert.Equal(5, back.X, precision: 9);
        Assert.Equal(5, back.Y, precision: 9);
    }

    [Fact]
    public void TryInvert_round_trips_a_combined_zoom_and_pan()
    {
        // This is exactly the schematics view matrix: scale then translate.
        Matrix view = Matrix.CreateScale(2.5, 2.5) * Matrix.CreateTranslation(-100, -40);

        Assert.True(RectGeometry.TryInvert(view, out Matrix inv));

        var original = new Point(37, 91);
        Point roundTripped = original.Transform(view).Transform(inv);

        Assert.Equal(original.X, roundTripped.X, precision: 6);
        Assert.Equal(original.Y, roundTripped.Y, precision: 6);
    }

    [Fact]
    public void TryInvert_reports_failure_for_a_singular_matrix_and_yields_the_identity()
    {
        // A zero scale collapses the plane - inverting it is meaningless, so callers must be
        // able to detect it rather than silently transforming by garbage.
        Assert.False(RectGeometry.TryInvert(Matrix.CreateScale(0, 0), out Matrix inv));
        Assert.Equal(Matrix.Identity, inv);
    }

    // --------------------------------------------------------- CreateNormalizedRect

    [Theory]
    [InlineData(0, 0, 10, 10)]     // drag down-right
    [InlineData(10, 10, 0, 0)]     // drag up-left
    [InlineData(0, 10, 10, 0)]     // drag up-right
    [InlineData(10, 0, 0, 10)]     // drag down-left
    public void A_drag_in_any_direction_yields_the_same_positive_rectangle(
        double x1, double y1, double x2, double y2)
    {
        Rect rect = RectGeometry.CreateNormalizedRect(new Point(x1, y1), new Point(x2, y2));

        Assert.Equal(0, rect.X);
        Assert.Equal(0, rect.Y);
        Assert.Equal(10, rect.Width);
        Assert.Equal(10, rect.Height);
    }

    [Fact]
    public void A_zero_length_drag_yields_an_empty_rectangle()
    {
        Rect rect = RectGeometry.CreateNormalizedRect(new Point(5, 5), new Point(5, 5));

        Assert.Equal(0, rect.Width);
        Assert.Equal(0, rect.Height);
    }

    // ------------------------------------------------------- GetImageContentRect

    [Fact]
    public void A_wide_image_is_width_constrained_and_anchored_at_the_origin()
    {
        // The image control is Left/Top aligned, so there is deliberately no centering offset.
        Rect content = RectGeometry.GetImageContentRect(new Size(200, 200), new PixelSize(400, 100));

        Assert.Equal(0, content.X);
        Assert.Equal(0, content.Y);
        Assert.Equal(200, content.Width);
        Assert.Equal(50, content.Height);
    }

    [Fact]
    public void A_tall_image_is_height_constrained_and_anchored_at_the_origin()
    {
        Rect content = RectGeometry.GetImageContentRect(new Size(200, 200), new PixelSize(100, 400));

        Assert.Equal(0, content.X);
        Assert.Equal(0, content.Y);
        Assert.Equal(50, content.Width);
        Assert.Equal(200, content.Height);
    }

    [Fact]
    public void A_matching_aspect_ratio_fills_the_control()
    {
        Rect content = RectGeometry.GetImageContentRect(new Size(200, 100), new PixelSize(400, 200));

        Assert.Equal(200, content.Width);
        Assert.Equal(100, content.Height);
    }

    [Fact]
    public void A_degenerate_control_size_falls_back_to_the_control_rect()
    {
        Rect content = RectGeometry.GetImageContentRect(new Size(0, 0), new PixelSize(400, 200));

        Assert.Equal(0, content.Width);
        Assert.Equal(0, content.Height);
    }

    // ------------------------------------------------- GetCenteredImageContentRect

    [Fact]
    public void A_wide_image_is_width_constrained_and_centered_vertically()
    {
        // Unlike GetImageContentRect, the thumbnail gallery's Image is Stretch-aligned, so its
        // content is centered in the control box rather than anchored at the origin.
        Rect content = RectGeometry.GetCenteredImageContentRect(new Size(200, 200), new PixelSize(400, 100));

        Assert.Equal(0, content.X);
        Assert.Equal(75, content.Y);
        Assert.Equal(200, content.Width);
        Assert.Equal(50, content.Height);
    }

    [Fact]
    public void A_tall_image_is_height_constrained_and_centered_horizontally()
    {
        Rect content = RectGeometry.GetCenteredImageContentRect(new Size(200, 200), new PixelSize(100, 400));

        Assert.Equal(75, content.X);
        Assert.Equal(0, content.Y);
        Assert.Equal(50, content.Width);
        Assert.Equal(200, content.Height);
    }

    [Fact]
    public void A_matching_aspect_ratio_fills_the_control_with_no_centering_offset()
    {
        Rect content = RectGeometry.GetCenteredImageContentRect(new Size(200, 100), new PixelSize(400, 200));

        Assert.Equal(0, content.X);
        Assert.Equal(0, content.Y);
        Assert.Equal(200, content.Width);
        Assert.Equal(100, content.Height);
    }

    [Fact]
    public void A_degenerate_control_size_falls_back_to_the_control_rect_when_centering()
    {
        Rect content = RectGeometry.GetCenteredImageContentRect(new Size(0, 0), new PixelSize(400, 200));

        Assert.Equal(0, content.Width);
        Assert.Equal(0, content.Height);
    }

    // A zero bitmap dimension is reachable, not theoretical: a thumbnail whose image failed to load
    // keeps OriginalPixelSize at 0x0. Unguarded, the aspect division gives Infinity or NaN, and the
    // NaN case returns a rect with NaN X/Width that poisons any layout it reaches.
    [Theory]
    [InlineData(0, 0)]
    [InlineData(400, 0)]
    [InlineData(0, 200)]
    public void A_degenerate_bitmap_size_never_produces_a_NaN_or_infinite_rect(int pixelWidth, int pixelHeight)
    {
        var controlSize = new Size(200, 100);
        var bitmapPixelSize = new PixelSize(pixelWidth, pixelHeight);

        foreach (Rect content in new[]
        {
            RectGeometry.GetImageContentRect(controlSize, bitmapPixelSize),
            RectGeometry.GetCenteredImageContentRect(controlSize, bitmapPixelSize),
        })
        {
            Assert.False(double.IsNaN(content.X) || double.IsInfinity(content.X));
            Assert.False(double.IsNaN(content.Y) || double.IsInfinity(content.Y));
            Assert.False(double.IsNaN(content.Width) || double.IsInfinity(content.Width));
            Assert.False(double.IsNaN(content.Height) || double.IsInfinity(content.Height));
        }
    }

    // ------------------------------------------------- Local <-> pixel conversions

    [Fact]
    public void A_rectangle_round_trips_between_local_and_pixel_space()
    {
        var contentRect = new Rect(0, 0, 200, 100);
        var pixelSize = new PixelSize(400, 200);
        var localRect = new Rect(20, 10, 40, 20);

        Rect pixels = RectGeometry.LocalToPixelRect(localRect, contentRect, pixelSize);
        Rect back = RectGeometry.PixelToLocalRect(pixels, contentRect, pixelSize);

        Assert.Equal(localRect.X, back.X, precision: 9);
        Assert.Equal(localRect.Y, back.Y, precision: 9);
        Assert.Equal(localRect.Width, back.Width, precision: 9);
        Assert.Equal(localRect.Height, back.Height, precision: 9);
    }

    [Fact]
    public void LocalToPixelRect_scales_by_the_content_to_bitmap_ratio()
    {
        Rect pixels = RectGeometry.LocalToPixelRect(
            new Rect(10, 5, 20, 10), new Rect(0, 0, 200, 100), new PixelSize(400, 200));

        Assert.Equal(20, pixels.X);
        Assert.Equal(10, pixels.Y);
        Assert.Equal(40, pixels.Width);
        Assert.Equal(20, pixels.Height);
    }

    [Fact]
    public void LocalToPixelRect_clamps_to_the_bitmap_bounds()
    {
        // A highlight dragged past the edge of the image must not address pixels off the bitmap.
        Rect pixels = RectGeometry.LocalToPixelRect(
            new Rect(150, 50, 200, 200), new Rect(0, 0, 200, 100), new PixelSize(400, 200));

        Assert.True(pixels.Right <= 400);
        Assert.True(pixels.Bottom <= 200);
    }

    [Fact]
    public void PixelToLocalRect_offsets_by_the_content_origin()
    {
        Rect local = RectGeometry.PixelToLocalRect(
            new Rect(0, 0, 400, 200), new Rect(30, 40, 200, 100), new PixelSize(400, 200));

        Assert.Equal(30, local.X);
        Assert.Equal(40, local.Y);
    }

    // -------------------------------------------------------------- TryParseDouble

    [Theory]
    [InlineData("1.5", 1.5)]
    [InlineData("-2.25", -2.25)]
    [InlineData("0", 0)]
    [InlineData("1e3", 1000)]
    public void TryParseDouble_reads_invariant_numbers(string text, double expected)
    {
        Assert.True(RectGeometry.TryParseDouble(text, out double value));
        Assert.Equal(expected, value, precision: 9);
    }

    [Fact]
    public void TryParseDouble_does_not_accept_a_comma_decimal_separator()
    {
        // Board data is authored on Danish and German machines too. "1,5" must not silently
        // become 15 by being read as a thousands separator.
        Assert.False(RectGeometry.TryParseDouble("1,5", out _));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("abc")]
    public void TryParseDouble_rejects_junk(string text)
    {
        Assert.False(RectGeometry.TryParseDouble(text, out _));
    }

    // --------------------------------------------------------- ParseColorOrDefault

    [Fact]
    public void A_hex_colour_is_parsed()
    {
        Color parsed = RectGeometry.ParseColorOrDefault("#FF8000", Colors.Black);

        Assert.Equal(0xFF, parsed.R);
        Assert.Equal(0x80, parsed.G);
        Assert.Equal(0x00, parsed.B);
    }

    [Fact]
    public void A_hex_colour_with_alpha_is_parsed()
    {
        Assert.Equal(0x80, RectGeometry.ParseColorOrDefault("#80FF0000", Colors.Black).A);
    }

    [Fact]
    public void A_named_colour_is_parsed()
    {
        Assert.Equal(Colors.IndianRed, RectGeometry.ParseColorOrDefault("IndianRed", Colors.Black));
    }

    [Fact]
    public void Surrounding_whitespace_is_tolerated()
    {
        Assert.Equal(Colors.Red, RectGeometry.ParseColorOrDefault("  Red  ", Colors.Black));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("not-a-colour")]
    [InlineData("#GGGGGG")]
    public void A_bad_colour_falls_back_instead_of_throwing(string text)
    {
        // Contributed board data can contain anything; a bad swatch must not crash the overlay.
        Assert.Equal(Colors.Magenta, RectGeometry.ParseColorOrDefault(text, Colors.Magenta));
    }

    // ------------------------------------------------------- ParseOpacityOrDefault

    [Theory]
    [InlineData("0.25", 0.25)]
    [InlineData("1", 1.0)]
    [InlineData("0", 0.0)]
    public void An_opacity_in_the_zero_to_one_range_is_taken_as_is(string text, double expected)
    {
        Assert.Equal(expected, RectGeometry.ParseOpacityOrDefault(text, 0.5), precision: 9);
    }

    [Theory]
    [InlineData("50%", 0.5)]
    [InlineData("100%", 1.0)]
    [InlineData("  25 %  ", 0.25)]
    public void A_percentage_is_converted(string text, double expected)
    {
        Assert.Equal(expected, RectGeometry.ParseOpacityOrDefault(text, 0.5), precision: 9);
    }

    [Fact]
    public void A_bare_number_above_one_is_treated_as_a_percentage()
    {
        // "20" in a board file means 20%, not 2000%.
        Assert.Equal(0.20, RectGeometry.ParseOpacityOrDefault("20", 0.5), precision: 9);
    }

    [Fact]
    public void An_out_of_range_value_is_clamped()
    {
        Assert.Equal(1.0, RectGeometry.ParseOpacityOrDefault("500%", 0.5), precision: 9);
        Assert.Equal(0.0, RectGeometry.ParseOpacityOrDefault("-1", 0.5), precision: 9);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("opaque")]
    public void An_unparseable_opacity_falls_back(string text)
    {
        Assert.Equal(0.33, RectGeometry.ParseOpacityOrDefault(text, 0.33), precision: 9);
    }

    // ------------------------------------------------------- PixelToLocalPoint

    // The point conversion must agree with the rect conversion, otherwise a snap guide would
    // drift away from the rectangle edge it was computed from.
    [Fact]
    public void A_pixel_point_maps_to_the_same_place_as_the_rect_conversion()
    {
        var contentRect = new Rect(0, 0, 400, 300);
        var pixelSize = new PixelSize(800, 600);

        var point = RectGeometry.PixelToLocalPoint(new Point(400, 300), contentRect, pixelSize);
        var rect = RectGeometry.PixelToLocalRect(new Rect(400, 300, 0, 0), contentRect, pixelSize);

        Assert.Equal(rect.X, point.X, precision: 9);
        Assert.Equal(rect.Y, point.Y, precision: 9);
    }

    [Fact]
    public void A_pixel_point_is_offset_by_the_content_rect_origin()
    {
        var point = RectGeometry.PixelToLocalPoint(
            new Point(0, 0), new Rect(10, 20, 400, 300), new PixelSize(800, 600));

        Assert.Equal(10.0, point.X, precision: 9);
        Assert.Equal(20.0, point.Y, precision: 9);
    }

    // ------------------------------------------------------- InsetRectForStroke

    // A stroke is drawn centred on the path, so the rect is pulled in by half the thickness on
    // every side to keep the drawn border inside the original bounds.
    [Fact]
    public void A_rect_is_inset_by_half_the_stroke_on_every_side()
    {
        var inset = RectGeometry.InsetRectForStroke(new Rect(10, 20, 100, 50), 4.0);

        Assert.Equal(new Rect(12, 22, 96, 46), inset);
    }

    [Fact]
    public void A_zero_thickness_stroke_leaves_the_rect_alone()
    {
        var rect = new Rect(10, 20, 100, 50);

        Assert.Equal(rect, RectGeometry.InsetRectForStroke(rect, 0.0));
    }

    // A rectangle thinner than its own stroke would otherwise invert into a negative size, which
    // Avalonia rejects - so width and height are floored at zero instead.
    [Fact]
    public void A_rect_thinner_than_its_stroke_collapses_to_zero_rather_than_inverting()
    {
        var inset = RectGeometry.InsetRectForStroke(new Rect(0, 0, 2, 1), 4.0);

        Assert.Equal(0.0, inset.Width, precision: 9);
        Assert.Equal(0.0, inset.Height, precision: 9);
    }

    // ------------------------------------------------- FindKeysWithRectsIntersecting

    [Fact]
    public void A_key_whose_rect_overlaps_the_target_is_returned()
    {
        var rectsByKey = new Dictionary<string, List<Rect>>
        {
            ["C1"] = new List<Rect> { new Rect(0, 0, 10, 10) }
        };

        var matches = RectGeometry.FindKeysWithRectsIntersecting(rectsByKey, new Rect(5, 5, 10, 10));

        Assert.Contains("C1", matches);
    }

    [Fact]
    public void A_key_whose_rect_does_not_overlap_the_target_is_excluded()
    {
        var rectsByKey = new Dictionary<string, List<Rect>>
        {
            ["C1"] = new List<Rect> { new Rect(0, 0, 10, 10) }
        };

        var matches = RectGeometry.FindKeysWithRectsIntersecting(rectsByKey, new Rect(100, 100, 10, 10));

        Assert.Empty(matches);
    }

    [Fact]
    public void A_key_with_multiple_rects_matches_if_any_one_of_them_touches()
    {
        // A component can have more than one highlight rectangle - only one needs to touch.
        var rectsByKey = new Dictionary<string, List<Rect>>
        {
            ["U1"] = new List<Rect> { new Rect(0, 0, 5, 5), new Rect(200, 200, 5, 5) }
        };

        var matches = RectGeometry.FindKeysWithRectsIntersecting(rectsByKey, new Rect(1, 1, 2, 2));

        Assert.Contains("U1", matches);
    }

    [Fact]
    public void Multiple_touched_keys_are_all_returned()
    {
        var rectsByKey = new Dictionary<string, List<Rect>>
        {
            ["C1"] = new List<Rect> { new Rect(0, 0, 10, 10) },
            ["C2"] = new List<Rect> { new Rect(5, 5, 10, 10) },
            ["C3"] = new List<Rect> { new Rect(500, 500, 10, 10) }
        };

        var matches = RectGeometry.FindKeysWithRectsIntersecting(rectsByKey, new Rect(0, 0, 20, 20));

        Assert.Equal(new HashSet<string> { "C1", "C2" }, matches);
    }

    [Fact]
    public void An_empty_rect_dictionary_yields_no_matches()
    {
        var matches = RectGeometry.FindKeysWithRectsIntersecting(new Dictionary<string, List<Rect>>(), new Rect(0, 0, 10, 10));

        Assert.Empty(matches);
    }

    // ------------------------------------------------------- GetCenterScaledControlRect

    // At scale 1 the control occupies exactly its layout box - the case where every wrong
    // formula still looks right, which is why the scaled cases below matter.
    [Fact]
    public void An_unscaled_control_occupies_its_layout_rect()
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(120, 80), new Size(40, 20), 1.0);

        Assert.Equal(new Rect(120, 80, 40, 20), rect);
    }

    // The whole point of the helper: a centered scale moves the TOP-LEFT as well as the size.
    // Halving a 40x20 box at (120,80) leaves 20x10 centered on (140,90) - so (130,85), not the
    // (120,80) that "position plus scaled size" would give. That difference is what put the
    // worklog pill hover rect off the pill at every zoom other than 100%.
    [Fact]
    public void A_shrunk_control_stays_centered_on_its_layout_box()
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(120, 80), new Size(40, 20), 0.5);

        Assert.Equal(new Rect(130, 85, 20, 10), rect);
    }

    [Fact]
    public void A_grown_control_expands_around_its_center()
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(100, 100), new Size(10, 10), 2.0);

        Assert.Equal(new Rect(95, 95, 20, 20), rect);
    }

    // The center is the one point a centered scale cannot move, at any scale.
    [Theory]
    [InlineData(0.25)]
    [InlineData(1.0)]
    [InlineData(3.5)]
    public void The_center_is_unchanged_by_any_scale(double scale)
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(120, 80), new Size(40, 20), scale);

        Assert.Equal(140, rect.Center.X, 6);
        Assert.Equal(90, rect.Center.Y, 6);
    }

    // A pill whose bitmap never loaded measures zero; it must collapse to a point rather than
    // produce a NaN or negative rect that Contains() would answer unpredictably.
    [Theory]
    [InlineData(0, 20)]
    [InlineData(40, 0)]
    [InlineData(0, 0)]
    public void A_degenerate_layout_size_collapses_to_a_point(double width, double height)
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(120, 80), new Size(width, height), 0.5);

        Assert.Equal(new Rect(120, 80, 0, 0), rect);
    }

    // A negative scale would mirror the control and yield a rect with negative extents; clamping
    // keeps Contains() meaningful for a caller that computed its scale from a bad view matrix.
    [Fact]
    public void A_negative_scale_is_clamped_rather_than_mirroring_the_rect()
    {
        var rect = RectGeometry.GetCenterScaledControlRect(new Point(120, 80), new Size(40, 20), -2.0);

        Assert.Equal(new Rect(140, 90, 0, 0), rect);
    }
}
