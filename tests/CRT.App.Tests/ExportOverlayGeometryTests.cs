using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// Where a worklog's marked area lands on a schematic in the exported PDF.
//
// THE BUG THESE EXIST FOR: the first version computed these fractions correctly and then handed
// them to QuestPDF's PaddingLeft/PaddingTop x100, believing those took a percentage. Every QuestPDF
// padding is an absolute length, so "58% across" became "58 points across" and an area covering a
// tenth of the board was drawn covering most of it. Nothing threw; the PDF simply disagreed with
// the screen, which is what was reported.
//
// The numbers below are taken from a REAL entry on a real board (workbook 1, entry #3 on the C128
// 310378 "I/O, Ports & Connectors" sheet, 3552x2477) precisely so they can be checked against what
// the UI draws from the same record rather than against a fixture invented to match the code.
public class ExportOverlayGeometryTests
{
    private const int BoardWidth = 3552;

    private const int BoardHeight = 2477;

    // The real stored area of entry #3.
    private const double AreaX = 2060.667771333886;

    private const double AreaY = 285.6371168185584;

    private const double AreaWidth = 373.0770505385249;

    private const double AreaHeight = 527.5542667771334;

    [Fact]
    public void A_real_marked_area_maps_to_the_same_fractions_the_ui_draws()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(
            AreaX, AreaY, AreaWidth, AreaHeight, BoardWidth, BoardHeight);

        Assert.NotNull(fractions);

        // ~58% across and ~11.5% down, about a tenth of the width and a fifth of the height. These
        // are the values the on-screen overlay produces from the same record; the version that
        // shipped drew this area spanning most of the board instead.
        Assert.Equal(0.580, fractions!.Value.Left, 3);
        Assert.Equal(0.115, fractions.Value.Top, 3);
        Assert.Equal(0.105, fractions.Value.Width, 3);
        Assert.Equal(0.213, fractions.Value.Height, 3);
    }

    // The area must occupy a SMALL part of the image - the specific thing that went wrong. Stated
    // as its own assertion because the exact fractions above could all be individually plausible
    // while the rectangle still swallowed the page.
    [Fact]
    public void A_real_marked_area_covers_only_a_small_fraction_of_the_board()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(
            AreaX, AreaY, AreaWidth, AreaHeight, BoardWidth, BoardHeight);

        Assert.True(fractions!.Value.Width < 0.2, $"width fraction was {fractions.Value.Width}");
        Assert.True(fractions.Value.Height < 0.3, $"height fraction was {fractions.Value.Height}");
    }

    [Fact]
    public void The_remainders_complete_the_image_exactly()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(
            AreaX, AreaY, AreaWidth, AreaHeight, BoardWidth, BoardHeight);

        // The layout builds its bands from these, so left + width + remaining must be the whole
        // image - a band set that does not add to 1 puts the overlay progressively out of step.
        Assert.Equal(1.0, fractions!.Value.Left + fractions.Value.Width + fractions.Value.RemainingRight, 6);
        Assert.Equal(1.0, fractions.Value.Top + fractions.Value.Height + fractions.Value.RemainingBottom, 6);
    }

    // An area drawn against a LARGER image than the one now on disk (the board scan was replaced)
    // is pulled back inside the picture rather than dropped - a rectangle slightly wrong still
    // points at the right part of the board, while a missing one says nothing at all.
    [Fact]
    public void An_area_reaching_past_the_image_is_clamped_inside_it()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(
            900, 900, 400, 400, 1000, 1000);

        Assert.NotNull(fractions);
        Assert.Equal(0.9, fractions!.Value.Left, 6);
        Assert.Equal(0.9, fractions.Value.Top, 6);

        // Clamped against what is left after the corner, not to the raw 0.4.
        Assert.Equal(0.1, fractions.Value.Width, 6);
        Assert.Equal(0.1, fractions.Value.Height, 6);
        Assert.Equal(0.0, fractions.Value.RemainingRight, 6);
    }

    // ###########################################################################################
    // An area hanging off the LEFT or TOP is CLIPPED, not slid inward.
    //
    // Clamping the origin alone kept the full width and moved the rectangle right, so an area at
    // x=-50 w=100 on a 1000px image drew 100px of rectangle starting at 0 - twice the visible size,
    // covering copper the entry does not mark. The visible part is 50px, i.e. 0.05 of the width.
    [Fact]
    public void An_area_starting_before_the_left_edge_is_clipped_not_shifted()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(-50, 10, 100, 100, 1000, 1000);

        Assert.NotNull(fractions);
        Assert.Equal(0.0, fractions!.Value.Left, 6);

        // Half the area is off the image, so half the width survives - NOT the full 0.1.
        Assert.Equal(0.05, fractions.Value.Width, 6);
    }

    [Fact]
    public void An_area_starting_above_the_top_edge_is_clipped_not_shifted()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(10, -80, 100, 100, 1000, 1000);

        Assert.NotNull(fractions);
        Assert.Equal(0.0, fractions!.Value.Top, 6);
        Assert.Equal(0.02, fractions.Value.Height, 6);
    }

    // An area larger than the image in every direction collapses to exactly the image - the drawn
    // rectangle can never be bigger than the picture it is drawn on.
    [Fact]
    public void An_area_larger_than_the_image_becomes_the_whole_image()
    {
        var fractions = ExportOverlayGeometry.TryBuildAreaFractions(-500, -500, 5000, 5000, 1000, 1000);

        Assert.NotNull(fractions);
        Assert.Equal(0.0, fractions!.Value.Left, 6);
        Assert.Equal(0.0, fractions.Value.Top, 6);
        Assert.Equal(1.0, fractions.Value.Width, 6);
        Assert.Equal(1.0, fractions.Value.Height, 6);
    }

    // Entirely off the LEFT, the mirror of the existing off-the-right case.
    [Fact]
    public void An_area_entirely_left_of_the_image_is_dropped()
    {
        Assert.Null(ExportOverlayGeometry.TryBuildAreaFractions(-200, 10, 100, 100, 1000, 1000));
    }

    [Fact]
    public void An_area_entirely_off_the_image_is_dropped()
    {
        Assert.Null(ExportOverlayGeometry.TryBuildAreaFractions(1200, 1200, 100, 100, 1000, 1000));
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1000, 0)]
    [InlineData(0, 1000)]
    [InlineData(-10, 10)]
    public void A_zero_or_negative_image_size_has_no_fractions(int width, int height)
    {
        Assert.Null(ExportOverlayGeometry.TryBuildAreaFractions(10, 10, 10, 10, width, height));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-5)]
    public void An_empty_area_is_dropped(double size)
    {
        Assert.Null(ExportOverlayGeometry.TryBuildAreaFractions(10, 10, size, size, 1000, 1000));
    }

    // ###########################################################################################
    // The band aspect ratio - the single number the exported layout is actually built from, so
    // that no page dimension ever enters the maths.
    // ###########################################################################################
    [Fact]
    public void A_band_covering_the_whole_image_has_the_images_own_aspect_ratio()
    {
        double? ratio = ExportOverlayGeometry.TryBuildBandAspectRatio(1.0, 1.0, BoardWidth, BoardHeight);

        Assert.Equal(BoardWidth / (double)BoardHeight, ratio!.Value, 6);
    }

    // A band half as wide and half as tall is the same SHAPE as the whole image - the ratio must
    // not change. This is the property that keeps a marked area square-ish rather than stretched.
    [Fact]
    public void A_proportionally_scaled_band_keeps_the_images_aspect_ratio()
    {
        double? whole = ExportOverlayGeometry.TryBuildBandAspectRatio(1.0, 1.0, BoardWidth, BoardHeight);
        double? half = ExportOverlayGeometry.TryBuildBandAspectRatio(0.5, 0.5, BoardWidth, BoardHeight);

        Assert.Equal(whole!.Value, half!.Value, 6);
    }

    [Fact]
    public void A_wide_flat_band_has_a_larger_ratio_than_a_tall_narrow_one()
    {
        double? wide = ExportOverlayGeometry.TryBuildBandAspectRatio(0.8, 0.1, BoardWidth, BoardHeight);
        double? tall = ExportOverlayGeometry.TryBuildBandAspectRatio(0.1, 0.8, BoardWidth, BoardHeight);

        Assert.True(wide!.Value > tall!.Value);
    }

    [Theory]
    [InlineData(0, 0.5)]
    [InlineData(0.5, 0)]
    [InlineData(-0.2, 0.5)]
    public void A_band_with_no_extent_has_no_ratio(double widthFraction, double heightFraction)
    {
        Assert.Null(ExportOverlayGeometry.TryBuildBandAspectRatio(
            widthFraction, heightFraction, BoardWidth, BoardHeight));
    }
}
