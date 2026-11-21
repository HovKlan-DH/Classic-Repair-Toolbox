using Avalonia;
using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// DialogPlacement - where a dialog that sizes itself to its content goes (owner report, 2026-11-01,
// of the Submit dialog: taller than the screen, its Submit button out of reach, and its next panel
// shown at the top of where the tall one had been).
// ###########################################################################################
public sealed class DialogPlacementTests
{
    // The working area less a margin above and below, in device-independent units - at 100 % and at
    // 150 %, where 1560 device pixels are 1040 units.
    [Theory]
    [InlineData(1040, 1.0, 992)]
    [InlineData(1560, 1.5, 992)]
    [InlineData(2120, 2.0, 1012)]
    public void A_dialog_is_no_taller_than_the_working_area_less_its_margins(int workingAreaHeight, double scaling, double expected)
    {
        Assert.Equal(expected, DialogPlacement.MaxHeight(workingAreaHeight, scaling));
    }

    // However small the screen, a dialog keeps room for a few lines and its buttons.
    [Fact]
    public void A_tiny_screen_still_leaves_the_dialog_its_lowest_height()
    {
        Assert.Equal(DialogPlacement.LowestMaxHeight, DialogPlacement.MaxHeight(250, 1.0));
    }

    // Centred over its owner, at the owner's scaling: a 520 x 300 dialog at 150 % is 780 x 450
    // device pixels.
    [Fact]
    public void A_dialog_is_centred_over_its_owner()
    {
        var owner = new PixelRect(100, 50, 1600, 1000);
        var screen = new PixelRect(0, 0, 2560, 1400);

        Assert.Equal(new PixelPoint(510, 325), DialogPlacement.CentredOver(owner, new Size(520, 300), 1.5, screen));
        Assert.Equal(new PixelPoint(640, 400), DialogPlacement.CentredOver(owner, new Size(520, 300), 1.0, screen));
    }

    // Where centring would put it out of the working area, it is moved in - its top edge first, so its
    // title bar is never above the screen.
    [Fact]
    public void A_dialog_is_kept_inside_the_working_area()
    {
        var screen = new PixelRect(0, 0, 1920, 1040);

        // An owner near the right and bottom edges.
        Assert.Equal(new PixelPoint(1400, 740), DialogPlacement.CentredOver(new PixelRect(1700, 900, 400, 200), new Size(520, 300), 1.0, screen));

        // A dialog taller than the screen stands at its top.
        Assert.Equal(new PixelPoint(700, 0), DialogPlacement.CentredOver(new PixelRect(0, 0, 1920, 1040), new Size(520, 1500), 1.0, screen));

        // A screen that tells nothing keeps it simply centred.
        Assert.Equal(new PixelPoint(1640, 850), DialogPlacement.CentredOver(new PixelRect(1700, 900, 400, 200), new Size(520, 300), 1.0, default));
    }

    // ###########################################################################################
    // A PLACE KEPT FROM LAST TIME IS MADE TO FIT THE SCREEN IT OPENS ON (code review, 2026-11-21: "Capture from oscilloscope
    // pin-by-pin"'s window, left 1280 x 1000 on a large monitor, opened 1000 high on a laptop of 768, its buttons off the bottom).
    // ###########################################################################################
    private static readonly Size Least = new(1100, 640);

    // A place that fits stays as it was kept.
    [Fact]
    public void A_kept_place_that_fits_its_screen_is_kept()
    {
        Assert.Equal(
            (new Size(1200, 760), new PixelPoint(40, 50)),
            DialogPlacement.KeptPlaceOn(new Size(1200, 760), Least, new PixelPoint(40, 50), new PixelRect(0, 0, 1920, 1040), 1.0));
    }

    // Taller than the working area less the frame: as tall as that, and moved up until the frame's foot is on the screen.
    [Fact]
    public void A_kept_size_too_large_for_the_screen_is_made_to_fit_and_moved_onto_it()
    {
        // 728 high less the 40 of the frame; 1296 wide with the frame's 16, so no further right than 70.
        Assert.Equal(
            (new Size(1280, 688), new PixelPoint(70, 0)),
            DialogPlacement.KeptPlaceOn(new Size(1280, 1000), Least, new PixelPoint(100, 50), new PixelRect(0, 0, 1366, 728), 1.0));

        // At 150 %: 2400 x 1440 device pixels are 1600 x 960 units, the window and its frame 1944 x 1440 device pixels.
        Assert.Equal(
            (new Size(1280, 920), new PixelPoint(456, 0)),
            DialogPlacement.KeptPlaceOn(new Size(1280, 1000), Least, new PixelPoint(1000, 300), new PixelRect(0, 0, 2400, 1440), 1.5));

        // A second screen to the right of the first: kept on it.
        Assert.Equal(
            (new Size(1280, 688), new PixelPoint(1990, 0)),
            DialogPlacement.KeptPlaceOn(new Size(1280, 1000), Least, new PixelPoint(2000, 100), new PixelRect(1920, 0, 1366, 728), 1.0));
    }

    // However small the screen, never under the window's least size; and a screen that tells nothing leaves the place as kept.
    [Fact]
    public void A_kept_place_is_never_under_the_least_size()
    {
        Assert.Equal(new Size(1100, 640), DialogPlacement.KeptPlaceOn(new Size(1280, 1000), Least, new PixelPoint(0, 0), new PixelRect(0, 0, 1024, 600), 1.0).Size);
        Assert.Equal(
            (new Size(1100, 1000), new PixelPoint(5000, 5000)),
            DialogPlacement.KeptPlaceOn(new Size(900, 1000), Least, new PixelPoint(5000, 5000), default, 1.0));
    }
}
