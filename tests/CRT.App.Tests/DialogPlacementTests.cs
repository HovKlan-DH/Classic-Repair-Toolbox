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
}
