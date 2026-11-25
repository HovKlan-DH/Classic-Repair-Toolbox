using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// TopLineGeometry - the "Contribute" tab's buttons beside its top line, or under it on a narrow tab (2026-11-25: on
// CRT's narrowest window they left the words some 120 wide, a dozen lines that pushed the workspace below the window).
// ###########################################################################################
public sealed class TopLineGeometryTests
{
    // A wide tab: the words keep far more than a line's room beside the buttons.
    [Fact]
    public void On_a_wide_row_the_buttons_stay_beside_the_words()
    {
        Assert.False(TopLineGeometry.ButtonsGoUnder(rowWidth: 1100, buttonsWidth: 430, gap: 16));
    }

    // CRT's narrowest window: some 560 wide, the buttons 430 - the words would have 114.
    [Fact]
    public void On_the_narrowest_window_the_buttons_go_under_the_words()
    {
        Assert.True(TopLineGeometry.ButtonsGoUnder(rowWidth: 560, buttonsWidth: 430, gap: 16));
    }

    // The words keep exactly their least: still beside; a pixel less: under.
    [Fact]
    public void The_least_text_width_itself_keeps_the_buttons_beside()
    {
        double row = TopLineGeometry.LeastTextWidth + 430 + 16;

        Assert.False(TopLineGeometry.ButtonsGoUnder(row, 430, 16));
        Assert.True(TopLineGeometry.ButtonsGoUnder(row - 1, 430, 16));
    }

    // A row not laid out yet has no width: the buttons stay where the markup puts them, beside.
    [Theory]
    [InlineData(0)]
    [InlineData(double.NaN)]
    public void A_row_not_laid_out_yet_keeps_the_buttons_beside(double rowWidth)
    {
        Assert.False(TopLineGeometry.ButtonsGoUnder(rowWidth, 430, 16));
    }
}
