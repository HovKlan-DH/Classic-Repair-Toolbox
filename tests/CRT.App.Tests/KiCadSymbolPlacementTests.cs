using Handlers.Geometry;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadSymbolPlacement - where a point of a library symbol lands on its sheet. A library symbol is
// drawn Y up and a sheet Y down, the angle turns it first and the mirror flips it on the sheet: the
// order that put a pin's tip on a wire on every KiCad project CRT ships (measured 2026-11-04). Worked
// out by hand for the point (2, 3) of a symbol placed at (100, 50).
// ###########################################################################################
public sealed class KiCadSymbolPlacementTests
{
    [Theory]
    [InlineData(0, null, 102, 47)]
    [InlineData(90, null, 97, 48)]
    [InlineData(180, null, 98, 53)]
    [InlineData(270, null, 103, 52)]
    [InlineData(0, "x", 102, 53)]
    [InlineData(0, "y", 98, 47)]
    [InlineData(180, "x", 98, 47)]
    [InlineData(180, "y", 102, 53)]
    public void A_point_lands_turned_and_flipped(double angle, string? mirror, double expectedX, double expectedY)
    {
        Assert.Equal((expectedX, expectedY), KiCadSymbolPlacement.OnSheet(2, 3, 100, 50, angle, mirror));
    }

    // *** THE ONE THAT WAS MEASURED: *** turned first, flipped after. The other order gives another point
    // for exactly these, and put 71% of the tips at 90 degrees off their wires.
    [Theory]
    [InlineData(90, "x", 97, 52)]
    [InlineData(90, "y", 103, 48)]
    [InlineData(270, "x", 103, 48)]
    [InlineData(270, "y", 97, 52)]
    public void A_symbol_both_turned_and_mirrored_is_turned_first_and_flipped_on_the_sheet(double angle, string mirror, double expectedX, double expectedY)
    {
        Assert.Equal((expectedX, expectedY), KiCadSymbolPlacement.OnSheet(2, 3, 100, 50, angle, mirror));
    }

    [Theory]
    [InlineData(-90, 270)]
    [InlineData(450, 90)]
    [InlineData(360, 0)]
    public void An_angle_past_a_full_turn_is_the_same_angle(double angle, double same)
    {
        Assert.Equal(KiCadSymbolPlacement.OnSheet(2, 3, 100, 50, same, null), KiCadSymbolPlacement.OnSheet(2, 3, 100, 50, angle, null));
    }
}
