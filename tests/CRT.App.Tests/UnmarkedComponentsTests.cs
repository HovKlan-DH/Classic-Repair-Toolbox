using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// UnmarkedComponents - the components the label editor lists beside a picture as not marked on it
// yet (CRT 4.0 step 3, slice 3c): once per label whatever its regions, the board's spelling and its
// first row's category, in the board's order; a label with a rectangle, in any capitals, is none.
// ###########################################################################################
public sealed class UnmarkedComponentsTests
{
    [Fact]
    public void The_components_with_no_rectangle_are_listed_once_in_the_boards_order()
    {
        ComponentEntry[] components =
        [
            new() { BoardLabel = " U8 ", Category = " IC ", Region = "PAL" },
            new() { BoardLabel = "u8", Category = "Other", Region = "NTSC" },
            new() { BoardLabel = "U1", Category = "IC" },
            new() { BoardLabel = "C12", Category = "Capacitor" },
            new() { BoardLabel = "", Category = "IC" },
        ];

        Assert.Equal(
            [new UnmarkedComponent("U8", "IC"), new UnmarkedComponent("C12", "Capacitor")],
            UnmarkedComponents.Of(components, [" u1 ", null]));

        Assert.Empty(UnmarkedComponents.Of(components, ["U8", "U1", "C12"]));
        Assert.Empty(UnmarkedComponents.Of([], []));
    }
}
