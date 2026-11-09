using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Which region CRT shows, and its colors (RegionChoice; 2026-11-09: "do the region steps 1-9, so we are prepared for other
// regions"). A board in PAL and NTSC shows exactly what it always did.
// ###########################################################################################
public sealed class RegionChoiceTests
{
    [Fact]
    public void A_pal_and_ntsc_board_shows_the_region_picked_anywhere_as_it_always_did()
    {
        Assert.Equal("NTSC", RegionChoice.ToShow(BoardRegions.PalAndNtsc, "NTSC", pickedOnBoard: null));
        Assert.Equal("PAL", RegionChoice.ToShow(BoardRegions.PalAndNtsc, "pal", pickedOnBoard: null));
    }

    // A region of the board's own picked on it comes back on it; a board that has neither falls back to its first.
    [Fact]
    public void A_board_of_its_own_regions_shows_the_one_picked_on_it_else_its_first()
    {
        string[] regions = ["EU", "US"];

        Assert.Equal("US", RegionChoice.ToShow(regions, "PAL", pickedOnBoard: "us"));
        Assert.Equal("EU", RegionChoice.ToShow(regions, "PAL", pickedOnBoard: null));
        Assert.Equal("NTSC", RegionChoice.ToShow(["PAL", "NTSC", "Japan"], "NTSC", pickedOnBoard: "Elsewhere"));
    }

    // A board without regions changes nothing of what was picked.
    [Fact]
    public void A_board_without_regions_keeps_what_was_picked()
    {
        Assert.Equal("NTSC", RegionChoice.ToShow([], "NTSC", pickedOnBoard: null));
    }

    [Fact]
    public void Pal_and_ntsc_keep_their_colors_and_the_boards_own_regions_take_turns_with_four_more()
    {
        string[] regions = ["PAL", "NTSC", "Japan", "EU", "US", "Asia", "UK"];

        Assert.Equal("Schematics_Region_PAL", RegionChoice.ColorsOf(regions, "pal"));
        Assert.Equal("Schematics_Region_NTSC", RegionChoice.ColorsOf(regions, "NTSC"));
        Assert.Equal("Schematics_Region_Own1", RegionChoice.ColorsOf(regions, "Japan"));
        Assert.Equal("Schematics_Region_Own4", RegionChoice.ColorsOf(regions, "Asia"));
        Assert.Equal("Schematics_Region_Own1", RegionChoice.ColorsOf(regions, "UK"));
        Assert.Equal("SchematicsRegion", RegionChoice.ColorsOf(regions, "Mars"));
        Assert.Null(RegionChoice.OwnNumberOf(regions, "PAL"));
    }
}
