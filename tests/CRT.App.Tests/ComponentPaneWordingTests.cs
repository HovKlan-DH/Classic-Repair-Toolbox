using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPaneWording - what the component pane says (CRT 4.0 step 2, 2026-10-25): its title,
// the region line, the offer to add the case's pins, and what the component editor window still
// edits for the component.
// ###########################################################################################
public sealed class ComponentPaneWordingTests
{
    [Fact]
    public void The_title_is_the_label_and_what_the_component_list_shows_beside_it()
    {
        Assert.Equal("U8 - CPU | 6510", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "U8", FriendlyName = "CPU", TechnicalNameOrValue = "6510" }));
        Assert.Equal("C12 - 10uF", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "C12", TechnicalNameOrValue = " 10uF " }));
        Assert.Equal("R1", ComponentPaneWording.Title(new ComponentEntry { BoardLabel = "R1" }));
    }

    // Said only for a component with a row of its own per region.
    [Fact]
    public void The_region_line_names_the_region_of_a_regional_row_only()
    {
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(""));
        Assert.Equal(string.Empty, ComponentPaneWording.RegionLine(null));
        Assert.StartsWith("The PAL version", ComponentPaneWording.RegionLine(" PAL "), StringComparison.Ordinal);
    }

    // ###########################################################################################
    // What the component editor still edits: the pictures and scope baselines of the region shown,
    // and the component's files - counted, in plain words.
    // ###########################################################################################
    [Fact]
    public void What_is_still_edited_elsewhere_is_counted_for_the_region_shown()
    {
        var board = new BoardData
        {
            ComponentImages =
            [
                new ComponentImageEntry { BoardLabel = "U8", Name = "Top", File = "Images/u8.png" },
                new ComponentImageEntry { BoardLabel = "U8", Name = "Pin 2", Pin = "2", File = "Images/u8-2.png", TimeDiv = "1us" },
                new ComponentImageEntry { BoardLabel = "U8", Region = "NTSC", Name = "Pin 3", Pin = "3", File = "Images/u8-3.png", TimeDiv = "1us" },
                new ComponentImageEntry { BoardLabel = "U9", Name = "Other", File = "Images/u9.png" },
            ],
            ComponentLocalFiles =
            [
                new ComponentLocalFileEntry { BoardLabel = "U8", Name = "Datasheet", File = "Files/6510.pdf" },
                new ComponentLocalFileEntry { BoardLabel = "u8", Name = "Notes", File = "Files/notes.txt" },
            ],
        };

        Assert.Equal(
            "This component has 1 picture, 1 scope baseline and 2 files. They are edited in the component editor for now.",
            ComponentPaneWording.Elsewhere(board, "U8", "PAL"));

        Assert.StartsWith("This component has no pictures", ComponentPaneWording.Elsewhere(board, "R1", "PAL"), StringComparison.Ordinal);
        Assert.Equal(
            "This component has 1 picture. They are edited in the component editor for now.",
            ComponentPaneWording.Elsewhere(board, "U9", "PAL"));
    }

    // ###########################################################################################
    // The offer to add the case's pins: every pin while none is listed, then the ones still missing
    // - and nothing for a case CRT cannot count.
    // ###########################################################################################
    [Fact]
    public void The_case_pins_offer_counts_what_is_missing()
    {
        Assert.Equal(8, ComponentPaneWording.MissingCasePins("DIP-8", []));
        Assert.Equal(6, ComponentPaneWording.MissingCasePins("DIP-8", [new ComponentPinEntry { Pin = "1" }, new ComponentPinEntry { Pin = " 8 " }]));
        Assert.Equal(0, ComponentPaneWording.MissingCasePins("SIP", []));
        Assert.Equal(0, ComponentPaneWording.MissingCasePins("TO-92", [new ComponentPinEntry { Pin = "1" }, new ComponentPinEntry { Pin = "2" }, new ComponentPinEntry { Pin = "3" }]));

        Assert.Equal("Add the 40 pins of DIP-40", ComponentPaneWording.AddCasePins("DIP-40", 40, noPinsYet: true));
        Assert.Equal("Add the 3 missing pins", ComponentPaneWording.AddCasePins("DIP-40", 3, noPinsYet: false));
        Assert.Equal("Add the missing pin", ComponentPaneWording.AddCasePins("DIP-40", 1, noPinsYet: false));
    }

    [Fact]
    public void A_component_no_longer_on_the_board_says_how_to_bring_it_back()
    {
        Assert.Contains("[U8]", ComponentPaneWording.NotOnBoard("U8"), StringComparison.Ordinal);
        Assert.Contains("Undo", ComponentPaneWording.NotOnBoard("U8"), StringComparison.Ordinal);
    }
}
