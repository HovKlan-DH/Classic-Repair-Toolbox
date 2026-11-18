using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ComponentPinDirections - a pin's direction as the "Schematics" tab says it on the schematic image (owner requests,
// 2026-11-17: the "Direction" badge of the pin under the pointer, and the numbered pins' colors): what the component's pin
// list says in the region shown, else the datasheet's of a chip CRT knows - the pinout's order.
// ###########################################################################################
public sealed class ComponentPinDirectionsTests
{
    private static BoardData Board(params ComponentPinEntry[] pins) => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U27", Category = "IC", TechnicalNameOrValue = "74LS08", Package = "DIP-14" },
            new ComponentEntry { BoardLabel = "U7", Category = "IC", TechnicalNameOrValue = "6510", Package = "DIP-40" },
        ],
        ComponentPins = [.. pins],
    };

    [Fact]
    public void A_pin_takes_the_direction_its_list_gives_it()
    {
        BoardData board = Board(
            new ComponentPinEntry { BoardLabel = "U7", Pin = "1", Name = "PHI0", Direction = "Input" },
            new ComponentPinEntry { BoardLabel = "U7", Pin = "2 [64K]", Name = "RDY", Direction = "Unknown" });

        ComponentPinDirections u7 = ComponentPinDirections.For(board, "u7", region: "PAL");

        Assert.Equal("Input", u7.WordOf("1"));
        Assert.Equal(PinDirectionKind.Input, u7.KindOf("1"));

        // A pin written with more after its number is that pin; "Unknown" is what the list says.
        Assert.Equal("Unknown", u7.WordOf("2"));
        Assert.Equal(PinDirectionKind.Other, u7.KindOf("2"));

        // A pin the list leaves out of a chip CRT has no datasheet of: nothing to say.
        Assert.Equal(string.Empty, u7.WordOf("3"));
        Assert.Null(u7.Chip);
    }

    // A chip CRT knows: a pin its list gives no direction takes the datasheet's - the list's own word first.
    [Fact]
    public void A_known_chips_pin_the_list_says_nothing_of_takes_the_datasheets_direction()
    {
        BoardData board = Board(new ComponentPinEntry { BoardLabel = "U27", Pin = "1", Name = "1A", Direction = "Input/output" });

        ComponentPinDirections u27 = ComponentPinDirections.For(board, "U27", region: "PAL");

        Assert.NotNull(u27.Chip);
        Assert.Equal("Input/output", u27.WordOf("1"));
        Assert.Equal("Output", u27.WordOf("3"));
        Assert.Equal("Ground", u27.WordOf("7"));
        Assert.Equal("Power", u27.WordOf("14"));
        Assert.Equal(PinDirectionKind.Output, u27.KindOf("3"));

        // Past the chip's pins, or no pad at all: nothing.
        Assert.Equal(string.Empty, u27.WordOf("15"));

        // A pad numbered "0" is no pin of the datasheet, whose pin 0 is the wires inside a unit (code review, 2026-11-18: it
        // read "Not connected", a 556's a terminal inside the chip).
        Assert.Equal(string.Empty, u27.WordOf("0"));
        Assert.Equal(string.Empty, u27.WordOf(" "));
    }

    // A pad named, not numbered ("A1"), is found by its name; another region's pins are not this region's.
    [Fact]
    public void A_named_pad_is_found_by_its_name_and_only_in_the_region_shown()
    {
        BoardData board = Board(
            new ComponentPinEntry { BoardLabel = "U7", Pin = "a1", Direction = "Output" },
            new ComponentPinEntry { BoardLabel = "U7", Region = "NTSC", Pin = "5", Direction = "Input" });

        Assert.Equal("Output", ComponentPinDirections.For(board, "U7", "PAL").WordOf("A1"));
        Assert.Equal(string.Empty, ComponentPinDirections.For(board, "U7", "PAL").WordOf("5"));
        Assert.Equal("Input", ComponentPinDirections.For(board, "U7", "NTSC").WordOf("5"));
    }

    [Fact]
    public void No_board_or_no_component_says_nothing()
    {
        Assert.Equal(string.Empty, ComponentPinDirections.For(null, "U7", "PAL").WordOf("1"));
        Assert.Equal(string.Empty, ComponentPinDirections.For(Board(), " ", "PAL").WordOf("1"));
        Assert.Equal(string.Empty, ComponentPinDirections.For(Board(), "U99", "PAL").WordOf("1"));
    }
}
