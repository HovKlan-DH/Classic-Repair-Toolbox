using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// KiCadNetConnections - the "Connected to:" lines of the "Schematics" tab's "KiCad nets" panel (owner requests,
// 2026-11-13: each component named as its component info window's title names it, in "[ ]", the board label and the pin
// bold - "[<b>U17</b> | 74LS373] pin <b>5</b>").
// ###########################################################################################
public sealed class KiCadNetConnectionsTests
{
    private static readonly ComponentEntry[] Components =
    [
        new() { BoardLabel = "U5", FriendlyName = "CIA 1", TechnicalNameOrValue = "6581" },
        new() { BoardLabel = "U6", FriendlyName = "CPU", TechnicalNameOrValue = "8502" },
        new() { BoardLabel = "RP9", TechnicalNameOrValue = "1KΩ 1/8W 10%" },
        new() { BoardLabel = "U17", TechnicalNameOrValue = "74LS373" },
    ];

    private static List<string> Texts(IEnumerable<KiCadNetConnectionLine> lines) => lines.Select(line => line.Text).ToList();

    // The owner's own example: each line the window's title in "[ ]", then the pin - in the board's order, not KiCad's.
    [Fact]
    public void Each_pad_is_named_by_its_components_window_title_in_the_boards_order()
    {
        List<KiCadNetConnectionLine> lines = KiCadNetConnections.Lines(
            [
                new("RP9", "6", "RESET"),
                new("U6", "11", "RESET"),
                new("U5", "13", "RESET"),
                new("U5", "1", "GND"),
            ],
            ["RESET"],
            Components,
            "PAL");

        Assert.Equal(["[U5 | CIA 1 | 6581] pin 13", "[U6 | CPU | 8502] pin 11", "[RP9 | 1KΩ 1/8W 10%] pin 6"], Texts(lines));
    }

    // What is drawn bold is the label and the pin alone, each a piece of its own - "<b>" never in the text, which a typed
    // friendly name could hold.
    [Fact]
    public void The_label_and_the_pin_are_pieces_of_their_own_for_the_bold()
    {
        KiCadNetConnectionLine line = Assert.Single(KiCadNetConnections.Lines([new("U17", "5", "D0")], ["D0"], Components, "PAL"));

        Assert.Equal("U17", line.Label);
        Assert.Equal(" | 74LS373] pin ", line.AfterLabel);
        Assert.Equal("5", line.Pin);
        Assert.Equal(string.Empty, line.AfterPin);
        Assert.Equal("[U17 | 74LS373] pin 5", line.Text);
    }

    // A reference the board has no component of is named by its reference alone, still in "[ ]", and comes after the
    // board's own; a pad with no number reads "?".
    [Fact]
    public void A_reference_the_board_does_not_have_keeps_its_reference_alone()
    {
        List<KiCadNetConnectionLine> lines = KiCadNetConnections.Lines(
            [new("J7", "2", "RESET"), new("U6", "", "RESET")],
            ["reset"],
            Components,
            "PAL");

        Assert.Equal(["[U6 | CPU | 8502] pin ?", "[J7] pin 2"], Texts(lines));
        Assert.Equal("] pin ", lines[1].AfterLabel);
    }

    // A component told apart per region is named by its row in the region shown, as its window names it.
    [Fact]
    public void A_regional_component_is_named_by_its_row_in_the_region_shown()
    {
        ComponentEntry[] components =
        [
            new() { BoardLabel = "U1", FriendlyName = "VIC", TechnicalNameOrValue = "6569", Region = "PAL" },
            new() { BoardLabel = "U1", FriendlyName = "VIC", TechnicalNameOrValue = "6567", Region = "NTSC" },
        ];

        KiCadPadOnNet[] pads = [new("U1", "17", "PHI0")];

        Assert.Equal(["[U1 | VIC | 6569] pin 17"], Texts(KiCadNetConnections.Lines(pads, ["PHI0"], components, "PAL")));
        Assert.Equal(["[U1 | VIC | 6567] pin 17"], Texts(KiCadNetConnections.Lines(pads, ["PHI0"], components, "NTSC")));
    }

    // With several nets highlighted each line says its net; pins in number order, each line once.
    [Fact]
    public void With_several_nets_each_line_says_its_net_and_pins_go_in_number_order()
    {
        List<KiCadNetConnectionLine> lines = KiCadNetConnections.Lines(
            [
                new("U6", "11", "RESET"),
                new("U6", "2", "PHI2"),
                new("U6", "2", "PHI2"),
                new("U6", "9", "OTHER"),
            ],
            ["RESET", "PHI2"],
            Components,
            "PAL");

        Assert.Equal(["[U6 | CPU | 8502] pin 2 [PHI2]", "[U6 | CPU | 8502] pin 11 [RESET]"], Texts(lines));
    }

    // No board read yet: the references, as before.
    [Fact]
    public void Without_a_board_the_references_are_listed()
    {
        Assert.Equal(["[U5] pin 13"], Texts(KiCadNetConnections.Lines([new("U5", "13", "RESET")], ["RESET"], null, null)));
    }
}
