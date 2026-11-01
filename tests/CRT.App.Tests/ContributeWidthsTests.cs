using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// ContributeWidths - the widths dragged on the Contribute tab, kept per board (owner request,
// 2026-10-29: "Should of course be saved - probably per board, as I assume some boards could look
// differently"). A board never dragged on takes the width dragged last, on any board.
// ###########################################################################################
public sealed class ContributeWidthsTests
{
    private static readonly (double Min, double Max) Range = (100, 500);

    [Fact]
    public void A_board_dragged_on_keeps_its_own_width()
    {
        var byBoard = new Dictionary<string, double>();

        double last = ContributeWidths.Remember(byBoard, "Commodore 64|250407", 320, Range);

        Assert.Equal(320, last);
        Assert.Equal(320, ContributeWidths.For(byBoard, last, "Commodore 64|250407", Range));
    }

    // One drag carries over to every board not dragged on - and a board that has its own keeps it.
    [Fact]
    public void A_board_never_dragged_on_takes_the_width_dragged_last()
    {
        var byBoard = new Dictionary<string, double>();

        double last = ContributeWidths.Remember(byBoard, "Commodore 64|250407", 320, Range);
        last = ContributeWidths.Remember(byBoard, "Commodore 128|310378", 410, Range);

        Assert.Equal(410, ContributeWidths.For(byBoard, last, "Commodore VIC-20|250403", Range));
        Assert.Equal(320, ContributeWidths.For(byBoard, last, "Commodore 64|250407", Range));
    }

    // The same board in other capitals is the same board - one entry, never two.
    [Fact]
    public void A_board_in_other_capitals_is_the_same_board()
    {
        var byBoard = new Dictionary<string, double> { ["Commodore 64|250407"] = 300 };

        ContributeWidths.Remember(byBoard, "commodore 64|250407", 350, Range);

        Assert.Equal(350, Assert.Single(byBoard).Value);
        Assert.Equal(350, ContributeWidths.For(byBoard, 200, "COMMODORE 64|250407", Range));
    }

    // A width out of range - dragged, or written by hand - is brought inside it, so a column is
    // never read back to nothing.
    [Theory]
    [InlineData(20, 100)]
    [InlineData(9000, 500)]
    [InlineData(double.NaN, 100)]
    public void A_width_is_always_in_range(double width, double expected)
    {
        var byBoard = new Dictionary<string, double> { ["A|1"] = width };

        Assert.Equal(expected, ContributeWidths.For(byBoard, 250, "A|1", Range));
        Assert.Equal(expected, ContributeWidths.Remember(new Dictionary<string, double>(), "A|1", width, Range));
    }

    // With no board shown, the width dragged last - and a drag then is kept as that alone.
    [Fact]
    public void No_board_is_the_width_dragged_last()
    {
        var byBoard = new Dictionary<string, double>();

        Assert.Equal(250, ContributeWidths.For(byBoard, 250, string.Empty, Range));
        Assert.Equal(260, ContributeWidths.Remember(byBoard, " ", 260, Range));
        Assert.Empty(byBoard);
    }

    // ###########################################################################################
    // THE SETTINGS BESIDE THE PREVIEWS KEEP THEIR ROOM (owner request, 2026-10-30: dragged to 0px). The
    // previews are shown at most as wide as leaves a card's settings SettingsMinWidth - never under the
    // range's least, and as chosen while the cards are not laid out yet.
    // ###########################################################################################
    [Theory]
    [InlineData(320, 980, 320)]
    [InlineData(900, 980, 980 - ContributeWidths.CardChrome - ContributeWidths.SettingsMinWidth)]
    [InlineData(900, 500, 200)]
    [InlineData(900, 0, 900)]
    [InlineData(5000, 0, 1000)]
    [InlineData(double.NaN, 980, 200)]
    public void The_previews_are_shown_no_wider_than_leaves_the_settings_their_room(double wanted, double cardsWidth, double shown)
    {
        Assert.Equal(shown, ContributeWidths.PreviewShown(wanted, cardsWidth), 3);
    }
}
