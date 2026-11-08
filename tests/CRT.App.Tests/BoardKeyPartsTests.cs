using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// BoardKeyParts - a board key, "{hardware}|{board}", in its two parts, by ONE rule (code review, 2026-11-07: MyDevicesScope,
// the device form and Main each split it their own way, so they could disagree about which hardware a device is of).
// ###########################################################################################
[Collection("DataManager")]
public sealed class BoardKeyPartsTests
{
    [Theory]
    [InlineData("Commodore 64|250407", "Commodore 64", "250407")]
    [InlineData(" Commodore 64 | 250407 ", "Commodore 64", "250407")]
    [InlineData("Commodore 64", "Commodore 64", "")]
    [InlineData("Commodore 64|", "Commodore 64", "")]
    [InlineData("|250407", "", "250407")]
    [InlineData("Amiga|500|rev 6", "Amiga", "500|rev 6")]
    [InlineData("", "", "")]
    [InlineData(null, "", "")]
    public void The_hardware_is_before_the_first_bar_and_the_board_after_it(string? boardKey, string hardware, string board)
    {
        Assert.Equal((hardware, board), BoardKeyParts.Split(boardKey));
    }

    // The three that split a key go by it: the "My Devices" scope, and Main's names for a key no drop-down lists.
    [Fact]
    public void The_scope_and_Mains_names_go_by_the_one_rule()
    {
        Assert.True(MyDevicesScope.Holds(MyDevicesScope.CurrentHardware, " Commodore 64 | 250407 ", "Commodore 64|250425"));
        Assert.False(MyDevicesScope.Holds(MyDevicesScope.CurrentHardware, "|250407", "|250425"));

        Assert.Equal(("Gone hardware", "Gone board"), CRT.Main.FormatBoardKeyAsFullNames("Gone hardware|Gone board"));
        Assert.Equal(("Gone hardware|", string.Empty), CRT.Main.FormatBoardKeyAsFullNames("Gone hardware|"));
    }
}
