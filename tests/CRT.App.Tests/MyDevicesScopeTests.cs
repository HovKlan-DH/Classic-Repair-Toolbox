using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// What the "Workbooks" tab lists (MyDevicesScope; owner request, 2026-11-06: "Show all devices", "Show only devices
// for selected hardware" in the middle, "Show only devices for selected board" - the workbook list narrowing the same
// way). A board key is "Hardware|Board", capitals aside.
// ###########################################################################################
public sealed class MyDevicesScopeTests
{
    private const string Shown = "Commodore 64|250407";

    [Theory]
    [InlineData("Commodore 64|250407", true)]
    [InlineData("commodore 64|250407", true)]
    [InlineData("Commodore 64|250425", false)]
    [InlineData("Commodore 128|310378", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void The_board_shown_lists_only_that_board(string? boardKey, bool listed)
    {
        Assert.Equal(listed, MyDevicesScope.Holds(MyDevicesScope.CurrentBoard, boardKey, Shown));
    }

    // Every board of the hardware shown - the C64's long and short boards alike - and none of another hardware.
    [Theory]
    [InlineData("Commodore 64|250407", true)]
    [InlineData("Commodore 64|250425", true)]
    [InlineData("COMMODORE 64|250469", true)]
    [InlineData("Commodore 64C|250469", false)]
    [InlineData("Commodore 128|310378", false)]
    [InlineData("", false)]
    public void The_hardware_shown_lists_every_board_of_it(string boardKey, bool listed)
    {
        Assert.Equal(listed, MyDevicesScope.Holds(MyDevicesScope.CurrentHardware, boardKey, Shown));
    }

    [Theory]
    [InlineData("Commodore 128|310378")]
    [InlineData("")]
    [InlineData(null)]
    public void All_lists_everything(string? boardKey)
    {
        Assert.True(MyDevicesScope.Holds(MyDevicesScope.AllBoards, boardKey, Shown));
    }

    // Stored as the workbooks' scope always was; anything unknown is the board shown, the default - which is also
    // how CRT 3.x, reading the same file, takes "CurrentHardware".
    [Theory]
    [InlineData("AllBoards", MyDevicesScope.AllBoards)]
    [InlineData("currenthardware", MyDevicesScope.CurrentHardware)]
    [InlineData("CurrentBoard", MyDevicesScope.CurrentBoard)]
    [InlineData("nonsense", MyDevicesScope.CurrentBoard)]
    [InlineData(null, MyDevicesScope.CurrentBoard)]
    public void A_scope_is_one_of_three(string? stored, string expected)
    {
        Assert.Equal(expected, MyDevicesScope.Normalise(stored));
        Assert.Equal(MyDevicesScope.Holds(expected, "Commodore 64|250425", Shown), MyDevicesScope.Holds(stored, "Commodore 64|250425", Shown));
    }
}
