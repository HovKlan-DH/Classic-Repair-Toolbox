using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// "SORTING OF DEVICES" (owner request, 2026-11-07: "Sort automatically by name", "Sort automatically by
// hardware/board/name", "Sort manually in list (drag up/down)" - by name the default) - the order "My devices"
// and the worklog bar's "Device" drop-down list the devices in.
// ###########################################################################################
public sealed class OwnedDeviceSortingTests
{
    private static OwnedDeviceRecord Device(int id, string name, string boardKey = "Commodore 64|250407", int place = 0) =>
        new() { Id = id, Name = name, BoardKey = boardKey, DisplayOrder = place };

    // The drop-downs' names for a board key: the key's own two halves, as Main's lookup gives them for a listed board.
    private static (string Hardware, string Board) Names(string key)
    {
        string[] parts = key.Split('|');
        return (parts[0], parts.Length > 1 ? parts[1] : string.Empty);
    }

    private static List<string> Sorted(IEnumerable<OwnedDeviceRecord> devices, string? sorting) =>
        OwnedDeviceSorting.Sort(devices, sorting, Names).Select(device => device.Name).ToList();

    [Theory]
    [InlineData(null, OwnedDeviceSorting.ByName)]
    [InlineData("", OwnedDeviceSorting.ByName)]
    [InlineData("something else", OwnedDeviceSorting.ByName)]
    [InlineData("manual", OwnedDeviceSorting.Manual)]
    [InlineData("HARDWAREBOARDNAME", OwnedDeviceSorting.ByHardwareBoardName)]
    public void A_stored_value_is_one_of_the_three_and_anything_else_is_by_name(string? stored, string expected)
    {
        Assert.Equal(expected, OwnedDeviceSorting.Normalise(stored));
    }

    // By name, numbers read as numbers - "#2" before "#10" - capitals aside; a tie by the order they were added.
    [Fact]
    public void By_name_numbers_read_as_numbers_and_a_tie_goes_by_the_order_added()
    {
        OwnedDeviceRecord[] devices =
        [
            Device(1, "C64 #10"),
            Device(2, "c64 #2"),
            Device(3, "Amiga 500"),
            Device(5, "Breadbin"),
            Device(4, "Breadbin"),
        ];

        Assert.Equal(["Amiga 500", "Breadbin", "Breadbin", "c64 #2", "C64 #10"], Sorted(devices, OwnedDeviceSorting.ByName));
        Assert.Equal([4, 5], OwnedDeviceSorting.Sort(devices, OwnedDeviceSorting.ByName, Names).Where(device => device.Name == "Breadbin").Select(device => device.Id));
    }

    // By hardware, then board, then name - as the drop-downs name them.
    [Fact]
    public void By_hardware_board_and_name_groups_the_devices_by_their_hardware_and_board()
    {
        OwnedDeviceRecord[] devices =
        [
            Device(1, "Zed", "Commodore 64|250425"),
            Device(2, "Alpha", "Commodore 64|250425"),
            Device(3, "Middle", "Commodore 64|250407"),
            Device(4, "First", "Amiga 500|Rev 6A"),
        ];

        Assert.Equal(["First", "Middle", "Alpha", "Zed"], Sorted(devices, OwnedDeviceSorting.ByHardwareBoardName));
    }

    // By hand: the places dragged, then every device never placed in the order it was added.
    [Fact]
    public void By_hand_the_places_dragged_come_first_and_one_never_placed_after_them_as_added()
    {
        OwnedDeviceRecord[] devices =
        [
            Device(1, "Last added before", place: 0),
            Device(2, "Second", place: 2),
            Device(3, "First", place: 1),
            Device(4, "Added later", place: 0),
        ];

        Assert.Equal(["First", "Second", "Last added before", "Added later"], Sorted(devices, OwnedDeviceSorting.Manual));
    }

    // ###########################################################################################
    // A drop under a scope or a search moves only the devices listed: the places they stood in are theirs again in the
    // new order, every other device keeps its own.
    // ###########################################################################################
    [Theory]
    [InlineData(new[] { 1, 2, 3, 4 }, new[] { 3, 1 }, new[] { 3, 2, 1, 4 })]
    [InlineData(new[] { 1, 2, 3, 4 }, new[] { 4, 3, 2, 1 }, new[] { 4, 3, 2, 1 })]
    [InlineData(new[] { 1, 2, 3 }, new[] { 2 }, new[] { 1, 2, 3 })]
    [InlineData(new[] { 1, 2, 3 }, new[] { 9, 3, 1 }, new[] { 3, 2, 1 })]
    public void A_drop_moves_the_devices_listed_and_every_other_keeps_its_place(int[] all, int[] listedNow, int[] expected)
    {
        Assert.Equal(expected, OwnedDeviceSorting.Reordered(all, listedNow));
    }
}
