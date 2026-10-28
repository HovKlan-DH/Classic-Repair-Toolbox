using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// OwnedDeviceRepairs - a device's history worked out from the workbooks (CRT 4.0 step 4): which
// are its repairs, which are on no device, which devices a workbook can go on, and their totals.
// ###########################################################################################
public sealed class OwnedDeviceRepairsTests
{
    private const string Board = "Commodore 64|250407";

    private static WorkbookRecord Workbook(int id, int? deviceId, string board = Board, string status = "Open") =>
        new() { Id = id, BoardKey = board, Title = $"Repair {id}", OwnedDeviceId = deviceId, Status = status };

    private static OwnedDeviceRecord Device(int id, string board = Board) =>
        new() { Id = id, Name = $"Device {id}", BoardKey = board };

    [Fact]
    public void A_devices_repairs_are_the_workbooks_naming_it_newest_first()
    {
        WorkbookRecord[] workbooks = [Workbook(1, 7), Workbook(2, null), Workbook(3, 7), Workbook(4, 8)];

        Assert.Equal([3, 1], OwnedDeviceRepairs.Of(7, workbooks).Select(w => w.Id));
    }

    // An id no device has (the device was deleted) is on no device, as an empty one is.
    [Fact]
    public void A_repair_naming_a_device_that_is_gone_is_on_no_device()
    {
        Assert.Null(OwnedDeviceRepairs.DeviceOf(Workbook(3, 99), [Device(7)]));
        Assert.Null(OwnedDeviceRepairs.DeviceOf(Workbook(2, null), [Device(7)]));
        Assert.Equal(7, OwnedDeviceRepairs.DeviceOf(Workbook(1, 7), [Device(7)])!.Id);
    }

    // ###########################################################################################
    // A workbook goes on a device with its board (capitals aside) - and stays offered the device it
    // is on, whose board may have been swapped since.
    // ###########################################################################################
    [Fact]
    public void A_workbook_can_go_on_a_device_with_its_board_or_the_one_it_is_on()
    {
        OwnedDeviceRecord same = Device(1, "commodore 64|250407");
        OwnedDeviceRecord other = Device(2, "Commodore 64|250425");
        OwnedDeviceRecord swapped = Device(3, "Commodore 128|310378");

        Assert.Equal([1], OwnedDeviceRepairs.ChoicesFor(Workbook(1, null), [same, other, swapped]).Select(d => d.Id));
        Assert.Equal([1, 3], OwnedDeviceRepairs.ChoicesFor(Workbook(1, 3), [same, other, swapped]).Select(d => d.Id));
    }

    // ###########################################################################################
    // The totals across a device's repairs: every worklog's hours and cost added up, the repairs
    // counted and the open ones among them.
    // ###########################################################################################
    [Fact]
    public void A_devices_history_adds_up_every_repairs_worklogs()
    {
        WorkbookRecord[] repairs = [Workbook(1, 7), Workbook(2, 7, status: "Closed")];
        var entries = new Dictionary<int, List<WorklogEntryRecord>>
        {
            [1] = [new WorklogEntryRecord { WorkDoneItems = [new WorklogWorkDoneRecord { HoursSpent = 1.5, Cost = 20 }] }],
            [2] = [new WorklogEntryRecord { State = "Closed", WorkDoneItems = [new WorklogWorkDoneRecord { HoursSpent = 0.5, Cost = 25 }] }],
        };

        OwnedDeviceRepairs.History history = OwnedDeviceRepairs.Summarize(repairs, id => entries[id]);

        Assert.Equal((2, 1), (history.RepairCount, history.OpenRepairCount));
        Assert.Equal(2.0, history.Totals.TotalHours, 3);
        Assert.Equal(45.0, history.Totals.TotalCost, 3);
        Assert.Equal(2, history.Totals.WorklogCount);

        string line = string.Concat(OwnedDeviceRepairs.HeadlineStats(history, "DKK").Select(stat => stat.Prefix + stat.Number + stat.Suffix + " "));
        Assert.StartsWith("2 repairs ", line, StringComparison.Ordinal);
        Assert.Contains("45 DKK", line, StringComparison.Ordinal);
        Assert.EndsWith("1 open ", line, StringComparison.Ordinal);
    }

    // The time and the cost are the workbook headline's own parts - one rule (code review, 2026-10-28).
    [Fact]
    public void A_devices_time_and_cost_are_spelt_as_a_workbooks_are()
    {
        var entries = new List<WorklogEntryRecord>
        {
            new() { WorkDoneItems = [new WorklogWorkDoneRecord { HoursSpent = 1.25, Cost = 12.5 }] },
        };

        OwnedDeviceRepairs.History history = OwnedDeviceRepairs.Summarize([Workbook(1, 7)], _ => entries);
        IReadOnlyList<WorkbookSummary.Stat> line = OwnedDeviceRepairs.HeadlineStats(history, "EUR");

        Assert.Equal(WorkbookSummary.BuildTimeAndCostStats(history.Totals, "EUR"), line.Skip(1).Take(line.Count - 2));
    }

    // No time and no cost logged: the line says only what there is.
    [Fact]
    public void A_devices_line_leaves_out_time_and_cost_it_does_not_have()
    {
        OwnedDeviceRepairs.History history = OwnedDeviceRepairs.Summarize([Workbook(1, 7)], _ => []);

        Assert.Equal(
            ["1 repair", "1 open"],
            OwnedDeviceRepairs.HeadlineStats(history, "DKK").Select(stat => stat.Prefix + stat.Number + stat.Suffix));
    }
}
