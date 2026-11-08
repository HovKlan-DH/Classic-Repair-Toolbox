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
    // A device picked beside the board drop-downs is worked on through its newest OPEN repair, on
    // whichever board that was made - the plan: "Choosing a machine selects its board and its open
    // repair". Other devices' repairs never count.
    // ###########################################################################################
    [Fact]
    public void A_device_is_worked_on_through_its_newest_open_repair_on_any_board()
    {
        WorkbookRecord[] workbooks =
        [
            Workbook(1, 7, status: "Open"),
            Workbook(2, 7, board: "Commodore 64|250425", status: "Open"),
            Workbook(3, 7, status: "Closed"),
            Workbook(4, 8, status: "Open"),
            Workbook(5, null, status: "Open"),
        ];

        Assert.Equal(2, OwnedDeviceRepairs.RepairToWorkOn(Device(7), workbooks, Anywhere)!.Id);
    }

    // Every board can be shown - the rule's question, answered for the cases that are not about it.
    private static bool Anywhere(string boardKey) => true;

    // ###########################################################################################
    // A repair on a board CRT cannot show (a later sync removed it) cannot be opened, so it is not
    // the one worked on - the next one is (code review, 2026-10-29: picking the device did nothing).
    // ###########################################################################################
    [Fact]
    public void A_repair_on_a_board_CRT_cannot_show_is_passed_over()
    {
        WorkbookRecord[] workbooks =
        [
            Workbook(1, 7, status: "Closed"),
            Workbook(2, 7, board: "Commodore 64|gone", status: "Open"),
        ];

        bool CanShow(string boardKey) => boardKey != "Commodore 64|gone";

        Assert.Equal(1, OwnedDeviceRepairs.RepairToWorkOn(Device(7), workbooks, CanShow)!.Id);
        Assert.Null(OwnedDeviceRepairs.RepairToWorkOn(Device(7), workbooks, _ => false));
    }

    // With none open, its newest repair on the board it has NOW (capitals aside) - one on a board
    // swapped out is no longer the device being worked on. With neither, none: there is no repair
    // to work on yet.
    [Fact]
    public void With_nothing_open_a_device_is_worked_on_through_its_newest_repair_on_its_board()
    {
        WorkbookRecord[] workbooks =
        [
            Workbook(1, 7, status: "Closed"),
            Workbook(2, 7, board: "COMMODORE 64|250407", status: "Closed"),
            Workbook(3, 7, board: "Commodore 64|250425", status: "Closed"),
            Workbook(4, 8, status: "Open"),
        ];

        Assert.Equal(2, OwnedDeviceRepairs.RepairToWorkOn(Device(7), workbooks, Anywhere)!.Id);
        Assert.Null(OwnedDeviceRepairs.RepairToWorkOn(Device(7, board: "Commodore 128|310378"), workbooks, Anywhere));
        Assert.Null(OwnedDeviceRepairs.RepairToWorkOn(Device(9), workbooks, Anywhere));
    }

    // ###########################################################################################
    // THE DEVICE PICKED (owner requests, 2026-11-06: the bar's "Device" drop-down and the device selected in "My
    // devices", one choice; asked, CRT works "Only the device's" workbooks while one is): the device saved, while
    // there is such a device and the board shown is its own (capitals aside) or holds one of its workbooks - another
    // board shown is all devices (null) again, rather than nothing worked in there. A device on a board CRT cannot
    // show has no board to go to and stays. Nothing saved is all devices, the default.
    // ###########################################################################################
    [Fact]
    public void The_device_picked_holds_while_the_board_shown_is_its_own_or_holds_its_workbooks()
    {
        const string Other = "Commodore 64|250425";
        OwnedDeviceRecord[] devices = [Device(7), Device(8, Other), Device(9, "Gone|board")];
        WorkbookRecord[] workbooks = [Workbook(1, 7, Other), Workbook(2, 8)];
        bool CanShow(string board) => board is Board or Other;

        Assert.Null(OwnedDeviceRepairs.BarFilter(null, devices, Board, workbooks, CanShow));

        // On its own board, whatever workbooks there are - none of its own included.
        Assert.Equal(7, OwnedDeviceRepairs.BarFilter(7, devices, Board, [], CanShow)!.Id);
        Assert.Equal(7, OwnedDeviceRepairs.BarFilter(7, devices, " commodore 64|250407 ", workbooks, CanShow)!.Id);

        // Another board: while one of its workbooks is there (a board swapped since), else all devices.
        Assert.Equal(7, OwnedDeviceRepairs.BarFilter(7, devices, Other, workbooks, CanShow)!.Id);
        Assert.Equal(8, OwnedDeviceRepairs.BarFilter(8, devices, Board, workbooks, CanShow)!.Id);
        Assert.Null(OwnedDeviceRepairs.BarFilter(8, devices, Board, [], CanShow));
        Assert.Null(OwnedDeviceRepairs.BarFilter(7, devices, "Commodore 128|310378", workbooks, CanShow));

        // A board CRT cannot show: nowhere to go - it stays picked.
        Assert.Equal(9, OwnedDeviceRepairs.BarFilter(9, devices, Board, [], CanShow)!.Id);

        // No board shown - the moment between a hardware and its board, on the way to the device's own: it stays.
        Assert.Equal(8, OwnedDeviceRepairs.BarFilter(8, devices, "", [], CanShow)!.Id);
        Assert.Equal(8, OwnedDeviceRepairs.BarFilter(8, devices, null, [], CanShow)!.Id);

        // A device that has gone: all devices.
        Assert.Null(OwnedDeviceRepairs.BarFilter(10, devices, Board, workbooks, CanShow));
    }

    // ###########################################################################################
    // While a device is picked the workbooks worked in are its own (owner decision, 2026-11-06: "Only the
    // device's") - the active one is resolved among them, so with none there is none; all devices, all of them.
    // ###########################################################################################
    [Fact]
    public void The_workbooks_worked_in_are_the_device_picked_s_or_all()
    {
        WorkbookRecord[] onBoard = [Workbook(3, 7), Workbook(2, null), Workbook(1, 8)];

        Assert.Equal([3, 2, 1], OwnedDeviceRepairs.WorkedIn(onBoard, null).Select(w => w.Id));
        Assert.Equal([3], OwnedDeviceRepairs.WorkedIn(onBoard, 7).Select(w => w.Id));
        Assert.Empty(OwnedDeviceRepairs.WorkedIn(onBoard, 9));
        Assert.Null(WorklogManager.ResolveActiveWorkbook(OwnedDeviceRepairs.WorkedIn(onBoard, 9), savedActiveId: 2));
        Assert.Equal(3, WorklogManager.ResolveActiveWorkbook(OwnedDeviceRepairs.WorkedIn(onBoard, 7), savedActiveId: 2)!.Id);
    }

    // ###########################################################################################
    // A workbook activated takes the pick along when it is not on the device picked - its own device, or all
    // devices for one on none (or on a device gone); the device's own keeps it, and all devices hold every one.
    // ###########################################################################################
    [Fact]
    public void A_workbook_activated_off_the_device_picked_takes_the_pick_to_its_own()
    {
        OwnedDeviceRecord[] devices = [Device(7), Device(8)];

        Assert.Equal(7, OwnedDeviceRepairs.PickFollowing(7, Workbook(1, 7), devices));
        Assert.Equal(8, OwnedDeviceRepairs.PickFollowing(7, Workbook(1, 8), devices));
        Assert.Null(OwnedDeviceRepairs.PickFollowing(7, Workbook(1, null), devices));
        Assert.Null(OwnedDeviceRepairs.PickFollowing(7, Workbook(1, 12), devices));
        Assert.Null(OwnedDeviceRepairs.PickFollowing(null, Workbook(1, 8), devices));
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
        Assert.StartsWith("2 workbooks ", line, StringComparison.Ordinal);
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
            ["1 workbook", "1 open"],
            OwnedDeviceRepairs.HeadlineStats(history, "DKK").Select(stat => stat.Prefix + stat.Number + stat.Suffix));
    }
}
