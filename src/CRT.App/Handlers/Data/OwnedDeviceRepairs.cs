using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A DEVICE'S HISTORY, WORKED OUT FROM THE WORKBOOKS (CRT 4.0 step 4 - the plan: "Its page is its
    // history: its repairs (workbooks), with time and cost totalled across all of them"). Which
    // workbooks are a device's, which device a workbook is on, which devices it can be put on, and
    // what the repairs add up to.
    //
    // Pure: handed the records, never reading a folder - the Workbooks tab has read them already.
    // ###########################################################################################
    public static class OwnedDeviceRepairs
    {
        // A device's repairs, newest first (the workbook list's own order).
        public static List<WorkbookRecord> Of(int deviceId, IEnumerable<WorkbookRecord> workbooks) =>
            workbooks
                .Where(workbook => workbook.OwnedDeviceId == deviceId)
                .OrderByDescending(workbook => workbook.Id)
                .ToList();

        // The device a workbook is on, or null - an id no device has any more (it was deleted; ids are
        // never handed out again) is none.
        public static OwnedDeviceRecord? DeviceOf(WorkbookRecord? workbook, IEnumerable<OwnedDeviceRecord> devices) =>
            workbook?.OwnedDeviceId is { } id ? devices.FirstOrDefault(device => device.Id == id) : null;

        // ###########################################################################################
        // The repair a device is worked on through when it is picked beside the board drop-downs (the
        // plan: "Choosing a machine selects its board and its open repair"): its newest OPEN repair,
        // on whichever board that was made; with none open, its newest repair on the board it has
        // now; else none - it has no repair to work on yet.
        //
        // Only a repair on a board CRT can show counts (`canShow`, by board key): one on a board a
        // later sync removed cannot be opened, and picking the device did nothing at all (code
        // review, 2026-10-29).
        // ###########################################################################################
        public static WorkbookRecord? RepairToWorkOn(OwnedDeviceRecord device, IEnumerable<WorkbookRecord> workbooks, Func<string, bool> canShow)
        {
            List<WorkbookRecord> repairs = Of(device.Id, workbooks).Where(repair => canShow(repair.BoardKey)).ToList();

            return repairs.FirstOrDefault(repair => WorklogManager.IsWorkbookStatusOpen(repair.Status))
                ?? repairs.FirstOrDefault(repair => string.Equals(repair.BoardKey, device.BoardKey, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // THE DEVICE PICKED - the worklog bar's "Device" drop-down and the one marked in "My devices", ONE choice
        // (owner requests, 2026-11-06: a device picked shows "all workbooks for this device"; "Show workbooks for
        // all devices" the first choice, "per default if never set before"; then "When in 'My Device' and I have
        // selected a device, then I do expect that it will only show the workbooks for this device" and, asked,
        // "Only the device's" - CRT works only in its workbooks while it is picked) - null for all devices.
        //
        // The one saved, while it is a device there is and the board shown is its own or holds one of its
        // workbooks: another board shown - the drop-downs at the top switched - puts the bar back on all devices,
        // rather than work in nothing there. A device whose board CRT cannot show (`canShow`) has no board to go
        // to, and stays picked till something else is - and so does any while NO board is shown: switching hardware
        // empties the boards' list before the board is picked, and the bar drawn in that moment let go of a device
        // picked on another hardware on its way to its own board (owner report, 2026-11-06: another device clicked
        // "does the exact same thing as 'Clear'").
        // ###########################################################################################
        public static OwnedDeviceRecord? BarFilter(
            int? savedDeviceId,
            IEnumerable<OwnedDeviceRecord> devices,
            string? boardShown,
            IEnumerable<WorkbookRecord> workbooks,
            Func<string, bool> canShow)
        {
            if (savedDeviceId is not { } id || devices.FirstOrDefault(device => device.Id == id) is not { } device)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(boardShown) || SameBoard(device.BoardKey, boardShown) || !canShow(device.BoardKey))
            {
                return device;
            }

            return workbooks.Any(workbook => workbook.OwnedDeviceId == device.Id && SameBoard(workbook.BoardKey, boardShown)) ? device : null;
        }

        // ###########################################################################################
        // The workbooks of the board shown CRT works in: every one, or only the device picked's (owner
        // decision, 2026-11-06: "Only the device's") - "the active workbook" is resolved among these
        // (WorklogManager.ResolveActiveWorkbook), so with none of the device's on the board there is none,
        // and "Add worklog" has nothing to write into rather than write into another device's.
        // ###########################################################################################
        public static List<WorkbookRecord> WorkedIn(IEnumerable<WorkbookRecord> boardWorkbooks, int? pickedDeviceId) =>
            pickedDeviceId is { } id
                ? boardWorkbooks.Where(workbook => workbook.OwnedDeviceId == id).ToList()
                : boardWorkbooks.ToList();

        // ###########################################################################################
        // The device picked once `workbook` is activated: kept while the workbook is on it (all devices hold
        // every one), else the workbook's own device - or all, for one on no device. A workbook made in "Add
        // workbook" on another device, or none, is then the one worked in rather than one nothing shows.
        // ###########################################################################################
        public static int? PickFollowing(int? pickedDeviceId, WorkbookRecord? workbook, IEnumerable<OwnedDeviceRecord> devices)
        {
            if (pickedDeviceId is null || workbook?.OwnedDeviceId == pickedDeviceId)
            {
                return pickedDeviceId;
            }

            return DeviceOf(workbook, devices)?.Id;
        }

        // Two board keys naming one board - trimmed, capitals aside, as every board-key lookup does.
        private static bool SameBoard(string? one, string? other) =>
            (one?.Trim() ?? string.Empty) is { Length: > 0 } key &&
            string.Equals(key, other?.Trim(), StringComparison.OrdinalIgnoreCase);

        // ###########################################################################################
        // The devices a workbook can be put on: those with the workbook's board (a repair was made
        // on a board, and a device has one), plus the device it is on already - its board may have
        // been swapped since. Capitals in the board key aside, as every board-key lookup does.
        // ###########################################################################################
        public static List<OwnedDeviceRecord> ChoicesFor(WorkbookRecord workbook, IEnumerable<OwnedDeviceRecord> devices) =>
            devices
                .Where(device => device.Id == workbook.OwnedDeviceId ||
                                 string.Equals(device.BoardKey, workbook.BoardKey, StringComparison.OrdinalIgnoreCase))
                .ToList();

        // ###########################################################################################
        // What a device's repairs add up to: how many, how many still open, and the worklogs' own
        // totals across all of them - ONE WorkbookSummary over every worklog, so the device page and
        // a workbook's strip count hours and cost by the same rule.
        // ###########################################################################################
        public sealed record History(int RepairCount, int OpenRepairCount, WorkbookSummary.Totals Totals);

        public static History Summarize(IReadOnlyList<WorkbookRecord> repairs, Func<int, IEnumerable<WorklogEntryRecord>> entriesOf) =>
            Summarize(repairs.Select(repair => new Repair(repair.Status, entriesOf(repair.Id))));

        // One repair as it is counted: its status and its worklogs - a workbook record's, or an
        // exported history's document's (DeviceHistoryExportModel), so both count by this one rule.
        public sealed record Repair(string Status, IEnumerable<WorklogEntryRecord> Worklogs);

        public static History Summarize(IEnumerable<Repair> repairs)
        {
            List<Repair> counted = repairs.ToList();

            return new History(
                counted.Count,
                counted.Count(repair => WorklogManager.IsWorkbookStatusOpen(repair.Status)),
                WorkbookSummary.Summarize(counted.SelectMany(repair => repair.Worklogs)));
        }

        // ###########################################################################################
        // The repairs line under a device's name, as parts so the numbers can be bold: "3 workbooks -
        // 2 hours - 45 DKK - 1 open" - workbooks, as its heading "Workbooks associated to this device" says
        // (owner request, 2026-11-06). Time and cost by the workbook headline's own rule
        // (WorkbookSummary.BuildTimeAndCostStats); "open" counts REPAIRS here.
        // ###########################################################################################
        public static IReadOnlyList<WorkbookSummary.Stat> HeadlineStats(History history, string? currencyCode)
        {
            var stats = new List<WorkbookSummary.Stat>
            {
                new(string.Empty, history.RepairCount.ToString(CultureInfo.InvariantCulture), history.RepairCount == 1 ? " workbook" : " workbooks"),
            };

            stats.AddRange(WorkbookSummary.BuildTimeAndCostStats(history.Totals, currencyCode));

            stats.Add(new WorkbookSummary.Stat(string.Empty, history.OpenRepairCount.ToString(CultureInfo.InvariantCulture), " open"));
            return stats;
        }
    }
}
