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
        // The repairs line under a device's name, as parts so the numbers can be bold: "3 repairs -
        // 2 hours - 45 DKK - 1 open". Time and cost by the workbook headline's own rule
        // (WorkbookSummary.BuildTimeAndCostStats); "open" counts REPAIRS here.
        // ###########################################################################################
        public static IReadOnlyList<WorkbookSummary.Stat> HeadlineStats(History history, string? currencyCode)
        {
            var stats = new List<WorkbookSummary.Stat>
            {
                new(string.Empty, history.RepairCount.ToString(CultureInfo.InvariantCulture), history.RepairCount == 1 ? " repair" : " repairs"),
            };

            stats.AddRange(WorkbookSummary.BuildTimeAndCostStats(history.Totals, currencyCode));

            stats.Add(new WorkbookSummary.Stat(string.Empty, history.OpenRepairCount.ToString(CultureInfo.InvariantCulture), " open"));
            return stats;
        }
    }
}
