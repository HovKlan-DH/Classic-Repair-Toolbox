using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT CRT SAYS ABOUT THE USER'S DEVICES (CRT 4.0 step 4) - "Device" wherever a person reads it
    // (owner decision, 2026-10-24: "My devices", "Add device"). Every sentence the devices' screens
    // show is here, so the tab, the dialogs and the tests say the same.
    // ###########################################################################################
    public static class OwnedDeviceWording
    {
        public const string ListHeading = "My devices";

        public const string NoDevicesYet =
            "Add the computers you own, to keep each one's repairs, changes and photos together.";

        public const string NameNeeded = "A device needs a name.";

        public const string NoRepairsYet =
            "No repairs on this device yet. \"New repair\" starts one on its board.";

        public const string NoChangesYet =
            "Nothing recorded yet - a replaced power supply, a fitted ROM or a swapped board goes here.";

        public const string NoPhotosYet = "No photos yet.";

        // ###########################################################################################
        // The picker beside the board drop-downs (slice 4c): its heading, and what it shows while the
        // workbook in the bar is on no device - plus the choice "Create new workbook" and "On my
        // device" put before the devices.
        // ###########################################################################################
        public const string PickerHeading = "Device";

        public const string NoDevicePicked = "No device";

        public const string NoDeviceChoice = "None";

        // "Create new workbook"'s title for a repair started on a device ("New repair", or a device
        // picked that has none yet).
        public static string NewRepairTitle(string deviceName) => $"New repair on {deviceName.Trim()}";

        // The worklog bar's words for a repair on a device - "Breadbin #2 - No video" (the plan).
        public static string BarText(string deviceName, string repairTitle) =>
            repairTitle.Trim().Length == 0 ? deviceName.Trim() : $"{deviceName.Trim()} - {repairTitle.Trim()}";

        // ###########################################################################################
        // Said after "Add photos": how many went in, and how many could not (code review, 2026-10-28:
        // photos that could not be copied were dropped without a word).
        // ###########################################################################################
        public static string PhotosAdded(int added, int notAdded)
        {
            string done = added == 1 ? "Added 1 photo." : $"Added {added.ToString(CultureInfo.InvariantCulture)} photos.";

            return notAdded switch
            {
                <= 0 => done,
                1 => $"{done} 1 could not be added - it is no picture CRT can show, open in another program, or no longer where it was picked.",
                _ => $"{done} {notAdded.ToString(CultureInfo.InvariantCulture)} could not be added - they are no pictures CRT can show, open in another program, or no longer where they were picked.",
            };
        }

        // A board key ("Commodore 64|250407") as a person reads it, "Commodore 64 - 250407".
        public static string BoardText(string? boardKey)
        {
            string key = boardKey?.Trim() ?? string.Empty;
            return key.Length == 0 ? "no board" : key.Replace("|", " - ", StringComparison.Ordinal);
        }

        // The change recorded when a device's board is swapped (OwnedDeviceStore.Update).
        public static string BoardSwapped(string fromBoardKey, string toBoardKey) =>
            $"Board changed from {BoardText(fromBoardKey)} to {BoardText(toBoardKey)}";

        // ###########################################################################################
        // The line under a device's name: its board as the drop-downs name it, then its region and
        // serial number when it has them - "Commodore 64 - 250407 - PAL - serial CA123456".
        // ###########################################################################################
        public static string DetailLine(OwnedDeviceRecord device, string boardText)
        {
            var parts = new List<string> { boardText };

            if (device.Region.Trim().Length > 0)
            {
                parts.Add(device.Region.Trim());
            }

            if (device.SerialNumber.Trim().Length > 0)
            {
                parts.Add("serial " + device.SerialNumber.Trim());
            }

            return string.Join(" - ", parts.Where(part => part.Length > 0));
        }

        public static string OriginLine(string? origin) =>
            origin?.Trim() is { Length: > 0 } text ? "Where it came from: " + text : string.Empty;

        // No count: the line under it starts with "2 repairs" (OwnedDeviceRepairs.HeadlineStats).
        public const string RepairsHeading = "Repairs";

        // ###########################################################################################
        // "Export history" (slice 4d): what the exported document calls itself, and its list of
        // repairs - "#12 No video (Open, started 2026-09-28, 3 worklogs)", or none recorded. In
        // brackets, as a title may hold " - " itself ("No video - VIC socket").
        // ###########################################################################################
        public const string HistoryLabel = "Device history";

        public const string NoRepairsRecorded = "No repairs recorded.";

        public static string HistoryRepairLine(int workbookId, string title, string status, string dateText, int worklogCount)
        {
            string worklogs = worklogCount == 1 ? "1 worklog" : $"{worklogCount.ToString(CultureInfo.InvariantCulture)} worklogs";
            string name = title.Trim().Length == 0 ? "(untitled)" : title.Trim();

            return $"#{workbookId.ToString(CultureInfo.InvariantCulture)} {name} ({status}, {dateText}, {worklogs})";
        }

        public static string ChangesHeading(int count) =>
            count == 0 ? "Changes made to it" : $"Changes made to it ({count.ToString(CultureInfo.InvariantCulture)})";

        public static string PhotosHeading(int count) =>
            count == 0 ? "Notes and photos" : $"Notes and photos ({count.ToString(CultureInfo.InvariantCulture)} {(count == 1 ? "photo" : "photos")})";

        // A change as its row reads it - "2025-06-14  JiffyDOS ROM fitted".
        public static string ChangeDate(DateTime date) =>
            date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

        // ###########################################################################################
        // Asked before a device is deleted: what goes (its photos, notes and changes) and what stays
        // (its repairs, which are workbooks of their own) - and where they stay: in the workbook list,
        // on no device (code review, 2026-10-28: it named a "repairs without a device" list CRT does
        // not have).
        // ###########################################################################################
        public static string DeleteQuestion(string name, int repairCount)
        {
            string repairs = repairCount switch
            {
                0 => "It has no repairs.",
                1 => "Its 1 repair stays in the workbook list, on no device.",
                _ => $"Its {repairCount.ToString(CultureInfo.InvariantCulture)} repairs stay in the workbook list, on no device.",
            };

            return $"Delete [{name.Trim()}]? Its photos, notes and changes are deleted with it. {repairs}";
        }
    }
}
