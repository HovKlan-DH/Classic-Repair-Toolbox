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

        // Under "My devices" while there is none - naming the button that adds one, so a device is seen to be
        // added before any workbook (owner report, 2026-11-06: "it is not logical that I need to create a
        // workbook to create a new device"). Above the tabs, the worklog bar's: the one beside "My devices"
        // went that day (owner request: "no need for it").
        public const string NoDevicesYet =
            "\"Add device\" above the tabs adds a computer you own, to keep its workbooks, comments and photos together.";

        // ###########################################################################################
        // The handle on a card in "My devices" while the devices are sorted by hand (owner request, 2026-11-07: "Sort
        // manually in list (drag up/down)").
        // ###########################################################################################
        public const string DragDeviceTip = "Drag to move this device up or down";

        // Under "My devices" while devices there are none of the Configuration tab's scope lists (MyDevicesScope) -
        // so the list is not read as empty.
        public static string NoDevicesInScope(string scope) => MyDevicesScope.Normalise(scope) switch
        {
            MyDevicesScope.CurrentHardware => "No devices for the selected hardware.",
            MyDevicesScope.CurrentBoard => "No devices for the selected board.",
            _ => "No devices.",
        };

        // Under "My devices" while a search finds none of them (WorklogSearchIndex.ForDevice).
        public const string NoDevicesMatch = "No devices match your search.";

        public const string NameNeeded = "A device needs a name.";

        // A device's page with no workbook on the device - said, though the page is only read (owner request, 2026-11-07:
        // "It is fine though, that it shows that no workbooks are associated"; its "New repair" button went that day).
        public const string NoRepairsYet = "No workbooks associated to this device yet.";

        // ###########################################################################################
        // "Clear" at the foot of "My devices" (owner request, 2026-11-06: 'at bottom of the "My devices"
        // panel ... there should be a "Clear" ... clicking that will unselect the device and by then show
        // all again').
        // ###########################################################################################
        public const string ClearChoice = "Clear";

        // The workbook list on the "Workbooks" tab while a device is picked and it has none.
        public const string NoWorkbooksOfDevice =
            "No workbooks associated to this device yet.\nUse \"Add workbook\" above the tabs to start one.";

        // ###########################################################################################
        // The worklog bar's "Workbook" drop-down while a device is picked and the board shown has no
        // workbook of it - CRT works only in the device's (owner decision, 2026-11-06: "Only the device's"):
        // none at all, or only on other boards, which the list then holds.
        // ###########################################################################################
        public const string NoWorkbookOfDeviceYet = "No workbook on this device yet";

        public const string NoWorkbookOfDeviceOnBoard = "No workbook of this device on this board";

        // ###########################################################################################
        // A device's comments - the form's list and the page's (owner request, 2026-11-06: "Comments" in the
        // device form, as the worklog's) - "Comments (2)".
        // ###########################################################################################
        public static string CommentsHeading(int count) =>
            count == 0 ? "Comments" : $"Comments ({count.ToString(CultureInfo.InvariantCulture)})";

        // ###########################################################################################
        // Its condition, 0 to 5 stars (owner request, 2026-11-06): what the form's stars say when pointed
        // at, and the exported history's line - which draws no star glyph a PDF font may lack.
        // ###########################################################################################
        // ###########################################################################################
        // Under "Hardware" and "Board" in "Edit device" while the device has workbooks (owner decision, 2026-11-06, asked:
        // "Locked while it has workbooks"): a workbook is associated only to a device of its own board, so the board
        // cannot change under them.
        // ###########################################################################################
        public static string BoardLocked(int workbookCount) =>
            (workbookCount == 1 ? "It has 1 workbook" : $"It has {workbookCount.ToString(CultureInfo.InvariantCulture)} workbooks") +
            ", so its hardware and board cannot change - a workbook is only associated to a device of its own board.";

        public const string ConditionHint = "Click a star to set the condition - click it again for none.";

        public static string ConditionLine(int stars)
        {
            int within = OwnedDeviceCondition.Within(stars);
            return $"Condition: {within.ToString(CultureInfo.InvariantCulture)} of {OwnedDeviceCondition.MaxStars.ToString(CultureInfo.InvariantCulture)} stars";
        }

        // ###########################################################################################
        // The device picker, the worklog bar's first drop-down (slice 4c; in the bar since 2026-11-06): its
        // label and its first choice - every workbook, whatever device it is on (owner request, 2026-11-06:
        // "rename 'No device' to 'Show workbooks for all devices' and then have that as first selection"; each
        // line's count of workbooks is a badge, WorklogInfoPillBuilder.BuildCountPill); plus the choice "Add
        // workbook" and "On my device" put before the devices.
        // ###########################################################################################
        public const string PickerHeading = "Device";

        public const string AllDevicesChoice = "Show workbooks for all devices";

        public const string NoDeviceChoice = "None";

        // "Add workbook"'s title for a repair started on a device ("New repair", or a device picked that
        // has none yet).
        public static string NewRepairTitle(string deviceName) => $"New repair on {deviceName.Trim()}";

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

        // ###########################################################################################
        // Its description - "Where it came from" until 2026-11-06 (owner request: 'change text "Where did it
        // come from" to "Description"') - several lines as typed, its line breaks kept, each "\n", so the
        // page and the exported PDF break where the user did. Shown as it is, under the line naming the
        // device: a description needs no label.
        // ###########################################################################################
        public static string DescriptionText(string? description) =>
            RichNote.TrimMarkup(description?.ReplaceLineEndings("\n"));

        // A comment's date and time, as the worklog's own comments show theirs.
        public static string CommentDate(DateTime date) => date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

        // ###########################################################################################
        // The workbooks on a device - its page's heading, and the exported history's (owner request,
        // 2026-11-06: "rename the 'Repairs' to 'Workbooks associated to this device'"). No count: the
        // line under it starts with "2 workbooks" (OwnedDeviceRepairs.HeadlineStats).
        // ###########################################################################################
        public const string RepairsHeading = "Workbooks associated to this device";

        // ###########################################################################################
        // "Export history" (slice 4d): what the exported document calls itself, and its list of
        // repairs - "#12 No video (Open, started 2026-09-28, 3 worklogs)", or none recorded. In
        // brackets, as a title may hold " - " itself ("No video - VIC socket").
        // ###########################################################################################
        public const string HistoryLabel = "Device history";

        public const string NoRepairsRecorded = "No workbooks recorded.";

        public static string HistoryRepairLine(int workbookId, string title, string status, string dateText, int worklogCount)
        {
            string worklogs = worklogCount == 1 ? "1 worklog" : $"{worklogCount.ToString(CultureInfo.InvariantCulture)} worklogs";
            string name = title.Trim().Length == 0 ? "(untitled)" : title.Trim();

            return $"#{workbookId.ToString(CultureInfo.InvariantCulture)} {name} ({status}, {dateText}, {worklogs})";
        }

        // ###########################################################################################
        // A device's photos and its links of interest, each under a headline of its own, counted as the worklog
        // editor counts its lists - "Photos (3)" (owner request, 2026-11-06: "Notes and photos" until then, "Notes"
        // a headline of its own since).
        // ###########################################################################################
        public static string PhotosHeading(int count) =>
            count == 0 ? "Photos" : $"Photos ({count.ToString(CultureInfo.InvariantCulture)})";

        public static string LinksHeading(int count) =>
            count == 0 ? "Links of interest" : $"Links of interest ({count.ToString(CultureInfo.InvariantCulture)})";

        // ###########################################################################################
        // Asked before a device is deleted: what goes (its photos, comments and links) and what stays
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

            return $"Delete [{name.Trim()}]? Its photos, comments and links are deleted with it. {repairs}";
        }
    }
}
