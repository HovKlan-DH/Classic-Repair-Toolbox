using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE WORKLOG BAR'S TWO DROP-DOWNS: "Device", AND THE WORKBOOKS IT NARROWS TO (owner requests, 2026-11-06: "I
    // would like a device selector as the first drop-down, and then it will show all workbooks for this device in
    // the other 'Workbook' selector. Inside each of the drop-downs there should be a badge shown as the last in each
    // line"; then "rename 'No device' to 'Show workbooks for all devices' and then have that as first selection.
    // When this is selected (and it is per default if never set before) then it will show all workbooks", "show a
    // counter for number of workbooks", and a device with no workbook should "go to the 'My Devices' tab and select
    // that device" rather than open "Add workbook" - the device picker stood above Hardware until that day, CRT 4.0
    // step 4's slice 4c).
    //
    // THE DEVICE DROP-DOWN IS THE DEVICE PICKED, REMEMBERED (UserSettings.WorklogBarDeviceId) - ONE CHOICE WITH THE
    // DEVICE MARKED IN "My devices" on the "Workbooks" tab (owner request, 2026-11-06: a device selected there shows
    // "only the workbooks for this device", a workbook selected there "I want that device selected in both the
    // drop-down and in the 'My devices' panel", and its "Clear" shows all again): "Show workbooks for all devices"
    // first - every workbook, whatever device it is on - then every device, each line with how many workbooks it has
    // and its status. Greyed out with no device.
    //
    // WHILE A DEVICE IS PICKED CRT WORKS ONLY IN ITS WORKBOOKS (owner decision, 2026-11-06, asked: "Only the
    // device's"): the active workbook is resolved among them (OwnedDeviceRepairs.WorkedIn), so with none of them on
    // the board shown there is none - the "Workbook" drop-down says so and "Add worklog" is hidden - never another
    // device's written into unseen. A workbook activated that is not the device's takes the pick along
    // (OwnedDeviceRepairs.PickFollowing), and another board shown with none of its workbooks is all devices again
    // (OwnedDeviceRepairs.BarFilter).
    //
    // Picking a device - here or in "My devices" (PickDevice) - works on its repair (OwnedDeviceRepairs.RepairToWorkOn),
    // its board selected first when another is shown; one with no repair to work on has its own board shown and
    // its page opened on the "Workbooks" tab. "Show workbooks for all devices" picked changes no workbook.
    //
    // Each line of the workbook drop-down ends in Open or Closed (BuildStatePill), and one on another board than the
    // one shown names that board, as picking it switches there.
    //
    // Filled by RefreshWorklogBar, the one place worklog state is drawn from, with what it read.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // The workbook picker's words with "Show workbooks for all devices" picked and workbooks to list, none of them on the
        // board shown (owner request, 2026-11-07: "Select a workbook"). With no workbook at all there is no picker
        // (FillBarPickers; owner request, 2026-11-08 - "No workbook on this board yet" until then).
        internal const string SelectWorkbookText = "Select a workbook";

        // ... and with workbooks there but none worked in - "Clear" under the "Workbooks" tab's workbook list.
        internal const string NoWorkbookPickedText = "No workbook picked";

        // Set while RefreshWorklogBar fills the picker, so its own selection is no pick.
        private bool thisFillingDevicePicker;

        // ###########################################################################################
        // What the last RefreshWorklogBar read - every workbook and every device - for a pick and the
        // forms it opens, never a second read of every folder (code review, 2026-10-29). Every worklog
        // and device change goes through that refresh, so they are what is on disk.
        // ###########################################################################################
        private List<WorkbookRecord> thisWorkbooksLastRead = [];

        private List<OwnedDeviceRecord> thisDevicesLastRead = [];

        // The device the bar's workbook is on, as the last refresh found it - null for none.
        private OwnedDeviceRecord? thisBarWorkbookDevice;

        // The device the open list was opened on, and whether an item was chosen in it (a click on an
        // item, or Enter) - so choosing the device already shown is a pick too, which no selection
        // change says.
        private int? thisPickerOpenedOn;

        private bool thisPickerItemChosen;

        // How many picks were made - for a test.
        internal int DevicePicksForTests { get; private set; }

        // The two pickers' lines - set once, in the constructor.
        private void WireDevicePicker()
        {
            this.DevicePickerHeadingText.Text = OwnedDeviceWording.PickerHeading;
            this.DevicePickerComboBox.PlaceholderText = OwnedDeviceWording.AllDevicesChoice;
            // A line: the device - or all devices - then its status, then its count of workbooks as an outlined
            // badge, the last (owner request, 2026-11-06).
            this.DevicePickerComboBox.ItemTemplate = new FuncDataTemplate<BarDeviceChoice>(
                (choice, _) => choice == null
                    ? null
                    : BarPickerLine(
                        choice.Device?.Name ?? OwnedDeviceWording.AllDevicesChoice,
                        choice.Device is { } device ? WorklogInfoPillBuilder.BuildDeviceStatusPill(device.Status) : null,
                        WorklogInfoPillBuilder.BuildCountPill(choice.WorkbookCount)));

            // A workbook's line opens with its device's name in a badge while every device's workbooks are listed (owner
            // request, 2026-11-07: "so I can separate the workbooks in the list") - never with one device picked.
            this.WorklogJobBox.ItemTemplate = new FuncDataTemplate<WorkbookRecord>(
                (workbook, _) => workbook == null
                    ? null
                    : BarPickerLine(
                        this.DeviceBadgeOnBar(workbook),
                        WorkbookPickerText(workbook, this.GetCurrentBoardKey()),
                        WorklogInfoPillBuilder.BuildStatePill(workbook.Status)));

            // A pick is made from the OPEN list (code review, 2026-10-29): closed, the arrow keys and
            // the wheel stepped the selection, every step a board load. Alt+Down, Enter and F4 still
            // open the list.
            this.DevicePickerComboBox.AddHandler(KeyDownEvent, this.OnDevicePickerKeyDown, RoutingStrategies.Tunnel);
            this.DevicePickerComboBox.AddHandler(PointerWheelChangedEvent, this.OnDevicePickerWheel, RoutingStrategies.Tunnel);

            // The device shown, chosen again from the open list (code review, 2026-10-29). On the
            // TUNNEL route: the ComboBox closes the list in its own handler on the bubble route, so a
            // handler there saw it closed already and a click was never counted (second review).
            this.DevicePickerComboBox.AddHandler(PointerReleasedEvent, this.OnDevicePickerPointerReleased, RoutingStrategies.Tunnel, handledEventsToo: true);
            this.DevicePickerComboBox.PropertyChanged += this.OnDevicePickerPropertyChanged;
        }

        // ###########################################################################################
        // A line of either picker: its words, cut short with "..." when the bar is narrow, then its badges at
        // the line's end, in the order handed - the list's lines are as wide as the list (the BarPicker style),
        // so the badges stand one under the other. Too narrow for them, the words stay and the badges go
        // (BarPickerLinePanel). A null badge is none.
        // ###########################################################################################
        internal static BarPickerLinePanel BarPickerLine(string text, params Control?[] badges) =>
            BarPickerLine(null, text, badges);

        // A line opening with `leading` before its words (BarPickerLinePanel) - none when null.
        internal static BarPickerLinePanel BarPickerLine(Control? leading, string text, params Control?[] badges)
        {
            var line = new BarPickerLinePanel(hasLeading: leading is not null);

            if (leading is not null)
            {
                leading.Margin = new Thickness(0, 0, 6, 0);
                leading.VerticalAlignment = VerticalAlignment.Center;
                line.Children.Add(leading);
            }

            var words = new TextBlock
            {
                Text = text,
                FontSize = 12,
                TextTrimming = TextTrimming.CharacterEllipsis,
                VerticalAlignment = VerticalAlignment.Center,
            };

            words.Classes.Add("BarPickerText");
            line.Children.Add(words);

            int placed = 0;

            foreach (Control badge in badges.OfType<Control>())
            {
                badge.Margin = new Thickness(placed == 0 ? 8 : 4, 0, 0, 0);
                badge.VerticalAlignment = VerticalAlignment.Center;
                line.Children.Add(badge);
                placed++;
            }

            return line;
        }

        private void OnDevicePickerKeyDown(object? sender, KeyEventArgs e)
        {
            if (this.DevicePickerComboBox.IsDropDownOpen)
            {
                // Enter and Space both take the item the open list is on.
                this.thisPickerItemChosen |= e.Key is Key.Enter or Key.Space;
            }
            else if (!e.KeyModifiers.HasFlag(KeyModifiers.Alt) &&
                     e.Key is Key.Up or Key.Down or Key.Home or Key.End or Key.PageUp or Key.PageDown)
            {
                e.Handled = true;
            }
        }

        // A click on an item of the open list.
        private void OnDevicePickerPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            if (this.DevicePickerComboBox.IsDropDownOpen && e.Source is Visual source && source.FindAncestorOfType<ComboBoxItem>(includeSelf: true) is not null)
            {
                this.thisPickerItemChosen = true;
            }
        }

        // ###########################################################################################
        // The list opened: on which device, nothing chosen yet. Closed: the device it was opened on
        // chosen again is worked on - its open repair, as any pick (the plan: "Choosing a machine
        // selects its board and its open repair") - since no selection changed to say so. Another
        // choice is SelectionChanged's. "Show workbooks for all devices" chosen again changes nothing.
        // ###########################################################################################
        private void OnDevicePickerPropertyChanged(object? sender, AvaloniaPropertyChangedEventArgs e)
        {
            if (e.Property != ComboBox.IsDropDownOpenProperty)
            {
                return;
            }

            if (this.DevicePickerComboBox.IsDropDownOpen)
            {
                this.thisPickerOpenedOn = this.DeviceWorkedOn?.Id;
                this.thisPickerItemChosen = false;
                return;
            }

            if (this.thisPickerItemChosen &&
                this.DevicePickerComboBox.SelectedItem is BarDeviceChoice { Device: { } device } choice &&
                device.Id == this.thisPickerOpenedOn)
            {
                Dispatcher.UIThread.Post(() => this.WorkOnDevice(choice));
            }

            this.thisPickerItemChosen = false;
        }

        private void OnDevicePickerWheel(object? sender, PointerWheelEventArgs e)
        {
            if (!this.DevicePickerComboBox.IsDropDownOpen)
            {
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Both pickers as the worklog now is: the device picked (`filter`, RefreshWorklogBar's - checked and
        // saved there) and the workbooks it lists, the active one picked - none while the board shown has
        // none of the device's, the drop-down saying why. `devices` and `allWorkbooks` are RefreshWorklogBar's
        // read.
        //
        // SelectedItem is looked up BY ID inside the list handed over: WorkbookRecord has no Equals, so a
        // ComboBox shows an item selected only when it is the very object in its ItemsSource. Both are filled
        // under their guards - this is RefreshWorklogBar drawing the bar, not a user picking, and a pick would
        // call ActivateWorkbook -> RefreshWorklogBar again for a selection nobody made.
        // ###########################################################################################
        private void FillBarPickers(
            List<OwnedDeviceRecord> devices,
            List<WorkbookRecord> allWorkbooks,
            WorkbookRecord? activeWorkbook,
            OwnedDeviceRecord? filter,
            bool noneChosen)
        {
            this.thisBarWorkbookDevice = OwnedDeviceRepairs.DeviceOf(activeWorkbook, devices);

            var choices = new List<BarDeviceChoice> { new(null, allWorkbooks.Count) };
            // In the order "Sorting of devices" asks for, as "My devices" lists them (owner request, 2026-11-07).
            choices.AddRange(TabWorkbooks.SortDevices(devices).Select(device => new BarDeviceChoice(device, OwnedDeviceRepairs.Of(device.Id, allWorkbooks).Count)));

            List<WorkbookRecord> workbooks = filter is null ? allWorkbooks : OwnedDeviceRepairs.Of(filter.Id, allWorkbooks);

            // What the workbook lines' device badges read - set before the lines are made from them.
            this.thisBarListsEveryDevice = filter is null;
            this.thisBarDeviceNames = devices.GroupBy(device => device.Id).ToDictionary(group => group.Key, group => group.First().Name);

            this.thisFillingDevicePicker = true;
            this._suppressWorklogJobBoxSave = true;

            try
            {
                this.DevicePickerComboBox.ItemsSource = choices;
                this.DevicePickerComboBox.SelectedItem = choices.First(choice => choice.Device?.Id == filter?.Id);

                // *** A DROP-DOWN WITH NOTHING EVER MADE FOR IT IS NOT THERE (owner request, 2026-11-08: "if no devices or
                // workbooks exists yet, because the user has not created any yet, then its drop-down boxes should not show
                // either"). *** Its label goes with it, and the other has the room (BarPickersPanel lays out what is shown).
                // A workbook on any board counts: the drop-down lists every board's.
                bool anyDevice = devices.Count > 0;
                bool anyWorkbook = allWorkbooks.Count > 0;
                this.DevicePickerHeadingText.IsVisible = anyDevice;
                this.DevicePickerComboBox.IsVisible = anyDevice;
                this.WorkbookPickerHeadingText.IsVisible = anyWorkbook;
                this.WorkbookPickerHeadingText.Margin = new Thickness(anyDevice ? 16 : 0, 0, 8, 0);

                this.WorklogJobBox.IsVisible = anyWorkbook;
                this.WorklogJobBox.ItemsSource = workbooks;
                this.WorklogJobBox.SelectedItem = activeWorkbook == null ? null : workbooks.FirstOrDefault(workbook => workbook.Id == activeWorkbook.Id);
                this.WorklogJobBox.IsEnabled = workbooks.Count > 0;
                this.WorklogJobBox.PlaceholderText = activeWorkbook != null
                    ? null
                    : noneChosen
                        ? NoWorkbookPickedText
                        : filter is null
                            ? SelectWorkbookText
                            : workbooks.Count == 0 ? OwnedDeviceWording.NoWorkbookOfDeviceYet : OwnedDeviceWording.NoWorkbookOfDeviceOnBoard;
            }
            finally
            {
                this.thisFillingDevicePicker = false;
                this._suppressWorklogJobBoxSave = false;
            }
        }

        // ###########################################################################################
        // The device badge a workbook's line opens with: the name of the device it is on, while "Show workbooks for all
        // devices" is picked - none with a device picked, where every line is that device's, nor for a workbook on no
        // device (or on one gone).
        // ###########################################################################################
        private bool thisBarListsEveryDevice = true;

        private Dictionary<int, string> thisBarDeviceNames = new();

        private Border? DeviceBadgeOnBar(WorkbookRecord workbook) =>
            this.thisBarListsEveryDevice &&
            workbook.OwnedDeviceId is { } id &&
            this.thisBarDeviceNames.TryGetValue(id, out string? name) &&
            name.Trim().Length > 0
                ? WorklogInfoPillBuilder.BuildDeviceNamePill(name)
                : null;

        // ###########################################################################################
        // A workbook as the bar's picker lists it - "#3 | No video" - and one on another board than the one
        // shown with that board, "#3 | No video (C64/250407)": picking it switches the drop-downs there. Its
        // status is the badge at the line's end.
        // ###########################################################################################
        internal static string WorkbookPickerText(WorkbookRecord workbook, string boardShown)
        {
            string text = $"#{workbook.Id} | {workbook.Title}";

            return string.Equals(workbook.BoardKey?.Trim(), boardShown?.Trim(), StringComparison.OrdinalIgnoreCase)
                ? text
                : $"{text} ({FormatBoardKeyForDisplay(workbook.BoardKey ?? string.Empty)})";
        }

        // The device the "Device" drop-down narrows the workbooks to - null for all devices.
        internal OwnedDeviceRecord? DeviceWorkedOn => (this.DevicePickerComboBox.SelectedItem as BarDeviceChoice)?.Device;

        // ###########################################################################################
        // The device "Add workbook" in the bar preselects: the one picked, else the one the bar's workbook is
        // on - when it has the board shown. A repair started from the bar is most likely another on it, and
        // the form shows it, to be changed there.
        // ###########################################################################################
        internal int? DeviceForNewWorkbook(string boardKey) =>
            (this.DeviceWorkedOn ?? this.thisBarWorkbookDevice) is { } device && string.Equals(device.BoardKey.Trim(), boardKey, StringComparison.OrdinalIgnoreCase)
                ? device.Id
                : null;

        private void OnDevicePickerSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisFillingDevicePicker || this.DevicePickerComboBox.SelectedItem is not BarDeviceChoice choice)
            {
                return;
            }

            // This choice is picked HERE - the list closing after it picks nothing more (code review,
            // 2026-10-29: a refresh had moved the picker while the list was open, and Enter on the
            // device it was opened on made two picks).
            this.thisPickerItemChosen = false;

            // After the drop-down has finished its own selection - what follows replaces its items.
            Dispatcher.UIThread.Post(() => this.WorkOnDevice(choice));
        }

        // ###########################################################################################
        // A choice picked in the drop-down. All devices: remembered, every workbook listed, the bar's workbook
        // as it was. A device: picked (PickDevice) - its repair worked on, or with none its page on the "My
        // Devices" tab (owner request, 2026-11-06: "it should not open the 'create new workbook' modal, but
        // instead it should go to the 'My Devices' tab and select that device").
        // ###########################################################################################
        internal void WorkOnDevice(BarDeviceChoice choice)
        {
            this.DevicePicksForTests++;

            try
            {
                if (choice.Device is not { } device)
                {
                    this.ClearDevicePick();
                    return;
                }

                this.PickDevice(device, showPage: false);
            }
            catch (Exception ex)
            {
                // Nothing may leave the posted pick: App lets such an exception end CRT.
                Logger.Critical($"Working on a device failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // A DEVICE PICKED - in the bar's drop-down, in "My devices" (`showPage`: its page opened whatever it
        // has), or added: remembered, and worked on as the plan has it ("Choosing a machine selects its board
        // and its open repair") - the repair OwnedDeviceRepairs.RepairToWorkOn names activated on its board.
        // With none to work on, its own board is shown (a board CRT lists), where no workbook is active - CRT
        // works only in the device's (owner decision, 2026-11-06) - and its page opened on the "Workbooks" tab.
        // ###########################################################################################
        internal void PickDevice(OwnedDeviceRecord device, bool showPage)
        {
            UserSettings.WorklogBarDeviceId = device.Id;

            WorkbookRecord? repair = OwnedDeviceRepairs.RepairToWorkOn(device, this.thisWorkbooksLastRead, this.CanShowBoard);

            if (repair is not null)
            {
                this.WorkOnRepairPicked(repair);
            }
            else if (string.Equals(device.BoardKey.Trim(), this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase) ||
                     FindListedEntryForBoardKey(device.BoardKey) is null ||
                     !this.SwitchToBoard(device.BoardKey))
            {
                // On its board already, or one CRT cannot show: drawn again here. Else its board's load draws
                // everything again.
                this.RefreshWorklogBar();
            }

            if (showPage || repair is null)
            {
                this.ShowDevicePage(device.Id);
            }
        }

        // ###########################################################################################
        // No device picked - "Show workbooks for all devices" in the drop-down, or "Clear" under "My devices"
        // (owner request, 2026-11-06: "clicking that will unselect the device and by then show all again"):
        // every workbook listed, the board's own active one worked in.
        // ###########################################################################################
        internal void ClearDevicePick()
        {
            UserSettings.WorklogBarDeviceId = null;
            this.RefreshWorklogBar();
        }

        // ###########################################################################################
        // NO WORKBOOK WORKED IN on the board shown - "Clear" under the "Workbooks" tab's workbook list (owner decision,
        // 2026-11-06, asked: "No workbook worked in"; WorklogManager.NoWorkbookChosen): the bar's "Workbook" drop-down
        // empty and "Add worklog" hidden till a workbook is picked, activated or made. Remembered per board; an area
        // being marked for the workbook let go of is called off, as activating another calls it off.
        // ###########################################################################################
        internal void ClearWorkbookPick()
        {
            string boardKey = this.GetCurrentBoardKey();

            if (string.IsNullOrWhiteSpace(boardKey))
            {
                return;
            }

            this.TabSchematicsControl.CancelWorklogEntryMode();
            UserSettings.SetActiveWorkbookId(boardKey, WorklogManager.NoWorkbookChosen);
            this.RefreshWorklogBar();
        }

        // The "Workbooks" tab shown on a device's page.
        private void ShowDevicePage(int deviceId)
        {
            this.TabWorkbooks.ShowDevicePage(deviceId);

            if (this.WorkbooksTabItem.IsVisible)
            {
                this.MainTabControl.SelectedItem = this.WorkbooksTabItem;
            }
        }

        // ###########################################################################################
        // The repair a pick found, activated on its board. Activated on the board shown, or a load started
        // that draws everything again - or the switch landed NOWHERE (the board named in other capitals):
        // then the bar is drawn again on the workbook it works in, never left naming another repair than the
        // one "Add worklog" writes into (code review, 2026-10-29).
        // ###########################################################################################
        private void WorkOnRepairPicked(WorkbookRecord repair)
        {
            string before = this.GetCurrentBoardKey();
            this.ActivateWorkbookAcrossBoards(repair.BoardKey, repair.Id);

            if (!string.Equals(before, repair.BoardKey, StringComparison.Ordinal) &&
                string.Equals(before, this.GetCurrentBoardKey(), StringComparison.Ordinal))
            {
                this.RefreshWorklogBar();
            }
        }

        // Whether CRT can show the board `boardKey` names: the one shown, or one the drop-downs list -
        // a board unticked in the Configuration tab's list is not (code review, 2026-10-29).
        // A blank key names no board - not even with no board shown (code review, 2026-10-29).
        internal bool CanShowBoard(string boardKey) =>
            !string.IsNullOrWhiteSpace(boardKey) &&
            (string.Equals(boardKey, this.GetCurrentBoardKey(), StringComparison.OrdinalIgnoreCase) ||
             FindListedEntryForBoardKey(boardKey) is not null);

        // ###########################################################################################
        // "Add workbook" - opened HERE and nowhere else: the bar's button and "New repair" on a device's
        // page. Whatever it closed with, a device added in it is shown (AfterCreateWorkbookForm) - in one
        // place, so no caller can forget it (code reviews, 2026-10-29). The device preselected is named in
        // the form's title. The workbook made, which its caller activates, or null.
        // ###########################################################################################
        internal async Task<WorkbookRecord?> ShowCreateWorkbookAsync(Window owner, string boardKey, int? deviceId)
        {
            var dialog = new CreateWorkbookWindow();
            dialog.Initialize(boardKey, deviceId, this.thisDevicesLastRead);

            WorkbookRecord? created = await dialog.ShowDialog<WorkbookRecord?>(owner);
            this.AfterCreateWorkbookForm(created, dialog.DeviceAdded);
            return created;
        }

        // A device added in the form is on disk whether or not a workbook was made: called off, it is
        // shown now - a workbook made is activated by its caller, which shows everything again.
        internal void AfterCreateWorkbookForm(WorkbookRecord? created, bool deviceAdded)
        {
            if (created is null && deviceAdded)
            {
                this.RefreshWorklogBar();
            }
        }
    }

    // ###########################################################################################
    // A line of the worklog bar's "Device" drop-down: a device, or - Device null - "Show workbooks for all devices",
    // and how many workbooks it lists (owner request, 2026-11-06: "show a counter for number of workbooks").
    // ###########################################################################################
    internal sealed record BarDeviceChoice(OwnedDeviceRecord? Device, int WorkbookCount);
}
