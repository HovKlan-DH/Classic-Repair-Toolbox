using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE DEVICE PICKER beside the hardware and board drop-downs (CRT 4.0 step 4, slice 4c - the
    // plan: "A Machine picker sits beside the hardware and board drop-downs. Choosing a machine
    // selects its board and its open repair. The bar reads 'Breadbin #2 - No video'").
    //
    // THE PICKER HAS NO STATE OF ITS OWN: it shows the device the workbook in the worklog bar is on,
    // "No device" while that is on none - so the device it names and the repair "Add worklog" writes
    // into can never disagree. Picking a device activates the repair it is worked on through
    // (OwnedDeviceRepairs.RepairToWorkOn), its board selected first when another is shown. A device
    // with no repair on its board yet gets its board selected and "New repair on <name>" asked; that
    // cancelled, the picker goes back to the bar's workbook.
    //
    // Shown while the worklog is on and the user has a device - nobody else sees it. Filled by
    // RefreshWorklogBar, the one place worklog state is drawn from, with the devices it read; the
    // bar's own workbook picker names each repair's device from the same read.
    //
    // Part of the Main partial class - see Main.axaml.cs for the file map.
    // ###########################################################################################
    public partial class Main
    {
        // Set while RefreshWorklogBar fills the picker, so its own selection is no pick.
        private bool thisFillingDevicePicker;

        // ###########################################################################################
        // What the last RefreshWorklogBar read - every workbook and every device - for a pick and the
        // forms it opens, never a second read of every folder (code review, 2026-10-29). Every worklog
        // and device change goes through that refresh, so they are what is on disk.
        // ###########################################################################################
        private List<WorkbookRecord> thisWorkbooksLastRead = [];

        private List<OwnedDeviceRecord> thisDevicesLastRead = [];

        // The device the open list was opened on, and whether an item was chosen in it (a click on an
        // item, or Enter) - so choosing the device already shown is a pick too, which no selection
        // change says.
        private int? thisPickerOpenedOn;

        private bool thisPickerItemChosen;

        // Stands in for "New repair on <name>" in a test, which cannot show a dialog - null in the app.
        internal Func<OwnedDeviceRecord, string, Task<WorkbookRecord?>>? StartRepairOverrideForTests { get; set; }

        // What the last pick started, for a test to wait on - and how many picks were started.
        internal Task LastDevicePickForTests { get; private set; } = Task.CompletedTask;

        internal int DevicePicksForTests { get; private set; }

        // The picker's items as a device's name - set once, in the constructor.
        private void WireDevicePicker()
        {
            this.DevicePickerHeadingText.Text = OwnedDeviceWording.PickerHeading;
            this.DevicePickerComboBox.PlaceholderText = OwnedDeviceWording.NoDevicePicked;
            this.DevicePickerComboBox.ItemTemplate = new FuncDataTemplate<OwnedDeviceRecord>(
                (device, _) => new TextBlock { Text = device?.Name ?? string.Empty });

            // A pick is made from the OPEN list (code review, 2026-10-29): closed, the arrow keys and
            // the wheel stepped the selection, every step a board load - and a device with no repair
            // opened "New repair" mid-scroll. Alt+Down, Enter and F4 still open the list.
            this.DevicePickerComboBox.AddHandler(KeyDownEvent, this.OnDevicePickerKeyDown, RoutingStrategies.Tunnel);
            this.DevicePickerComboBox.AddHandler(PointerWheelChangedEvent, this.OnDevicePickerWheel, RoutingStrategies.Tunnel);

            // The device shown, chosen again from the open list (code review, 2026-10-29). On the
            // TUNNEL route: the ComboBox closes the list in its own handler on the bubble route, so a
            // handler there saw it closed already and a click was never counted (second review).
            this.DevicePickerComboBox.AddHandler(PointerReleasedEvent, this.OnDevicePickerPointerReleased, RoutingStrategies.Tunnel, handledEventsToo: true);
            this.DevicePickerComboBox.PropertyChanged += this.OnDevicePickerPropertyChanged;
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
        // device chosen is SelectionChanged's.
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

            if (this.thisPickerItemChosen && this.DeviceWorkedOn is { } device && device.Id == this.thisPickerOpenedOn)
            {
                Dispatcher.UIThread.Post(() => this.LastDevicePickForTests = this.WorkOnDeviceAsync(device));
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
        // The picker as the worklog now is: every device, the one the bar's workbook is on picked.
        // `devices` is RefreshWorklogBar's read - empty while the worklog is off, which hides it.
        // ###########################################################################################
        private void FillDevicePicker(List<OwnedDeviceRecord> devices, WorkbookRecord? activeWorkbook)
        {
            this.thisDevicesLastRead = devices;

            this.DevicePickerPanel.IsVisible = UserSettings.EnableWorklog && devices.Count > 0;

            this.thisFillingDevicePicker = true;

            try
            {
                this.DevicePickerComboBox.ItemsSource = devices;
                this.DevicePickerComboBox.SelectedItem = OwnedDeviceRepairs.DeviceOf(activeWorkbook, devices);
            }
            finally
            {
                this.thisFillingDevicePicker = false;
            }
        }

        // The device a workbook's repair is on, by name - null for none - from the last refresh's devices.
        private string? DeviceNameOf(WorkbookRecord workbook) =>
            OwnedDeviceRepairs.DeviceOf(workbook, this.thisDevicesLastRead)?.Name;

        // ###########################################################################################
        // A workbook as the bar's picker lists it - "#3 | No video (C64/250407)" - and a repair on a
        // device names the device first, "#3 | Breadbin #2 - No video (C64/250407)" (the plan: "The
        // bar reads 'Breadbin #2 - No video'"). The board stays: the list holds every board's.
        // ###########################################################################################
        internal static string WorkbookPickerText(WorkbookRecord workbook, string? deviceName)
        {
            string title = deviceName is { Length: > 0 } name
                ? OwnedDeviceWording.BarText(name, workbook.Title)
                : workbook.Title;

            return $"#{workbook.Id} | {title} ({FormatBoardKeyForDisplay(workbook.BoardKey)})";
        }

        // The device the picker shows - the one the bar's workbook is on.
        internal OwnedDeviceRecord? DeviceWorkedOn => this.DevicePickerComboBox.SelectedItem as OwnedDeviceRecord;

        // ###########################################################################################
        // The device "Create new workbook" in the bar preselects: the one worked on, when it has the
        // board shown - a repair started from the bar is most likely another on it, and the form
        // shows it, to be changed there.
        // ###########################################################################################
        internal int? DeviceForNewWorkbook(string boardKey) =>
            this.DeviceWorkedOn is { } device && string.Equals(device.BoardKey.Trim(), boardKey, StringComparison.OrdinalIgnoreCase)
                ? device.Id
                : null;

        private void OnDevicePickerSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisFillingDevicePicker || this.DevicePickerComboBox.SelectedItem is not OwnedDeviceRecord device)
            {
                return;
            }

            // This choice is picked HERE - the list closing after it picks nothing more (code review,
            // 2026-10-29: a refresh had moved the picker while the list was open, and Enter on the
            // device it was opened on made two picks).
            this.thisPickerItemChosen = false;

            // After the drop-down has finished its own selection - what follows replaces its items.
            Dispatcher.UIThread.Post(() => this.LastDevicePickForTests = this.WorkOnDeviceAsync(device));
        }

        // ###########################################################################################
        // A device picked: the repair it is worked on through, activated on its board - or, with none,
        // its board and a new repair on it. Whatever comes of it, the bar and the picker are drawn
        // again, so a pick that leads nowhere puts the picker back.
        // ###########################################################################################
        internal async Task WorkOnDeviceAsync(OwnedDeviceRecord device)
        {
            this.DevicePicksForTests++;

            try
            {
                if (OwnedDeviceRepairs.RepairToWorkOn(device, this.thisWorkbooksLastRead, this.CanShowBoard) is { } repair)
                {
                    string before = this.GetCurrentBoardKey();
                    this.ActivateWorkbookAcrossBoards(repair.BoardKey, repair.Id);

                    // Activated on the board shown, or a load started that draws everything again - or
                    // the switch landed NOWHERE (the board named in other capitals): then the picker
                    // goes back on the bar's workbook, never left naming another repair than the one
                    // "Add worklog" writes into (code review, 2026-10-29).
                    if (!string.Equals(before, repair.BoardKey, StringComparison.Ordinal) &&
                        string.Equals(before, this.GetCurrentBoardKey(), StringComparison.Ordinal))
                    {
                        this.RefreshWorklogBar();
                    }

                    return;
                }

                string boardKey = this.ShowBoardOfDevice(device);

                if (boardKey.Length == 0)
                {
                    Logger.Warning($"Device [{device.Name}] is on a board the drop-downs do not show - [{device.BoardKey}]");
                    this.RefreshWorklogBar();
                    return;
                }

                if (await this.StartRepairOnDeviceAsync(device, boardKey) is { } created)
                {
                    this.ActivateWorkbook(created.BoardKey, created.Id);
                    return;
                }

                this.RefreshWorklogBar();
            }
            catch (Exception ex)
            {
                // Nothing may leave the posted pick: App lets such an exception end CRT.
                Logger.Critical($"Working on a device failed - [{ex}]");
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
        // The device's board shown - selected in the drop-downs when another is - and its key; empty
        // when the drop-downs do not list it, or the switch did not land on it, so there is nothing
        // to start a repair on (code review, 2026-10-29: the board still shown was handed back, and
        // the device's new repair went on it).
        // ###########################################################################################
        private string ShowBoardOfDevice(OwnedDeviceRecord device)
        {
            string wanted = device.BoardKey.Trim();
            string current = this.GetCurrentBoardKey();

            if (current.Length > 0 && string.Equals(current, wanted, StringComparison.OrdinalIgnoreCase))
            {
                return current;
            }

            if (FindListedEntryForBoardKey(wanted) is not { } entry)
            {
                return string.Empty;
            }

            // As a workbook picked for another board does: nothing drawn for the board being left.
            this.TabSchematicsControl.CancelWorklogEntryMode();
            this.SelectBoard(entry);

            string shown = this.GetCurrentBoardKey();
            return string.Equals(shown, wanted, StringComparison.OrdinalIgnoreCase) ? shown : string.Empty;
        }

        private async Task<WorkbookRecord?> StartRepairOnDeviceAsync(OwnedDeviceRecord device, string boardKey)
        {
            if (this.StartRepairOverrideForTests is { } start)
            {
                return await start(device, boardKey);
            }

            return await this.ShowCreateWorkbookAsync(this, boardKey, device.Id);
        }

        // ###########################################################################################
        // "Create new workbook" - opened HERE and nowhere else: the bar's button, "New repair" on a
        // device's page, and a device picked that has no repair yet. Whatever it closed with, a device
        // added in it is shown (AfterCreateWorkbookForm) - in one place, so no caller can forget it
        // (code reviews, 2026-10-29). The device preselected is named in the form's title. The
        // workbook made, which its caller activates, or null.
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
}
