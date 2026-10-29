using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The device picker above Hardware and the worklog bar naming a repair's device (Main.Devices.cs -
// CRT 4.0 step 4, slice 4c; the plan: "A Machine picker sits beside the hardware and board
// drop-downs. Choosing a machine selects its board and its open repair. The bar reads 'Breadbin #2
// - No video'"). The picker shows the device the bar's workbook is on and nothing of its own, so
// the device named and the repair "Add worklog" writes into cannot disagree. Main's StartAsync is
// never called (MainWindowTests' header); the board shown is BoardKeyOverrideForTests.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainDevicesTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";
    private const string OtherBoardKey = "Commodore 64|250425";

    private readonly TempWorkspace thisWorkspace = new();

    public MainDevicesTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        UserSettings.EnableWorklog = true;
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static OwnedDeviceRecord AddDevice(string name, string boardKey = BoardKey) =>
        OwnedDeviceStore.Create(new OwnedDeviceDetails(name, boardKey, "PAL", string.Empty, string.Empty, OwnedDeviceStatus.InRepair))!;

    private static WorkbookRecord AddRepair(string title, int? deviceId, string boardKey = BoardKey) =>
        WorklogManager.CreateWorkbook(boardKey, title, string.Empty, deviceId)!;

    private static CRT.Main BuildMain()
    {
        var main = new CRT.Main { BoardKeyOverrideForTests = BoardKey };
        main.RefreshWorklogBar();
        return main;
    }

    // What the user does: a device picked in the drop-down, and what that started run to its end.
    private static async Task PickAsync(CRT.Main main, OwnedDeviceRecord device)
    {
        ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");
        picker.SelectedItem = picker.Items.OfType<OwnedDeviceRecord>().Single(each => each.Id == device.Id);

        Dispatcher.UIThread.RunJobs();
        await main.LastDevicePickForTests;
    }

    // Nobody without a device sees it, nor anybody with the worklog off - devices are the worklog's.
    [Fact]
    public void The_picker_is_shown_only_with_the_worklog_on_and_a_device()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            StackPanel panel = main.GetControl<StackPanel>("DevicePickerPanel");

            Assert.False(panel.IsVisible);

            AddDevice("Breadbin #2");
            main.RefreshWorklogBar();

            Assert.True(panel.IsVisible);
            Assert.Equal("Device", main.GetControl<TextBlock>("DevicePickerHeadingText").Text);

            UserSettings.EnableWorklog = false;
            main.ApplyWorklogBarVisibility();

            Assert.False(panel.IsVisible);
        });
    }

    // ###########################################################################################
    // It shows the device the bar's workbook is on - "No device" while that is on none, never the
    // device picked last.
    // ###########################################################################################
    [Fact]
    public void The_picker_shows_the_device_the_bars_workbook_is_on()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord onDevice = AddRepair("No video", breadbin.Id);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");

            // The newest is active - the one on no device.
            Assert.Null(main.DeviceWorkedOn);
            Assert.Equal("No device", picker.PlaceholderText);

            main.ActivateWorkbook(BoardKey, onDevice.Id);
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
            Assert.Equal(["Breadbin #2"], picker.Items.OfType<OwnedDeviceRecord>().Select(device => device.Name));

            main.ActivateWorkbook(BoardKey, onNone.Id);
            Assert.Null(main.DeviceWorkedOn);
        });
    }

    // ###########################################################################################
    // The bar's workbook picker: a repair on a device reads "Breadbin #2 - No video" - with its id
    // and board, as every workbook there, the list holding every board's.
    // ###########################################################################################
    [Fact]
    public void The_bar_names_the_device_a_repair_is_on()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord onDevice = AddRepair("No video", breadbin.Id);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox bar = main.GetControl<ComboBox>("WorklogJobBox");

            string TextOf(WorkbookRecord workbook) =>
                Assert.IsType<TextBlock>(bar.ItemTemplate!.Build(workbook)).Text ?? string.Empty;

            // No synced data in a test, so the board is its key as stored.
            Assert.Equal($"#{onDevice.Id} | Breadbin #2 - No video ({BoardKey})", TextOf(onDevice));
            Assert.Equal($"#{onNone.Id} | Dead keyboard ({BoardKey})", TextOf(onNone));
        });
    }

    // A workbook naming a device that has gone reads as on none.
    [Fact]
    public void A_repair_on_a_device_that_is_gone_is_named_as_on_none()
    {
        var workbook = new WorkbookRecord { Id = 4, Title = "No video", BoardKey = BoardKey, OwnedDeviceId = 99 };

        Assert.Equal($"#4 | No video ({BoardKey})", CRT.Main.WorkbookPickerText(workbook, null));
        Assert.Equal($"#4 | Breadbin #2 - No video ({BoardKey})", CRT.Main.WorkbookPickerText(workbook, "Breadbin #2"));
    }

    // ###########################################################################################
    // Picked: its open repair becomes the bar's workbook - the one "Add worklog" writes into - and
    // the picker shows the device, now that the workbook says it.
    // ###########################################################################################
    [Fact]
    public async Task Picking_a_device_works_on_its_open_repair()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord open = AddRepair("No video", breadbin.Id);
        AddRepair("Dead keyboard", null);

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            Assert.Null(main.DeviceWorkedOn);

            await PickAsync(main, breadbin);

            Assert.Equal(open.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
        });
    }

    // ###########################################################################################
    // A device with no repair yet: "New repair on <name>" on its board, and the repair started
    // becomes the bar's workbook. Called off, the picker goes back to the bar's workbook - nothing
    // is made.
    // ###########################################################################################
    [Fact]
    public async Task Picking_a_device_with_no_repair_starts_one_on_its_board()
    {
        OwnedDeviceRecord plus4 = AddDevice("Plus/4");
        AddRepair("Dead keyboard", null);
        var asked = new List<(string Device, string Board)>();

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            main.StartRepairOverrideForTests = (device, boardKey) =>
            {
                asked.Add((device.Name, boardKey));
                return Task.FromResult(WorklogManager.CreateWorkbook(boardKey, "Black screen", string.Empty, device.Id));
            };

            await PickAsync(main, plus4);

            Assert.Equal(("Plus/4", BoardKey), Assert.Single(asked));
            WorkbookRecord started = OwnedDeviceRepairs.Of(plus4.Id, WorklogManager.GetAllWorkbooks()).Single();
            Assert.Equal(started.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal(plus4.Id, main.DeviceWorkedOn!.Id);
        });
    }

    [Fact]
    public async Task A_new_repair_called_off_puts_the_picker_back()
    {
        OwnedDeviceRecord plus4 = AddDevice("Plus/4");
        AddRepair("Dead keyboard", null);

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            main.StartRepairOverrideForTests = (_, _) => Task.FromResult<WorkbookRecord?>(null);

            await PickAsync(main, plus4);

            Assert.Null(main.DeviceWorkedOn);
            Assert.Null(main.GetControl<ComboBox>("DevicePickerComboBox").SelectedItem);
            Assert.Single(WorklogManager.GetAllWorkbooks());
        });
    }

    // A device on a board CRT does not list has nothing to start a repair on: nothing is asked, and
    // the picker goes back. (A board no test's data can hold - DataManager's list is shared.)
    [Fact]
    public async Task A_device_on_a_board_CRT_does_not_list_starts_nothing()
    {
        OwnedDeviceRecord c128 = AddDevice("C128D", "Nowhere|" + Guid.NewGuid().ToString("N"));
        bool asked = false;

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            main.StartRepairOverrideForTests = (_, _) =>
            {
                asked = true;
                return Task.FromResult<WorkbookRecord?>(null);
            };

            await PickAsync(main, c128);

            Assert.False(asked);
            Assert.Null(main.DeviceWorkedOn);
        });
    }

    // ###########################################################################################
    // A device added in "Create new workbook" is on disk whether or not the workbook is then made:
    // the form called off, the picker shows it all the same (code review, 2026-10-29: hidden until
    // some other refresh came).
    // ###########################################################################################
    [Fact]
    public void A_device_added_in_a_called_off_Create_new_workbook_is_in_the_picker()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            Assert.False(main.GetControl<StackPanel>("DevicePickerPanel").IsVisible);

            AddDevice("Breadbin #2");
            main.AfterCreateWorkbookForm(created: null, deviceAdded: true);

            Assert.True(main.GetControl<StackPanel>("DevicePickerPanel").IsVisible);
            Assert.False(main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText").IsVisible);
        });
    }

    // ###########################################################################################
    // An open repair on a board CRT no longer lists cannot be shown: the device is worked on through
    // its repair on the board it has now instead (code review, 2026-10-29: the pick did nothing).
    // ###########################################################################################
    [Fact]
    public async Task A_device_whose_open_repair_is_on_a_board_gone_is_worked_on_on_its_board()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());
        Assert.False(WorklogManager.IsWorkbookStatusOpen(WorklogManager.GetAllWorkbooks().Single(each => each.Id == closed.Id).Status));
        AddRepair("No video", breadbin.Id, "Nowhere|" + Guid.NewGuid().ToString("N"));
        AddRepair("Dead keyboard", null);

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();

            await PickAsync(main, breadbin);

            Assert.Equal(closed.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
        });
    }

    // ###########################################################################################
    // The closed picker takes no pick from the arrow keys (code review, 2026-10-29): each step was a
    // board load, and a device with no repair opened "New repair" mid-scroll. The list is opened to
    // pick from.
    // ###########################################################################################
    [Fact]
    public void The_closed_picker_takes_no_pick_from_the_arrow_keys()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddDevice("C64C silver");
        AddRepair("No video", breadbin.Id);
        bool asked = false;

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.StartRepairOverrideForTests = (_, _) =>
            {
                asked = true;
                return Task.FromResult<WorkbookRecord?>(null);
            };

            ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");
            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Down, Source = picker });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("Breadbin #2", Assert.IsType<OwnedDeviceRecord>(picker.SelectedItem).Name);
            Assert.False(asked);
        });
    }

    // ###########################################################################################
    // A board unticked in the Configuration tab's list is not in the drop-downs, so CRT cannot show
    // it (code review, 2026-10-29: picking a device on one blanked the drop-downs). A board CRT
    // lists can be, and the board shown always can.
    // ###########################################################################################
    [Fact]
    public void A_board_hidden_from_the_lists_cannot_be_shown()
    {
        string hidden = "Commodore 128|310378 " + Guid.NewGuid().ToString("N");
        string listed = "Commodore 128|310379 " + Guid.NewGuid().ToString("N");

        try
        {
            ListBoards(hidden, listed);
            UserSettings.SetCatalogueKeyChecked(CatalogueVisibility.BuildKey(hidden.Split('|')[0], hidden.Split('|')[1]), false);

            UiTest.Run(() =>
            {
                CRT.Main main = BuildMain();

                Assert.False(main.CanShowBoard(hidden));
                Assert.True(main.CanShowBoard(listed));
                Assert.True(main.CanShowBoard(BoardKey));
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // A device with no repair whose board the drop-downs do not land on gets no "New repair" - it
    // went on the board still shown (code review, 2026-10-29). Here the drop-downs hold nothing, as
    // in every test, so they land nowhere.
    // ###########################################################################################
    [Fact]
    public async Task A_device_whose_board_is_not_landed_on_starts_no_repair_on_another()
    {
        string listed = "Commodore 128|310378 " + Guid.NewGuid().ToString("N");
        OwnedDeviceRecord c128 = AddDevice("C128D", listed);
        bool asked = false;

        try
        {
            ListBoards(listed);

            await UiTest.RunAsync(async () =>
            {
                CRT.Main main = BuildMain();
                main.StartRepairOverrideForTests = (_, _) =>
                {
                    asked = true;
                    return Task.FromResult<WorkbookRecord?>(null);
                };

                await PickAsync(main, c128);

                Assert.False(asked);
                Assert.Null(main.DeviceWorkedOn);
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // The device the picker already shows, chosen again from the open list: its open repair is
    // worked on - the plan's "Choosing a machine selects its board and its open repair" - when the
    // bar holds one of its closed repairs (code review, 2026-10-29: no selection change, no pick).
    // ###########################################################################################
    [Fact]
    public async Task Choosing_the_device_shown_again_works_on_its_open_repair()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord open = AddRepair("No video", breadbin.Id);
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());

        await UiTest.RunAsync(async () =>
        {
            // The bar holds the newest repair - the closed one - so the picker shows Breadbin #2.
            CRT.Main main = BuildMain();
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");
            picker.IsDropDownOpen = true;
            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter, Source = picker });
            picker.IsDropDownOpen = false;

            Dispatcher.UIThread.RunJobs();
            await main.LastDevicePickForTests;

            Assert.Equal(open.Id, UserSettings.GetActiveWorkbookId(BoardKey));
        });
    }

    // ###########################################################################################
    // The same, by a CLICK on the device's item in the open list - the popup clicked as a person
    // clicks it (code review, 2026-10-29: the list closed before the click was counted, so a click
    // did nothing; only Enter worked).
    // ###########################################################################################
    [Fact]
    public async Task Clicking_the_device_shown_again_works_on_its_open_repair()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord open = AddRepair("No video", breadbin.Id);
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            main.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");
                picker.IsDropDownOpen = true;
                Dispatcher.UIThread.RunJobs();

                ComboBoxItem item = picker.GetRealizedContainers().OfType<ComboBoxItem>()
                    .Single(container => container.Content is OwnedDeviceRecord device && device.Id == breadbin.Id);
                TopLevel popup = TopLevel.GetTopLevel(item)!;
                Point at = item.TranslatePoint(new Point(item.Bounds.Width / 2, item.Bounds.Height / 2), popup)!.Value;

                popup.MouseDown(at, MouseButton.Left);
                popup.MouseUp(at, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();
                await main.LastDevicePickForTests;

                Assert.False(picker.IsDropDownOpen);
                Assert.Equal(open.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // A pick whose board switch lands nowhere - the repair's board named in other capitals, which
    // the drop-downs take as the board already shown - puts the picker back on the bar's workbook
    // (code review, 2026-10-29: it stayed on the device, naming another repair than the one "Add
    // worklog" writes into).
    // ###########################################################################################
    [Fact]
    public async Task A_pick_whose_board_switch_lands_nowhere_puts_the_picker_back()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id, BoardKey.ToUpperInvariant());
        AddRepair("Dead keyboard", null);

        try
        {
            ListBoards(BoardKey);

            await UiTest.RunAsync(async () =>
            {
                CRT.Main main = BuildMain();

                await PickAsync(main, breadbin);

                Assert.Null(main.DeviceWorkedOn);
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // ONE pick for one choice: the list opened on Breadbin #2, a refresh moved the picker to the
    // other device's repair meanwhile, and Enter on Breadbin #2 both changes the selection back and
    // closes the list - it is worked on once, not twice (code review, 2026-10-29: a device with no
    // repair got "New repair" twice in a row).
    // ###########################################################################################
    [Fact]
    public async Task Enter_after_a_refresh_moved_the_picker_makes_one_pick()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        OwnedDeviceRecord silver = AddDevice("C64C silver");
        AddRepair("No video", breadbin.Id);
        WorkbookRecord silverRepair = AddRepair("Dead keyboard", silver.Id);
        AddRepair("PSU replaced", breadbin.Id);

        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = main.GetControl<ComboBox>("DevicePickerComboBox");
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            picker.IsDropDownOpen = true;

            UserSettings.SetActiveWorkbookId(BoardKey, silverRepair.Id);
            main.RefreshWorklogBar();
            Assert.Equal(silver.Id, main.DeviceWorkedOn!.Id);

            int picksBefore = main.DevicePicksForTests;
            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter, Source = picker });
            picker.SelectedItem = picker.Items.OfType<OwnedDeviceRecord>().Single(device => device.Id == breadbin.Id);
            picker.IsDropDownOpen = false;

            Dispatcher.UIThread.RunJobs();
            await main.LastDevicePickForTests;

            Assert.Equal(1, main.DevicePicksForTests - picksBefore);
        });
    }

    // Lists `boardKeys` as the boards CRT has (DataManager's main file) - put back with Unload.
    private void ListBoards(params string[] boardKeys) =>
        DataManager.LoadFrom(
            this.thisWorkspace.Path_("Data"),
            MainFiles.Write(
                this.thisWorkspace.Path_("Data"),
                boardKeys.Select(key => new MainFiles.Board("Test/" + Guid.NewGuid().ToString("N"), key.Split('|')[0], key.Split('|')[1])).ToArray()));

    // ###########################################################################################
    // "Create new workbook" in the bar preselects the device worked on - only on its own board.
    // ###########################################################################################
    [Fact]
    public void Create_new_workbook_in_the_bar_preselects_the_device_worked_on()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();

            Assert.Equal(breadbin.Id, main.DeviceForNewWorkbook(BoardKey));
            Assert.Null(main.DeviceForNewWorkbook(OtherBoardKey));

            AddRepair("Dead keyboard", null);
            main.RefreshWorklogBar();

            Assert.Null(main.DeviceForNewWorkbook(BoardKey));
        });
    }
}
