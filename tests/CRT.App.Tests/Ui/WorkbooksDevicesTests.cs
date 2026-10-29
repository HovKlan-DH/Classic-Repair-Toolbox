using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "My devices" on the Workbooks tab (CRT 4.0 step 4, TabWorkbooks.Devices.cs): the user's own
// computers above the workbooks, a device's page - its repairs with their totals, the changes made
// to it, its notes and photos - in place of the workbook, and "On my device" under a workbook's
// title. Also the two forms it opens: "Add device" and "Create new workbook" with its device.
//
// HeadlessUi, as WorkbooksListTests is (see its collection note): it builds controls, and points
// WorklogManager - the Workbooks folder devices share - at a folder of its own per test.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class WorkbooksDevicesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    // Unique per test, so the active workbook a test saves cannot leak into the next.
    private readonly string thisBoardKey = "Commodore 64|250407 " + Guid.NewGuid().ToString("N");
    private readonly string thisOtherBoardKey = "Commodore 64|250425 " + Guid.NewGuid().ToString("N");
    private readonly List<(string BoardKey, int WorkbookId)> thisActivated = new();

    public WorkbooksDevicesTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N")));
    }

    public void Dispose()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Detached-" + Guid.NewGuid().ToString("N")));
        this.thisWorkspace.Dispose();
    }

    private TabWorkbooks BuildTab()
    {
        var tab = new TabWorkbooks { BoardKeyOverrideForTests = this.thisBoardKey };

        tab.ActivateWorkbookOverrideForTests = (boardKey, workbookId) =>
        {
            this.thisActivated.Add((boardKey, workbookId));
            UserSettings.SetActiveWorkbookId(boardKey, workbookId);
            tab.RefreshWorkbooks();
        };

        tab.RefreshWorkbooks();
        return tab;
    }

    private OwnedDeviceRecord AddDevice(string name = "Breadbin #2", string? boardKey = null) =>
        OwnedDeviceStore.Create(new OwnedDeviceDetails(name, boardKey ?? this.thisBoardKey, "PAL", "CA123456", "Flea market", ""))!;

    private static List<string> Texts(Control control) =>
        control.GetSelfAndVisualDescendants().OfType<TextBlock>().Select(text => text.Text ?? string.Empty).ToList();

    private static string Text(TabWorkbooks tab, string name) => tab.GetControl<TextBlock>(name).Text ?? string.Empty;

    private string Picture(string name)
    {
        string path = Path.Combine(this.thisWorkspace.Root, "Picked", name);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, [1, 2, 3]);
        return path;
    }

    [Fact]
    public void My_devices_says_what_it_is_for_until_there_is_one_then_lists_every_device()
    {
        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();

            Assert.True(tab.GetControl<TextBlock>("NoDevicesText").IsVisible);
            Assert.Equal(OwnedDeviceWording.NoDevicesYet, Text(tab, "NoDevicesText"));
            Assert.Contains("Accept", tab.GetControl<Button>("AddDeviceButton").Classes);

            this.AddDevice("Breadbin #2");
            this.AddDevice("Plus/4", "Commodore Plus/4|PCB");
            tab.RefreshWorkbooks();

            Assert.False(tab.GetControl<TextBlock>("NoDevicesText").IsVisible);
            StackPanel list = tab.GetControl<StackPanel>("DeviceListPanel");
            Assert.Equal(2, list.Children.Count);
            Assert.Contains("Breadbin #2", Texts((Control)list.Children[0]));
            Assert.Contains("Working", Texts((Control)list.Children[0]));
        });
    }

    // ###########################################################################################
    // A device clicked: its page in place of the workbook - what it is, and its repairs on ANY
    // board (one made before its board was swapped says which), totalled, with none of the repairs
    // that are on no device or on another device.
    // ###########################################################################################
    [Fact]
    public void A_device_clicked_shows_its_page_with_its_repairs_on_any_board()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceRecord other = this.AddDevice("C64C silver");
        WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "", device.Id);
        WorklogManager.CreateWorkbook(this.thisOtherBoardKey, "PSU replaced", "", device.Id);
        WorklogManager.CreateWorkbook(this.thisBoardKey, "Someone else's repair", "", other.Id);
        WorklogManager.CreateWorkbook(this.thisBoardKey, "On no device", "");

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();

            tab.SelectDevice(device.Id);

            Assert.True(tab.IsDevicePageShown);
            Assert.False(tab.GetControl<Grid>("WorkbookViewGrid").IsVisible);
            Assert.Equal("Breadbin #2", Text(tab, "DeviceNameText"));
            Assert.Contains("Working", Texts(tab.GetControl<Decorator>("DeviceStatusHost")));
            Assert.EndsWith(" - PAL - serial CA123456", Text(tab, "DeviceDetailText"), StringComparison.Ordinal);
            Assert.Equal("Where it came from: Flea market", Text(tab, "DeviceOriginText"));

            Assert.StartsWith("2 repairs", tab.GetControl<TextBlock>("DeviceRepairsSummaryText").Inlines!.Text, StringComparison.Ordinal);
            StackPanel repairs = tab.GetControl<StackPanel>("DeviceRepairsPanel");
            Assert.Equal(["PSU replaced", "No video"], repairs.Children.Select(row => ((WorkbookRecord)((Control)row).Tag!).Title));
            Assert.Contains(Texts((Control)repairs.Children[0]), text => text.StartsWith("On Commodore 64", StringComparison.Ordinal));
            Assert.DoesNotContain(Texts((Control)repairs.Children[1]), text => text.StartsWith("On Commodore 64", StringComparison.Ordinal));

            // The device's card is the one marked - no workbook card beside it.
            Assert.Contains("Selected", ((Control)tab.GetControl<StackPanel>("DeviceListPanel").Children[0]).Classes);
            Assert.DoesNotContain(tab.GetControl<StackPanel>("WorkbookListPanel").Children, card => card.Classes.Contains("Selected"));
        });
    }

    // ###########################################################################################
    // A device's status pill is the tab's ONE informational pill look (code review, 2026-10-28: the
    // tab drew its own) - the workbook state pill's border, rounding and padding, in its own colour.
    // ###########################################################################################
    [Fact]
    public void A_devices_status_pill_is_the_informational_pill()
    {
        UiTest.Run(() =>
        {
            Border device = WorklogInfoPillBuilder.BuildDeviceStatusPill("Faulty", fontSize: 10);
            Border state = WorklogInfoPillBuilder.BuildStatePill("Open", fontSize: 10);

            Assert.Equal((state.BorderThickness, state.CornerRadius, state.Padding), (device.BorderThickness, device.CornerRadius, device.Padding));
            Assert.Equal("Faulty", ((TextBlock)device.Child!).Text);
        });
    }

    // A device's page takes its repairs from every workbook the refresh was handed - Main has read
    // them already - never a second read of the Workbooks folder (code review, 2026-10-28).
    [Fact]
    public void A_devices_page_takes_its_repairs_from_the_workbooks_the_refresh_was_handed()
    {
        OwnedDeviceRecord device = this.AddDevice();

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            var handed = new WorkbookRecord { Id = 99, BoardKey = this.thisOtherBoardKey, Title = "Handed over", OwnedDeviceId = device.Id };
            tab.RefreshWorkbooks([], [handed]);

            Assert.Equal(["Handed over"], tab.GetControl<StackPanel>("DeviceRepairsPanel").Children.Select(row => ((WorkbookRecord)((Control)row).Tag!).Title));
        });
    }

    // A repair opened from the page: the workbook back in its place, activated on its own board.
    [Fact]
    public void Opening_a_repair_brings_the_workbook_back_and_activates_it()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord repair = WorklogManager.CreateWorkbook(this.thisOtherBoardKey, "PSU replaced", "", device.Id)!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            tab.OpenRepair(repair);

            Assert.False(tab.IsDevicePageShown);
            Assert.True(tab.GetControl<Grid>("WorkbookViewGrid").IsVisible);
            Assert.Equal([(this.thisOtherBoardKey, repair.Id)], this.thisActivated);
        });
    }

    // A workbook card clicked brings the workbook back too.
    [Fact]
    public void A_workbook_clicked_brings_the_workbook_back()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "")!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            tab.SelectWorkbookForTests(workbook.Id);

            Assert.False(tab.IsDevicePageShown);
            Assert.Null(tab.SelectedDeviceIdForTests);
            Assert.Contains("Selected", ((Control)Assert.Single(tab.GetControl<StackPanel>("WorkbookListPanel").Children)).Classes);
            Assert.DoesNotContain("Selected", ((Control)tab.GetControl<StackPanel>("DeviceListPanel").Children[0]).Classes);
        });
    }

    [Fact]
    public void A_change_typed_is_added_with_its_date_and_the_line_emptied()
    {
        OwnedDeviceRecord device = this.AddDevice();

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);
            Assert.True(tab.GetControl<TextBlock>("NoDeviceChangesText").IsVisible);

            tab.GetControl<CalendarDatePicker>("ChangeDatePicker").SelectedDate = new DateTime(2025, 6, 14);
            tab.GetControl<TextBox>("ChangeTextBox").Text = "JiffyDOS ROM fitted";
            tab.AddDeviceChange();

            OwnedDeviceChangeRecord change = Assert.Single(OwnedDeviceStore.Get(device.Id)!.Changes);
            Assert.Equal((new DateTime(2025, 6, 14), "JiffyDOS ROM fitted"), (change.Date, change.Text));
            Assert.Equal(string.Empty, tab.GetControl<TextBox>("ChangeTextBox").Text);
            Assert.Equal("Changes made to it (1)", Text(tab, "DeviceChangesHeadingText"));

            Control row = (Control)Assert.Single(tab.GetControl<StackPanel>("DeviceChangesPanel").Children);
            Assert.Contains("2025-06-14", Texts(row));

            // Its Remove takes it off.
            Button remove = row.GetSelfAndVisualDescendants().OfType<Button>().Single();
            remove.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.Empty(OwnedDeviceStore.Get(device.Id)!.Changes);
        });
    }

    [Fact]
    public void Notes_are_saved_when_the_box_is_left()
    {
        OwnedDeviceRecord device = this.AddDevice();

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            tab.GetControl<TextBox>("DeviceNotesTextBox").Text = "Recapped in 2024.";
            tab.SaveDeviceNotes();

            Assert.Equal("Recapped in 2024.", OwnedDeviceStore.Get(device.Id)!.Notes);
        });
    }

    // ###########################################################################################
    // Notes typed for one device and not saved yet are saved to THAT device when another is picked
    // (code review, 2026-10-28): a card click takes no focus from the box, so the next device's page
    // showed the first one's notes and leaving the box saved them over the second device's.
    // ###########################################################################################
    [Fact]
    public void Notes_typed_for_one_device_are_saved_to_it_when_another_is_picked()
    {
        OwnedDeviceRecord first = this.AddDevice("Breadbin #2");
        OwnedDeviceRecord second = this.AddDevice("C64C silver");
        OwnedDeviceStore.SetNotes(second.Id, "Its own notes");

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            var window = new Window { Content = tab, Width = 1200, Height = 900 };
            window.Show();

            tab.SelectDevice(first.Id);
            Dispatcher.UIThread.RunJobs();
            TextBox notes = tab.GetControl<TextBox>("DeviceNotesTextBox");
            Assert.True(notes.Focus());
            Assert.True(notes.IsFocused);
            notes.Text = "Recapped in 2024.";

            tab.SelectDevice(second.Id);

            Assert.Equal("Its own notes", notes.Text);
            Assert.Equal("Recapped in 2024.", OwnedDeviceStore.Get(first.Id)!.Notes);

            // The box left on the second device's page: its own notes stay its own.
            tab.SaveDeviceNotes();
            Assert.Equal("Its own notes", OwnedDeviceStore.Get(second.Id)!.Notes);

            window.Close();
        });
    }

    // Photos are copied in under CRT's please-wait, off the UI thread (code review, 2026-10-28).
    [Fact]
    public async Task Photos_added_show_on_the_page_and_Remove_takes_one_off()
    {
        OwnedDeviceRecord device = this.AddDevice();
        string photo = this.Picture("front.jpg");

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            var overlay = new BusyOverlay();
            var root = new Grid();
            root.Children.Add(tab);
            root.Children.Add(overlay);
            var window = new Window { Content = root, Width = 1200, Height = 900 };
            window.Show();
            tab.SelectDevice(device.Id);

            Task adding = tab.AddDevicePhotosAsync([photo]);
            Assert.True(overlay.IsBusy);
            await adding;

            // Decoded once, however often the tab is drawn again (code review, 2026-10-28).
            tab.RefreshWorkbooks();
            tab.RefreshWorkbooks();
            Assert.Equal(1, tab.DevicePhotosDecodedForTests);

            // What was added is said - and one that could not be, gone since it was picked, too.
            Assert.Equal(OwnedDeviceWording.PhotosAdded(1, 0), Text(tab, "DevicePhotosNoteText"));
            await tab.AddDevicePhotosAsync([Path.Combine(this.thisWorkspace.Root, "Picked", "gone.jpg")]);
            Assert.Equal(OwnedDeviceWording.PhotosAdded(0, 1), Text(tab, "DevicePhotosNoteText"));

            WrapPanel photos = tab.GetControl<WrapPanel>("DevicePhotosPanel");
            Control tile = (Control)Assert.Single(photos.Children);
            Assert.Equal("Notes and photos (1 photo)", Text(tab, "DevicePhotosHeadingText"));

            tile.GetSelfAndVisualDescendants().OfType<Button>().Single().RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            Assert.Empty(OwnedDeviceStore.Get(device.Id)!.Photos);
            Assert.Empty(photos.Children);

            window.Close();
        });
    }

    // ###########################################################################################
    // Delete device asks first, naming how many repairs stay; "no" keeps it, "yes" deletes it and
    // brings the workbook back - its repair still there, on no device.
    // ###########################################################################################
    [Fact]
    public async Task Delete_device_asks_first_and_leaves_its_repairs()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord repair = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "", device.Id)!;

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            var asked = new List<int>();
            bool answer = false;
            tab.ConfirmDeleteDeviceOverrideForTests = (_, repairs) =>
            {
                asked.Add(repairs);
                return Task.FromResult(answer);
            };

            await tab.DeleteSelectedDeviceAsync();
            Assert.NotNull(OwnedDeviceStore.Get(device.Id));

            answer = true;
            await tab.DeleteSelectedDeviceAsync();

            Assert.Equal([1, 1], asked);
            Assert.Null(OwnedDeviceStore.Get(device.Id));
            Assert.False(tab.IsDevicePageShown);
            Assert.Equal(repair.Id, Assert.Single(WorklogManager.GetAllWorkbooks()).Id);
        });
    }

    // ###########################################################################################
    // "On my device" under a workbook's title: offered while a device has the workbook's board -
    // None and those devices - and the repair put on one names it on its card.
    // ###########################################################################################
    [Fact]
    public void A_workbook_is_put_on_a_device_with_its_board_from_under_its_title()
    {
        OwnedDeviceRecord device = this.AddDevice();
        this.AddDevice("C64C silver", this.thisOtherBoardKey);
        WorkbookRecord workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "")!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();

            Assert.True(tab.GetControl<StackPanel>("WorkbookDevicePanel").IsVisible);
            ComboBox choices = tab.GetControl<ComboBox>("WorkbookDeviceComboBox");
            Assert.Equal(["None", "Breadbin #2"], choices.Items.Select(item => item!.ToString()));
            Assert.Equal("None", choices.SelectedItem!.ToString());

            tab.PutShownWorkbookOnDevice(device.Id);

            Assert.Equal(device.Id, WorklogManager.GetAllWorkbooks().Single().OwnedDeviceId);
            Assert.Equal("Breadbin #2", tab.GetControl<ComboBox>("WorkbookDeviceComboBox").SelectedItem!.ToString());
            Assert.Contains("On Breadbin #2", Texts((Control)tab.GetControl<StackPanel>("WorkbookListPanel").Children[0]));

            // Picked in the drop-down: written once the drop-down has finished.
            choices.SelectedItem = choices.Items.Cast<object>().First(item => item.ToString() == "None");
            Dispatcher.UIThread.RunJobs();
            Assert.Null(WorklogManager.GetAllWorkbooks().Single().OwnedDeviceId);
        });
    }

    [Fact]
    public void With_no_device_on_its_board_nothing_is_offered_under_a_workbook()
    {
        this.AddDevice("C64C silver", this.thisOtherBoardKey);
        WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "");

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();

            Assert.False(tab.GetControl<StackPanel>("WorkbookDevicePanel").IsVisible);
        });
    }

    // ###########################################################################################
    // "Create new workbook" asks which device the repair is on - the plan: "one of yours on this
    // board, a new one, or none" - with "New repair" on a device's page preselecting it and naming
    // it in the title.
    // ###########################################################################################
    [Fact]
    public void Create_new_workbook_puts_the_repair_on_the_device_picked()
    {
        OwnedDeviceRecord device = this.AddDevice();
        this.AddDevice("C64C silver", this.thisOtherBoardKey);

        UiTest.Run(() =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey, device.Id);

            Assert.True(window.GetControl<StackPanel>("DevicePanel").IsVisible);
            Assert.Equal(["None", "Breadbin #2"], window.GetControl<ComboBox>("DeviceComboBox").Items.Select(item => item!.ToString()));
            Assert.Equal("Breadbin #2", window.GetControl<ComboBox>("DeviceComboBox").SelectedItem!.ToString());
            Assert.Equal("New repair on Breadbin #2", window.Title);

            window.CreateForTests("No video");

            Assert.Equal(device.Id, window.Created!.OwnedDeviceId);
        });
    }

    // With no device on this board, "None" and "Add device" are still offered - starting a repair is
    // where a new device comes up. The form's own title stays its own.
    [Fact]
    public void Create_new_workbook_offers_a_new_device_with_none_on_the_board()
    {
        this.AddDevice("C64C silver", this.thisOtherBoardKey);

        UiTest.Run(() =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey);

            Assert.True(window.GetControl<StackPanel>("DevicePanel").IsVisible);
            Assert.Equal(["None"], window.GetControl<ComboBox>("DeviceComboBox").Items.Select(item => item!.ToString()));
            Assert.Equal("None", window.GetControl<ComboBox>("DeviceComboBox").SelectedItem!.ToString());
            Assert.Equal("Add device", window.GetControl<Button>("AddDeviceButton").Content);
            Assert.Equal("Create new workbook", window.Title);

            window.CreateForTests("No video");

            Assert.Null(window.Created!.OwnedDeviceId);
        });
    }

    // ###########################################################################################
    // "Add device" beside the list opens "Add device" on this board - and the device it adds is
    // picked for the repair, which goes on it; nothing added, the choice stays as it was.
    // ###########################################################################################
    [Fact]
    public async Task A_device_added_from_Create_new_workbook_gets_the_repair()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey);
            window.AddDeviceOverrideForTests = () => Task.FromResult<OwnedDeviceRecord?>(this.AddDevice("Plus/4 #1"));

            window.GetControl<Button>("AddDeviceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await window.LastNewDeviceForTests;

            ComboBox choices = window.GetControl<ComboBox>("DeviceComboBox");
            Assert.Equal(["None", "Plus/4 #1"], choices.Items.Select(item => item!.ToString()));
            Assert.Equal("Plus/4 #1", choices.SelectedItem!.ToString());
            Assert.True(window.DeviceAdded);

            window.CreateForTests("Black screen");

            Assert.Equal(OwnedDeviceStore.GetAll().Single().Id, window.Created!.OwnedDeviceId);
        });
    }

    [Fact]
    public async Task No_device_added_leaves_the_choice_as_it_was()
    {
        OwnedDeviceRecord device = this.AddDevice();

        await UiTest.RunAsync(async () =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey, device.Id);
            window.AddDeviceOverrideForTests = () => Task.FromResult<OwnedDeviceRecord?>(null);

            window.GetControl<Button>("AddDeviceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await window.LastNewDeviceForTests;

            ComboBox choices = window.GetControl<ComboBox>("DeviceComboBox");
            Assert.Equal(["None", "Breadbin #2"], choices.Items.Select(item => item!.ToString()));
            Assert.Equal("Breadbin #2", choices.SelectedItem!.ToString());
            Assert.False(window.DeviceAdded);

            window.CreateForTests("No video");

            Assert.Equal(device.Id, window.Created!.OwnedDeviceId);
            Assert.Single(OwnedDeviceStore.GetAll());
        });
    }

    // ###########################################################################################
    // Arrowing through "On my device" only picks a device - it never opens "Add device" (code review,
    // 2026-10-29: "A new device" was an item of the list, and reaching it with the Down key opened
    // the dialog).
    // ###########################################################################################
    [Fact]
    public void Going_through_the_device_list_never_opens_Add_device()
    {
        this.AddDevice();
        bool opened = false;

        UiTest.Run(() =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey);
            window.AddDeviceOverrideForTests = () =>
            {
                opened = true;
                return Task.FromResult<OwnedDeviceRecord?>(null);
            };

            ComboBox choices = window.GetControl<ComboBox>("DeviceComboBox");

            for (int index = 0; index < choices.ItemCount; index++)
            {
                choices.SelectedIndex = index;
                Dispatcher.UIThread.RunJobs();
            }

            Assert.False(opened);
        });
    }

    // ###########################################################################################
    // A device that comes back from "Add device" on ANOTHER board is no device for this repair: not
    // listed, not picked (code review, 2026-10-29: a C64 repair was filed on a C128 device). Its
    // board is this board's in that form, so only a stand-in can bring one back - the rule holds all
    // the same.
    // ###########################################################################################
    [Fact]
    public async Task A_device_added_on_another_board_is_not_picked_for_this_repair()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey);
            window.AddDeviceOverrideForTests = () => Task.FromResult<OwnedDeviceRecord?>(this.AddDevice("C128D", this.thisOtherBoardKey));

            window.GetControl<Button>("AddDeviceButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await window.LastNewDeviceForTests;

            ComboBox choices = window.GetControl<ComboBox>("DeviceComboBox");
            Assert.Equal(["None"], choices.Items.Select(item => item!.ToString()));
            Assert.Equal("None", choices.SelectedItem!.ToString());

            window.CreateForTests("No video");

            Assert.Null(window.Created!.OwnedDeviceId);
        });
    }

    // "Add device" opened from a repair keeps the repair's board - the hardware and board are shown,
    // not changeable.
    [Fact]
    public void Add_device_for_a_repair_keeps_the_repairs_board()
    {
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Commodore 64", BoardName = "250407" },
            new() { HardwareName = "Commodore 128", BoardName = "310378" },
        ];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407", boardFixed: true);

            Assert.False(window.GetControl<ComboBox>("HardwareComboBox").IsEnabled);
            Assert.False(window.GetControl<ComboBox>("BoardComboBox").IsEnabled);
            Assert.Equal("Commodore 64|250407", window.Details().BoardKey);
        });
    }

    // ###########################################################################################
    // The repair's board kept even when CRT no longer lists it - it is the board the device is for
    // (code review, 2026-10-29: the locked Board box fell back to another board, or none, and the
    // device was saved there without a word).
    // ###########################################################################################
    [Fact]
    public void Add_device_for_a_repair_keeps_a_board_CRT_no_longer_lists()
    {
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250425" }];

        UiTest.Run(() =>
        {
            var listedHardware = new OwnedDeviceWindow();
            listedHardware.Initialize(boards, "Commodore 64|250407", boardFixed: true);
            Assert.Equal("Commodore 64|250407", listedHardware.Details().BoardKey);

            var unlistedHardware = new OwnedDeviceWindow();
            unlistedHardware.Initialize(boards, "Commodore 128|310378", boardFixed: true);
            Assert.Equal("Commodore 128|310378", unlistedHardware.Details().BoardKey);
        });
    }

    // ###########################################################################################
    // The history of a device whose repair is on the board shown, before CRT has finished reading
    // that board (or could not): its pictures read from its board file, as any other board's (code
    // review, 2026-10-29: those repairs went out without their board).
    // ###########################################################################################
    [Fact]
    public async Task A_repair_on_the_board_shown_before_it_is_read_still_gets_its_pictures()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord repair = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", string.Empty, device.Id)!;
        WorklogManager.AddEntry(repair.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        var board = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "Images/top.png" }] };

        try
        {
            Dictionary<string, string> stored = PublishedBoardTree.Load(
                this.thisWorkspace.Path_("Data"), this.thisBoardKey, board, ("Images/top.png", [1, 2, 3]));

            await UiTest.RunAsync(async () =>
            {
                TabWorkbooks tab = this.BuildTab();
                tab.SelectDevice(device.Id);

                DeviceHistoryExportModel.Document history = (await tab.BuildDeviceHistoryForTestsAsync())!;

                Assert.Equal(stored["Images/top.png"], Assert.Single(Assert.Single(history.Repairs).Sections).SchematicImagePath);
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // A workbook whose index.json names no board ("boardKey": null - hand-edited or damaged) costs
    // that repair its board pictures, never the whole history (code review, 2026-10-29: the export
    // threw, and no file was written for any repair).
    // ###########################################################################################
    [Fact]
    public async Task A_repair_naming_no_board_is_still_in_the_history()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord repair = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", string.Empty, device.Id)!;

        string index = Path.Combine(WorklogManager.WorkbookRoot, WorklogManager.BuildWorkbookFolderName(repair.Id), "index.json");
        File.WriteAllText(index, File.ReadAllText(index).Replace($"\"{this.thisBoardKey}\"", "null", StringComparison.Ordinal));
        Assert.Null(WorklogManager.GetAllWorkbooks().Single(each => each.Id == repair.Id).BoardKey);

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            DeviceHistoryExportModel.Document history = (await tab.BuildDeviceHistoryForTestsAsync())!;

            Assert.Equal(repair.Id, Assert.Single(history.Repairs).WorkbookId);
        });
    }

    // ###########################################################################################
    // The title names the device the repair goes on - and follows the drop-down (code review,
    // 2026-10-29: "New repair on Breadbin #2" stayed while the repair went on none).
    // ###########################################################################################
    [Fact]
    public void The_title_follows_the_device_picked()
    {
        OwnedDeviceRecord device = this.AddDevice();

        UiTest.Run(() =>
        {
            var window = new CreateWorkbookWindow();
            window.Initialize(this.thisBoardKey, device.Id);
            Assert.Equal("New repair on Breadbin #2", window.Title);

            ComboBox choices = window.GetControl<ComboBox>("DeviceComboBox");
            choices.SelectedIndex = 0;
            Assert.Equal("Create new workbook", window.Title);

            choices.SelectedIndex = 1;
            Assert.Equal("New repair on Breadbin #2", window.Title);
        });
    }

    // Editing a workbook asks nothing about devices - "On my device" under its title does that.
    [Fact]
    public void Edit_workbook_asks_no_device()
    {
        this.AddDevice();

        UiTest.Run(() =>
        {
            WorkbookRecord workbook = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", string.Empty)!;

            var window = new CreateWorkbookWindow();
            window.InitializeForEdit(workbook);

            Assert.False(window.GetControl<StackPanel>("DevicePanel").IsVisible);
        });
    }

    // ###########################################################################################
    // "Add device": the board shown preselected among the boards CRT knows, no region, Working -
    // and what is saved is what the form shows.
    // ###########################################################################################
    [Fact]
    public void Add_device_starts_on_the_board_shown_and_saves_what_the_form_holds()
    {
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Commodore 64", BoardName = "250407" },
            new() { HardwareName = "Commodore 64", BoardName = "250425" },
            new() { HardwareName = "Commodore 128", BoardName = "310378" },
        ];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250425");
            window.Show();

            Assert.Equal("Commodore 64", window.GetControl<ComboBox>("HardwareComboBox").SelectedItem);
            Assert.Equal(["250407", "250425"], window.GetControl<ComboBox>("BoardComboBox").Items.Cast<string>());
            Assert.Equal("250425", window.GetControl<ComboBox>("BoardComboBox").SelectedItem);
            Assert.Equal(OwnedDeviceWindow.NoRegion, window.GetControl<ComboBox>("RegionComboBox").SelectedItem);
            Assert.Equal("Working", window.GetControl<ComboBox>("StatusComboBox").SelectedItem);
            Assert.False(window.GetControl<Button>("SaveButton").IsEnabled);

            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.GetControl<ComboBox>("RegionComboBox").SelectedItem = "PAL";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            OwnedDeviceRecord saved = Assert.Single(OwnedDeviceStore.GetAll());
            Assert.Equal(("Breadbin #2", "Commodore 64|250425", "PAL"), (saved.Name, saved.BoardKey, saved.Region));
            Assert.Equal(saved.Id, window.Saved!.Id);
        });
    }

    // ###########################################################################################
    // Enter saves the form from a box - but Enter on a focused button is the button's own: on Cancel
    // it saved the form (code review, 2026-10-28).
    // ###########################################################################################
    [Fact]
    public void Enter_saves_from_a_box_but_never_from_the_Cancel_button()
    {
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407");
            window.Show();
            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";

            Button cancel = window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, "Cancel"));
            cancel.RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.Enter });

            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.Null(window.Saved);

            window.GetControl<TextBox>("NameTextBox").RaiseEvent(new Avalonia.Input.KeyEventArgs { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.Enter });

            Assert.Single(OwnedDeviceStore.GetAll());
        });
    }

    // Edited onto another board: the swap is recorded; a board CRT no longer lists stays offered.
    [Fact]
    public void Edit_device_keeps_a_board_CRT_no_longer_lists_and_records_a_swap()
    {
        OwnedDeviceRecord device = this.AddDevice("Breadbin #2", "Commodore 64|250466");
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.InitializeForEdit(device, boards);
            window.Show();

            Assert.Equal("Save device", window.GetControl<Button>("SaveButton").Content);
            Assert.Equal(["250407", "250466"], window.GetControl<ComboBox>("BoardComboBox").Items.Cast<string>());
            Assert.Equal("250466", window.GetControl<ComboBox>("BoardComboBox").SelectedItem);

            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "250407";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            OwnedDeviceRecord saved = OwnedDeviceStore.Get(device.Id)!;
            Assert.Equal("Commodore 64|250407", saved.BoardKey);
            Assert.Equal("Board changed from Commodore 64 - 250466 to Commodore 64 - 250407", Assert.Single(saved.Changes).Text);
        });
    }

    // ###########################################################################################
    // "EXPORT HISTORY" (slice 4d - the plan: "Export history puts every repair into one PDF or
    // ZIP"): the device page's two buttons, the file suggested by the device's name, written where it
    // was picked with the button's own extension - and nothing when the dialog is called off.
    // ###########################################################################################
    [Fact]
    public async Task Export_history_writes_the_devices_history_to_the_file_picked()
    {
        WorkbookPdfExporter.ConfigureQuestPdf();
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord repair = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", string.Empty, device.Id)!;
        WorklogManager.AddEntry(repair.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());

        string target = Path.Combine(this.thisWorkspace.Root, "Exports", "history");
        Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        var asked = new List<(bool AsZip, string Suggested)>();

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            Assert.Equal("Export history to PDF", tab.GetControl<Button>("ExportHistoryButton").Content);
            Assert.Equal("Export history to ZIP", tab.GetControl<Button>("ExportHistoryZipButton").Content);

            tab.PickHistoryPathOverrideForTests = (asZip, suggested) =>
            {
                asked.Add((asZip, suggested));
                return Task.FromResult<string?>(target);
            };

            await tab.ExportDeviceHistoryAsync(asZip: false);
        });

        (bool asZip, string suggested) = Assert.Single(asked);
        Assert.False(asZip);
        Assert.StartsWith($"Device_{device.Id}_Breadbin-2_", suggested);
        Assert.True(new FileInfo(target + ".pdf").Length > 0);
    }

    [Fact]
    public async Task Export_history_called_off_writes_nothing()
    {
        OwnedDeviceRecord device = this.AddDevice();
        string exports = this.thisWorkspace.Path_("Exports");
        Directory.CreateDirectory(exports);

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);
            tab.PickHistoryPathOverrideForTests = (_, _) => Task.FromResult<string?>(null);

            await tab.ExportDeviceHistoryAsync(asZip: true);
        });

        Assert.Empty(Directory.GetFiles(exports));
    }

    // ###########################################################################################
    // Each repair is drawn on ITS board's pictures: the board shown's as CRT holds them, another
    // board's read from its board file (CRT 4.0's stored picture here), and one on a board CRT no
    // longer lists without - never every repair on the board shown's pictures.
    // ###########################################################################################
    [Fact]
    public async Task Each_repair_in_the_history_is_drawn_on_its_own_boards_pictures()
    {
        OwnedDeviceRecord device = this.AddDevice();
        string gone = "Commodore 64|gone " + Guid.NewGuid().ToString("N");

        WorkbookRecord onShown = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", string.Empty, device.Id)!;
        WorkbookRecord onOther = WorklogManager.CreateWorkbook(this.thisOtherBoardKey, "Before the swap", string.Empty, device.Id)!;
        WorkbookRecord onGone = WorklogManager.CreateWorkbook(gone, "Long ago", string.Empty, device.Id)!;

        foreach (WorkbookRecord repair in new[] { onShown, onOther, onGone })
        {
            WorklogManager.AddEntry(repair.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "Bad cap", "", "Issue", "Open", Array.Empty<string>());
        }

        string shownPicture = this.Picture("shown.png");
        var shownBoard = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = shownPicture }] };
        var otherBoard = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "Images/top.png" }] };

        try
        {
            Dictionary<string, string> stored = PublishedBoardTree.Load(
                this.thisWorkspace.Path_("Data"), this.thisOtherBoardKey, otherBoard, ("Images/top.png", [1, 2, 3]));

            await UiTest.RunAsync(async () =>
            {
                TabWorkbooks tab = this.BuildTab();
                tab.CurrentBoardDataOverrideForTests = shownBoard;
                tab.SelectDevice(device.Id);

                DeviceHistoryExportModel.Document history = (await tab.BuildDeviceHistoryForTestsAsync())!;

                string? PictureOf(WorkbookRecord repair) =>
                    Assert.Single(history.Repairs.Single(each => each.WorkbookId == repair.Id).Sections).SchematicImagePath;

                Assert.Equal([onShown.Id, onOther.Id, onGone.Id], history.Repairs.Select(each => each.WorkbookId));
                Assert.Equal(shownPicture, PictureOf(onShown));
                Assert.Equal(stored["Images/top.png"], PictureOf(onOther));
                Assert.Null(PictureOf(onGone));
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }
}
