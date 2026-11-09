using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The worklog bar's "Device" and "Workbook" drop-downs (Main.Devices.cs - CRT 4.0 step 4, slice 4c, above
// Hardware until 2026-11-06, the bar's first drop-down since). Owner requests of that day: "a device selector as
// the first drop-down, and then it will show all workbooks for this device in the other 'Workbook' selector", a
// badge at the end of each line, "rename 'No device' to 'Show workbooks for all devices' and then have that as
// first selection ... per default if never set before", "a counter for number of workbooks", a device with no
// workbook going "to the 'My Devices' tab" with that device selected, and the two drop-downs as wide as their
// text. The device drop-down is a remembered filter - never one hiding the workbook "Add worklog" writes into.
// Main's StartAsync is never called (MainWindowTests' header); the board shown is BoardKeyOverrideForTests.
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

    private static ComboBox DevicesOf(CRT.Main main) => main.GetControl<ComboBox>("DevicePickerComboBox");

    private static ComboBox WorkbooksOf(CRT.Main main) => main.GetControl<ComboBox>("WorklogJobBox");

    // The device drop-down's choices, as its lines read them: the device - null for all devices - and its count.
    private static List<BarDeviceChoice> ChoicesOf(CRT.Main main) => DevicesOf(main).Items.OfType<BarDeviceChoice>().ToList();

    private static List<int> WorkbookIdsOf(CRT.Main main) => WorkbooksOf(main).Items.OfType<WorkbookRecord>().Select(workbook => workbook.Id).ToList();

    // What the user does: a choice picked in the drop-down - a device, or null for all devices - and what that started.
    private static void Pick(CRT.Main main, OwnedDeviceRecord? device)
    {
        ComboBox picker = DevicesOf(main);
        picker.SelectedItem = ChoicesOf(main).Single(choice => choice.Device?.Id == device?.Id);
        Dispatcher.UIThread.RunJobs();
    }

    // The words of a line of either picker - after the device badge it may open with (2026-11-07), before the badges
    // at its end.
    private static int WordsAt(Control line) =>
        Assert.IsType<BarPickerLinePanel>(line).Children.ToList().FindIndex(child => child.Classes.Contains("BarPickerText"));

    private static string LineText(Control line) =>
        Assert.IsType<TextBlock>(Assert.IsType<BarPickerLinePanel>(line).Children[WordsAt(line)]).Text ?? string.Empty;

    // The first badge after a line's words, and its words - its last text, after a workbook's padlock glyph.
    private static string BadgeText(Control line) =>
        Assert.IsType<Border>(Assert.IsType<BarPickerLinePanel>(line).Children[WordsAt(line) + 1])
            .GetLogicalDescendants().OfType<TextBlock>().Last().Text ?? string.Empty;

    // The device badge a workbook's line opens with - null with none.
    private static string? LeadingText(Control line) =>
        WordsAt(line) == 0
            ? null
            : Assert.IsType<Border>(Assert.IsType<BarPickerLinePanel>(line).Children[0]).GetLogicalDescendants().OfType<TextBlock>().Single().Text;

    // The last badge of a line - a device line's count of workbooks - and its number.
    private static string CountText(Control line) =>
        Assert.IsType<Border>(Assert.IsType<BarPickerLinePanel>(line).Children[^1])
            .GetLogicalDescendants().OfType<TextBlock>().Single().Text ?? string.Empty;

    // ###########################################################################################
    // THE BAR ADDS A DEVICE TOO, WITH NO DEVICE AND NO WORKBOOK THERE YET (owner request, 2026-11-06: "if you
    // have created no devices and no workbooks, then all you see in that top-bar is 'Create new workbook'"):
    // "Add device", green as it adds; added, the "Workbooks" tab is shown on its page, and no workbook was made.
    // Called off, nothing is added and the tab shown stays.
    // ###########################################################################################
    [Fact]
    public async Task The_bar_adds_a_device_with_no_workbook_and_shows_its_page()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            Button add = main.GetControl<Button>("WorklogAddDeviceButton");
            TabControl tabs = main.GetControl<TabControl>("MainTabControl");
            TabWorkbooks workbooks = main.GetControl<TabWorkbooks>("TabWorkbooks");

            Assert.True(main.GetControl<Border>("WorklogBar").IsVisible);
            Assert.True(add.IsVisible);
            Assert.Contains("Accept", add.Classes);
            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.Empty(WorklogManager.GetAllWorkbooks());

            workbooks.BoardsOverrideForTests = [new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407" }];
            workbooks.BoardRegionsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>([]);

            // Called off: nothing.
            workbooks.ShowDeviceFormOverrideForTests = _ => Task.FromResult<OwnedDeviceRecord?>(null);
            object? shownBefore = tabs.SelectedItem;

            await main.AddDeviceFromBarAsync();

            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.Same(shownBefore, tabs.SelectedItem);

            workbooks.ShowDeviceFormOverrideForTests = form =>
            {
                form.Show();
                form.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
                form.GetControl<Button>("SaveButton").RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));

                return Task.FromResult(form.Saved);
            };

            await main.AddDeviceFromBarAsync();

            OwnedDeviceRecord added = Assert.Single(OwnedDeviceStore.GetAll());
            Assert.Equal("Breadbin #2", added.Name);
            Assert.Empty(WorklogManager.GetAllWorkbooks());
            Assert.Same(main.GetControl<TabItem>("WorkbooksTabItem"), tabs.SelectedItem);
            Assert.Equal(added.Id, workbooks.SelectedDeviceIdForTests);
        });
    }

    // ###########################################################################################
    // THE BAR'S "ADD DEVICE" OPENS ITS FORM WHICHEVER TAB IS SHOWN (owner report, 2026-11-06: "clicking the 'Add
    // device' in the top-bar result in nothing - it does not open the modal for it"). The form was shown over the
    // window the "Workbooks" tab is in - and a tab not shown is out of the window's visual tree, so from any other
    // tab there was no window to show it over, and nothing happened. Clicked as a person clicks it, no stand-in.
    // ###########################################################################################
    [Fact]
    public void The_bars_Add_device_opens_its_form_whichever_tab_is_shown()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();
            main.Show();

            try
            {
                TabControl tabs = main.GetControl<TabControl>("MainTabControl");
                tabs.SelectedItem = main.GetControl<TabItem>("SchematicsTabItem");
                Dispatcher.UIThread.RunJobs();

                Button add = main.GetControl<Button>("WorklogAddDeviceButton");
                add.RaiseEvent(new Avalonia.Interactivity.RoutedEventArgs(Button.ClickEvent));
                Dispatcher.UIThread.RunJobs();

                OwnedDeviceWindow form = Assert.Single(main.OwnedWindows.OfType<OwnedDeviceWindow>());
                form.Close();
                Dispatcher.UIThread.RunJobs();

                Assert.Empty(OwnedDeviceStore.GetAll());
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // THE DEVICE DROP-DOWN IS THE BAR'S FIRST (owner request, 2026-11-06 - above Hardware until then): "Show workbooks
    // for all devices" first and picked, and gone with the bar when the worklog is off. NOT THERE UNTIL A DEVICE IS
    // (owner request, 2026-11-08: "if no devices or workbooks exists yet ... then its drop-down boxes should not show
    // either" - greyed out until then), its label with it, the "Workbook" one then standing first.
    // ###########################################################################################
    [Fact]
    public void The_device_drop_down_is_the_bars_first_with_all_devices_first_and_not_there_with_no_device()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            Border bar = main.GetControl<Border>("WorklogBar");
            ComboBox picker = DevicesOf(main);
            Panel pickers = main.GetControl<Panel>("WorklogPickersPanel");
            TextBlock heading = main.GetControl<TextBlock>("DevicePickerHeadingText");
            TextBlock workbookHeading = main.GetControl<TextBlock>("WorkbookPickerHeadingText");

            Assert.Contains(bar, picker.GetLogicalAncestors());
            Assert.True(pickers.Children.IndexOf(picker) < pickers.Children.IndexOf(WorkbooksOf(main)));
            Assert.Equal("Device", heading.Text);
            Assert.False(picker.IsVisible);
            Assert.False(heading.IsVisible);
            Assert.Equal(0, workbookHeading.Margin.Left);
            Assert.Null(Assert.IsType<BarDeviceChoice>(picker.SelectedItem).Device);

            AddDevice("Breadbin #2");
            main.RefreshWorklogBar();

            Assert.True(picker.IsVisible);
            Assert.True(heading.IsVisible);
            Assert.True(picker.IsEnabled);
            Assert.True(workbookHeading.Margin.Left > 0);
            Assert.Equal([null, "Breadbin #2"], ChoicesOf(main).Select(choice => choice.Device?.Name));

            UserSettings.EnableWorklog = false;
            main.ApplyWorklogBarVisibility();

            Assert.False(bar.IsVisible);
        });
    }

    // ###########################################################################################
    // "SHOW WORKBOOKS FOR ALL DEVICES" - THE FIRST CHOICE, PICKED TILL ANOTHER IS (owner request, 2026-11-06:
    // "per default if never set before") - lists EVERY workbook, whatever device and board it is on; each line of
    // the device drop-down says how many workbooks it holds ("a counter for number of workbooks") and a device's
    // ends in its status, a workbook's in Open or Closed.
    // ###########################################################################################
    [Fact]
    public void All_devices_is_the_default_lists_every_workbook_and_each_line_counts_its_workbooks()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddDevice("C64C silver");
        WorkbookRecord onDevice = AddRepair("No video", breadbin.Id);
        WorkbookRecord elsewhere = AddRepair("PSU replaced", breadbin.Id, OtherBoardKey);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = DevicesOf(main);
            ComboBox bar = WorkbooksOf(main);

            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.Null(main.DeviceWorkedOn);
            Assert.Null(Assert.IsType<BarDeviceChoice>(picker.SelectedItem).Device);
            Assert.Equal([onNone.Id, elsewhere.Id, onDevice.Id], WorkbookIdsOf(main));
            Assert.Equal(onNone.Id, Assert.IsType<WorkbookRecord>(bar.SelectedItem).Id);

            List<BarDeviceChoice> choices = ChoicesOf(main);
            Assert.Equal([3, 2, 0], choices.Select(choice => choice.WorkbookCount));

            // The count an outlined badge, the line's last (owner request, 2026-11-06: "I actually did mean an
            // outlined badge, and please show that as the last, right of the status").
            Control all = picker.ItemTemplate!.Build(choices[0])!;
            Assert.Equal("Show workbooks for all devices", LineText(all));
            Assert.Equal(2, Assert.IsType<BarPickerLinePanel>(all).Children.Count);
            Assert.Equal("3", CountText(all));

            Control device = picker.ItemTemplate!.Build(choices[1])!;
            Assert.Equal("Breadbin #2", LineText(device));
            Assert.Equal(OwnedDeviceStatus.InRepair, BadgeText(device));
            Assert.Equal("2", CountText(device));
            Assert.Equal(3, Assert.IsType<BarPickerLinePanel>(device).Children.Count);
            Assert.Equal("0", CountText(picker.ItemTemplate!.Build(choices[2])!));

            // A workbook on the board shown reads without it; on another, with it (no synced data in a test, so
            // the board is its key as stored).
            Control line = bar.ItemTemplate!.Build(onDevice)!;
            Assert.Equal($"#{onDevice.Id} | No video", LineText(line));
            Assert.Equal("Open", BadgeText(line));
            Assert.Equal($"#{elsewhere.Id} | PSU replaced ({OtherBoardKey})", LineText(bar.ItemTemplate!.Build(elsewhere)!));

            // ###########################################################################################
            // EVERY DEVICE'S WORKBOOKS LISTED, EACH LINE OPENS WITH ITS DEVICE'S NAME IN A BADGE (owner request,
            // 2026-11-07: "when I have selected 'Show workbooks for all devices', then I would like for the workbook
            // drop-down to show the device name first in a badge ... so I can separate the workbooks in the list") - a
            // workbook on no device none; and with one device picked, no line has one.
            // ###########################################################################################
            Assert.Equal("Breadbin #2", LeadingText(line));
            Assert.Null(LeadingText(bar.ItemTemplate!.Build(onNone)!));

            Pick(main, breadbin);
            Assert.Null(LeadingText(WorkbooksOf(main).ItemTemplate!.Build(onDevice)!));
        });
    }

    // ###########################################################################################
    // A DEVICE PICKED NARROWS THE WORKBOOKS TO ITS OWN, on any board, works on its repair, and is remembered -
    // CRT opened again shows the same; all devices picked again lists every workbook, the bar's workbook as it was.
    // ###########################################################################################
    [Fact]
    public void A_device_picked_lists_its_workbooks_and_is_remembered()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord onDevice = AddRepair("No video", breadbin.Id);
        WorkbookRecord elsewhere = AddRepair("PSU replaced", breadbin.Id, OtherBoardKey);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();

            Pick(main, breadbin);

            Assert.Equal(breadbin.Id, UserSettings.WorklogBarDeviceId);
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
            Assert.Equal(onDevice.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal([elsewhere.Id, onDevice.Id], WorkbookIdsOf(main));

            CRT.Main again = BuildMain();
            Assert.Equal(breadbin.Id, again.DeviceWorkedOn!.Id);
            Assert.Equal([elsewhere.Id, onDevice.Id], WorkbookIdsOf(again));

            Pick(main, null);

            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.Null(main.DeviceWorkedOn);
            Assert.Equal([onNone.Id, elsewhere.Id, onDevice.Id], WorkbookIdsOf(main));
            Assert.Equal(onDevice.Id, UserSettings.GetActiveWorkbookId(BoardKey));
        });
    }

    // ###########################################################################################
    // A WORKBOOK ACTIVATED OFF THE DEVICE PICKED TAKES THE PICK ALONG (OwnedDeviceRepairs.PickFollowing): CRT works only
    // in the device picked's workbooks (owner decision, 2026-11-06: "Only the device's"), so one activated on another
    // device picks that device, and one on none all devices - remembered either way; the bar works in it.
    // ###########################################################################################
    [Fact]
    public void A_workbook_activated_off_the_device_picked_takes_the_pick_to_its_own_device()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        OwnedDeviceRecord silver = AddDevice("C64C silver");
        WorkbookRecord onDevice = AddRepair("No video", breadbin.Id);
        WorkbookRecord onSilver = AddRepair("Dead SID", silver.Id);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            Pick(main, breadbin);
            Assert.Equal(onDevice.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);

            main.ActivateWorkbook(BoardKey, onSilver.Id);

            Assert.Equal(silver.Id, main.DeviceWorkedOn!.Id);
            Assert.Equal(silver.Id, UserSettings.WorklogBarDeviceId);
            Assert.Equal(onSilver.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);

            main.ActivateWorkbook(BoardKey, onNone.Id);

            Assert.Null(main.DeviceWorkedOn);
            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.Equal(onNone.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);
        });
    }

    // ###########################################################################################
    // WHILE A DEVICE IS PICKED CRT WORKS ONLY IN ITS WORKBOOKS (owner decision, 2026-11-06, asked: "Only the device's"):
    // the active workbook is its own on the board shown - one saved for the board on another device is passed over, and
    // "Add worklog" and the component popup's capture write into the device's (Main.ResolveActiveWorkbookForBoard) -
    // and with none of its own there none at all: the "Workbook" drop-down says so, "Add worklog" is hidden. "Show
    // workbooks for all devices" brings the board's own back.
    // ###########################################################################################
    [Fact]
    public void A_device_picked_works_only_in_its_own_workbooks_and_with_none_on_the_board_in_none()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        OwnedDeviceRecord silver = AddDevice("C64C silver");
        WorkbookRecord onBreadbin = AddRepair("No video", breadbin.Id);
        WorkbookRecord onNone = AddRepair("Dead keyboard", null);
        AddRepair("Dead SID", silver.Id, OtherBoardKey);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            Pick(main, breadbin);

            // The board's saved workbook on no device is passed over while Breadbin #2 is picked.
            UserSettings.SetActiveWorkbookId(BoardKey, onNone.Id);
            main.RefreshWorklogBar();

            Assert.Equal(onBreadbin.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);
            Assert.Equal(onBreadbin.Id, CRT.Main.ResolveActiveWorkbookForBoard(BoardKey)!.Id);
            Assert.True(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);

            // Silver's only workbook is on another board: it stays picked on the board shown - its own - and
            // nothing there is worked in.
            UserSettings.WorklogBarDeviceId = silver.Id;
            main.RefreshWorklogBar();

            Assert.Equal(silver.Id, main.DeviceWorkedOn!.Id);
            Assert.Null(WorkbooksOf(main).SelectedItem);
            Assert.Equal(OwnedDeviceWording.NoWorkbookOfDeviceOnBoard, WorkbooksOf(main).PlaceholderText);
            Assert.Null(CRT.Main.ResolveActiveWorkbookForBoard(BoardKey));
            Assert.False(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);

            Pick(main, null);

            Assert.Equal(onNone.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);
            Assert.True(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);
        });
    }

    // ###########################################################################################
    // Another board shown with none of the device picked's workbooks - the drop-downs at the top switched - is all
    // devices again (OwnedDeviceRepairs.BarFilter), rather than nothing worked in there.
    // ###########################################################################################
    [Fact]
    public void Another_board_shown_with_none_of_the_devices_workbooks_is_all_devices_again()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id);
        WorkbookRecord elsewhere = AddRepair("Dead keyboard", null, OtherBoardKey);

        try
        {
            // Both boards listed: the device's own is one CRT can go back to.
            ListBoards(BoardKey, OtherBoardKey);

            UiTest.Run(() =>
            {
                CRT.Main main = BuildMain();
                Pick(main, breadbin);
                Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

                main.BoardKeyOverrideForTests = OtherBoardKey;
                main.RefreshWorklogBar();

                Assert.Null(main.DeviceWorkedOn);
                Assert.Null(UserSettings.WorklogBarDeviceId);
                Assert.Equal(elsewhere.Id, Assert.IsType<WorkbookRecord>(WorkbooksOf(main).SelectedItem).Id);
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // A device that has gone is no filter: all devices - and with the worklog off nothing is read, so what was
    // picked is kept for when it comes back.
    [Fact]
    public void A_device_gone_reads_as_all_devices_and_the_worklog_off_keeps_the_pick()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id);

        UiTest.Run(() =>
        {
            UserSettings.WorklogBarDeviceId = breadbin.Id;
            UserSettings.EnableWorklog = false;
            CRT.Main main = BuildMain();
            Assert.Equal(breadbin.Id, UserSettings.WorklogBarDeviceId);

            UserSettings.EnableWorklog = true;
            main.RefreshWorklogBar();
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            Assert.True(OwnedDeviceStore.Delete(breadbin.Id));
            main.RefreshWorklogBar();

            Assert.Null(main.DeviceWorkedOn);
            Assert.Null(UserSettings.WorklogBarDeviceId);
        });
    }

    // ###########################################################################################
    // A DEVICE WITH NO WORKBOOK OPENS ITS PAGE ON THE "MY DEVICES" TAB (owner request, 2026-11-06: "it should not
    // open the 'create new workbook' modal, but instead it should go to the 'My Devices' tab and select that
    // device"): nothing is made, and it STAYS PICKED - CRT works only in its workbooks, and it has none, so the
    // "Workbook" drop-down says so and "Add worklog" is hidden (owner decision, 2026-11-06: "Only the device's"; the
    // drop-down went back as it was until then).
    // ###########################################################################################
    [Fact]
    public void Picking_a_device_with_no_workbook_opens_its_page_on_the_My_Devices_tab()
    {
        OwnedDeviceRecord plus4 = AddDevice("Plus/4");
        AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();
            TabControl tabs = main.GetControl<TabControl>("MainTabControl");
            Assert.NotSame(main.GetControl<TabItem>("WorkbooksTabItem"), tabs.SelectedItem);

            Pick(main, plus4);

            Assert.Same(main.GetControl<TabItem>("WorkbooksTabItem"), tabs.SelectedItem);
            Assert.Equal(plus4.Id, main.TabWorkbooks.SelectedDeviceIdForTests);
            Assert.True(main.TabWorkbooks.IsDevicePageShown);

            Assert.Single(WorklogManager.GetAllWorkbooks());
            Assert.Empty(main.OwnedWindows);
            Assert.Equal(plus4.Id, main.DeviceWorkedOn!.Id);
            Assert.Equal(plus4.Id, UserSettings.WorklogBarDeviceId);
            Assert.Null(WorkbooksOf(main).SelectedItem);
            Assert.Empty(WorkbookIdsOf(main));
            Assert.False(WorkbooksOf(main).IsEnabled);
            Assert.Equal(OwnedDeviceWording.NoWorkbookOfDeviceYet, WorkbooksOf(main).PlaceholderText);
            Assert.False(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);
        });
    }

    // A device whose only workbook is on a board CRT no longer lists has none to work on either: its page, picked.
    [Fact]
    public void A_device_whose_workbooks_are_all_on_boards_gone_opens_its_page()
    {
        OwnedDeviceRecord c128 = AddDevice("C128D");
        AddRepair("No video", c128.Id, "Nowhere|" + Guid.NewGuid().ToString("N"));

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            Pick(main, c128);

            Assert.Same(main.GetControl<TabItem>("WorkbooksTabItem"), main.GetControl<TabControl>("MainTabControl").SelectedItem);
            Assert.Equal(c128.Id, main.TabWorkbooks.SelectedDeviceIdForTests);
            Assert.Equal(c128.Id, main.DeviceWorkedOn!.Id);
            Assert.Null(WorkbooksOf(main).SelectedItem);
        });
    }

    // ###########################################################################################
    // Too narrow for its badges, a line keeps its words and lets them go (2026-11-06: on a narrow window the
    // closed device box read "Working" and not "C64"); wide enough, they stand side by side at the line's end,
    // the count last.
    // ###########################################################################################
    [Fact]
    public void A_line_too_narrow_for_its_badges_keeps_its_words_and_lets_them_go()
    {
        UiTest.Run(() =>
        {
            BarPickerLinePanel Laid(double width)
            {
                BarPickerLinePanel line = CRT.Main.BarPickerLine(
                    "C64 Breadbin Henrik",
                    Handlers.Theming.WorklogInfoPillBuilder.BuildDeviceStatusPill(OwnedDeviceStatus.Working),
                    Handlers.Theming.WorklogInfoPillBuilder.BuildCountPill(3));
                var window = new Window { Content = new Border { Width = width, Child = line }, Width = 600, Height = 100 };
                window.Show();
                window.UpdateLayout();
                window.Close();
                return line;
            }

            BarPickerLinePanel wide = Laid(400);
            Assert.True(wide.BadgeShown);
            Assert.Equal(400, wide.Children[2].Bounds.Right, 1.0);
            Assert.True(wide.Children[1].Bounds.Right <= wide.Children[2].Bounds.Left + 0.5, "the count is not after the status");
            Assert.True(wide.Children[0].Bounds.Width >= wide.Children[0].DesiredSize.Width - 0.5);

            double badges = wide.Children[1].DesiredSize.Width + wide.Children[2].DesiredSize.Width;
            BarPickerLinePanel narrow = Laid(badges + 10);
            Assert.False(narrow.BadgeShown);
            Assert.Equal(narrow.Bounds.Width, narrow.Children[0].Bounds.Width, 1.0);
            Assert.True(narrow.Children[1].Bounds.Left >= narrow.Bounds.Width - 0.5, "a badge is laid out where the line is drawn");
            Assert.True(narrow.ClipToBounds);
        });
    }

    // ###########################################################################################
    // A line opening with a device badge (2026-11-07): the badge first, the words after it, the state at the end - and
    // a long device name cut to its share of the line, so the workbook's own words keep their room.
    // ###########################################################################################
    [Fact]
    public void A_line_opening_with_a_device_badge_keeps_the_words_after_it_and_cuts_a_long_name_short()
    {
        UiTest.Run(() =>
        {
            BarPickerLinePanel Laid(string device, double width)
            {
                BarPickerLinePanel line = CRT.Main.BarPickerLine(
                    Handlers.Theming.WorklogInfoPillBuilder.BuildDeviceNamePill(device),
                    "#3 | No video",
                    Handlers.Theming.WorklogInfoPillBuilder.BuildStatePill("Open"));
                var window = new Window { Content = new Border { Width = width, Child = line }, Width = 900, Height = 100 };
                window.Show();
                window.UpdateLayout();
                window.Close();
                return line;
            }

            BarPickerLinePanel line = Laid("C64", 400);
            Assert.Equal(0, line.Children[0].Bounds.Left, 0.5);
            Assert.True(line.Children[1].Bounds.Left >= line.Children[0].Bounds.Right - 0.5, "the words stand before the device badge");
            Assert.Equal(400, line.Children[2].Bounds.Right, 1.0);

            BarPickerLinePanel crowded = Laid("C64 Breadbin Henrik with a very long name indeed", 300);
            Assert.True(crowded.LeadingWidthShown <= 300 * Handlers.Geometry.BarPickerLineGeometry.LeadingShare + 0.5);
            Assert.True(crowded.Children[1].Bounds.Width > 0, "the workbook's words have no room");
        });
    }

    // On the board shown a workbook reads without its board; on another, with it - capitals aside.
    [Fact]
    public void A_workbook_names_its_board_only_when_it_is_not_the_board_shown()
    {
        var workbook = new WorkbookRecord { Id = 4, Title = "No video", BoardKey = BoardKey, OwnedDeviceId = 99 };

        Assert.Equal("#4 | No video", CRT.Main.WorkbookPickerText(workbook, BoardKey));
        Assert.Equal("#4 | No video", CRT.Main.WorkbookPickerText(workbook, BoardKey.ToUpperInvariant()));
        Assert.Equal($"#4 | No video ({BoardKey})", CRT.Main.WorkbookPickerText(workbook, OtherBoardKey));
    }

    // ###########################################################################################
    // With no workbook at all the workbook picker is not there (owner request, 2026-11-08 - it said "No workbook on this
    // board yet", greyed out, until then). With "Show workbooks for all devices" picked and workbooks of other boards to
    // list, it asks for one (owner request, 2026-11-07: '"Please select a workbook"').
    // ###########################################################################################
    [Fact]
    public void With_no_workbook_on_the_board_the_workbook_picker_asks_for_one_of_another_board()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox bar = WorkbooksOf(main);

            Assert.Null(bar.SelectedItem);
            Assert.False(bar.IsVisible);

            AddRepair("Dead keyboard", null, OtherBoardKey);
            main.RefreshWorklogBar();

            Assert.True(bar.IsVisible);
            Assert.True(bar.IsEnabled);
            Assert.Null(bar.SelectedItem);
            Assert.Equal("Select a workbook", bar.PlaceholderText);

            AddRepair("No video", null);
            main.RefreshWorklogBar();

            Assert.NotNull(bar.SelectedItem);
            Assert.Null(bar.PlaceholderText);
        });
    }

    // ###########################################################################################
    // THE BAR'S BUTTONS: "Add worklog", "Add workbook" ("Create new workbook" until 2026-11-06), then "Add device"
    // last (owner request), each green as it adds - and the workbook's counts and dates beside the picker are gone
    // (owner request: "remove the highlighted text in the top-bar, to reduce what is shown").
    // ###########################################################################################
    [Fact]
    public void The_bar_has_its_three_buttons_in_order_and_no_text_beside_the_pickers()
    {
        AddRepair("No video", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            Border bar = main.GetControl<Border>("WorklogBar");
            Button addWorklog = main.GetControl<Button>("WorklogAddEntryButton");
            Button addWorkbook = main.GetControl<Button>("WorklogCreateWorkbookButton");
            Button addDevice = main.GetControl<Button>("WorklogAddDeviceButton");

            var buttons = Assert.IsAssignableFrom<Panel>(addDevice.Parent).Children.OfType<Button>().Where(button => button.IsVisible).ToList();
            Assert.Equal(["Add worklog", "Add workbook", "Add device"], buttons.Select(button => button.Content as string));
            Assert.All(buttons, button => Assert.Contains("Accept", button.Classes));
            Assert.Same(addWorklog, buttons[0]);
            Assert.Same(addWorkbook, buttons[1]);

            Assert.DoesNotContain(
                bar.GetLogicalDescendants().OfType<TextBlock>().Where(text => text.IsVisible),
                text => (text.Text ?? string.Empty).Contains("started", StringComparison.Ordinal) || (text.Text ?? string.Empty).Contains("worklogs yet", StringComparison.Ordinal));
        });
    }

    // ###########################################################################################
    // EVERY BUTTON OF THE BAR IS ON SCREEN, HOWEVER LONG THE NAMES (owner report, 2026-11-06: "the 'Create new
    // workbook' button is not fully visible"): the drop-downs give way, their words cut short, never a button. At a
    // narrow window and the owner's 1536. Not at CRT's narrowest, 800: the headless platform measures text with
    // stand-in fonts nearly twice as wide as CRT's ("Add worklog" 143px for some 85), so what fits there is seen
    // in the running app, not here.
    // ###########################################################################################
    [Theory]
    [InlineData(1100)]
    [InlineData(1536)]
    public void Every_button_of_the_bar_is_on_screen_however_long_the_names(int windowWidth)
    {
        OwnedDeviceRecord device = AddDevice("C64 Breadbin from the flea market with the broken keyboard");
        AddRepair("No video and a dead keyboard after the PSU went up in smoke on a Sunday", device.Id);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.Width = windowWidth;
            main.Height = 800;
            main.ApplyWorklogBarVisibility();
            main.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                main.UpdateLayout();

                Border bar = main.GetControl<Border>("WorklogBar");
                double barRight = bar.TranslatePoint(new Avalonia.Point(bar.Bounds.Width, 0), main)!.Value.X;

                foreach (string name in new[] { "WorklogAddEntryButton", "WorklogCreateWorkbookButton", "WorklogAddDeviceButton" })
                {
                    Button button = main.GetControl<Button>(name);
                    Assert.True(button.IsVisible, $"{name} is hidden");
                    double right = button.TranslatePoint(new Avalonia.Point(button.Bounds.Width, 0), main)!.Value.X;
                    Assert.True(right <= barRight + 0.5, $"{name} ends at {right}, past the bar's {barRight}, at {windowWidth} wide");
                    Assert.True(button.Bounds.Width >= button.DesiredSize.Width - 0.5, $"{name} is squeezed to {button.Bounds.Width} of {button.DesiredSize.Width}");
                }

                // The drop-downs' place ends before the buttons, and nothing of them is drawn past it.
                Panel pickers = main.GetControl<Panel>("WorklogPickersPanel");
                var rightGroup = Assert.IsAssignableFrom<Control>(main.GetControl<Button>("WorklogAddEntryButton").Parent);
                double pickersRight = pickers.TranslatePoint(new Avalonia.Point(pickers.Bounds.Width, 0), main)!.Value.X;
                double buttonsLeft = rightGroup.TranslatePoint(default, main)!.Value.X;
                Assert.True(pickers.ClipToBounds);
                Assert.True(pickersRight <= buttonsLeft, $"the drop-downs' place ends at {pickersRight}, under the buttons from {buttonsLeft}");

                ComboBox workbooks = WorkbooksOf(main);
                double pickerRight = workbooks.TranslatePoint(new Avalonia.Point(workbooks.Bounds.Width, 0), main)!.Value.X;
                Assert.True(pickerRight <= pickersRight + 0.5, $"the workbook picker is cut off at {windowWidth} wide");
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // THE DROP-DOWNS ARE AS WIDE AS THEIR TEXT (owner request, 2026-11-06: "have the two drop-downs resize, so it
    // will show as much as the text in the drop-downs as possible" - fixed shares, the device's the smallest, until
    // then): with room, each shows its whole line, and the device's widens and narrows with the choice it shows.
    // Wide, as the headless platform's stand-in fonts are wide (the header of the test above).
    // ###########################################################################################
    [Fact]
    public void The_drop_downs_are_as_wide_as_the_text_they_show()
    {
        OwnedDeviceRecord c64 = AddDevice("C64");
        AddRepair("No video and a dead keyboard after the PSU went up in smoke", c64.Id);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.Width = 2600;
            main.Height = 600;
            main.ApplyWorklogBarVisibility();
            main.Show();

            try
            {
                double Layout()
                {
                    Dispatcher.UIThread.RunJobs();
                    main.UpdateLayout();
                    return DevicesOf(main).Bounds.Width;
                }

                // Every line's words drawn whole: as wide as they are unbounded.
                void AssertWhole(ComboBox box)
                {
                    TextBlock words = box.GetVisualDescendants().OfType<TextBlock>().First(text => text.Classes.Contains("BarPickerText"));
                    var natural = new TextBlock { Text = words.Text, FontSize = words.FontSize, FontFamily = words.FontFamily };
                    natural.Measure(Size.Infinity);
                    Assert.True(words.Bounds.Width >= natural.DesiredSize.Width - 0.5, $"[{words.Text}] is cut to {words.Bounds.Width} of {natural.DesiredSize.Width}");
                }

                double allWidth = Layout();
                AssertWhole(DevicesOf(main));
                AssertWhole(WorkbooksOf(main));

                Pick(main, c64);
                double c64Width = Layout();
                AssertWhole(DevicesOf(main));
                AssertWhole(WorkbooksOf(main));

                Assert.True(c64Width < allWidth, $"the device box is {c64Width} for \"C64 (1)\" and {allWidth} for all devices");
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // A device added in "Add workbook" is on disk whether or not the workbook is then made: the form
    // called off, the drop-down lists it all the same (code review, 2026-10-29: hidden until some other
    // refresh came).
    // ###########################################################################################
    [Fact]
    public void A_device_added_in_a_called_off_Add_workbook_is_in_the_drop_down()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = DevicesOf(main);
            Assert.False(picker.IsVisible);

            AddDevice("Breadbin #2");
            main.AfterCreateWorkbookForm(created: null, deviceAdded: true);

            Assert.True(picker.IsVisible);
            Assert.Equal([null, "Breadbin #2"], ChoicesOf(main).Select(choice => choice.Device?.Name));
            Assert.False(main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText").IsVisible);
        });
    }

    // ###########################################################################################
    // THE "Workbook" DROP-DOWN IS NOT THERE UNTIL A WORKBOOK IS, ON ANY BOARD (owner request, 2026-11-08: "if no devices
    // or workbooks exists yet, because the user has not created any yet, then its drop-down boxes should not show
    // either"). One on another board is a workbook made: the drop-down lists it, as it lists every board's.
    // ###########################################################################################
    [Fact]
    public void The_workbook_drop_down_is_not_there_until_a_workbook_is_made_on_any_board()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.ApplyWorklogBarVisibility();

            ComboBox workbooks = WorkbooksOf(main);
            TextBlock heading = main.GetControl<TextBlock>("WorkbookPickerHeadingText");

            Assert.False(workbooks.IsVisible);
            Assert.False(heading.IsVisible);

            AddRepair("No video", deviceId: null, boardKey: OtherBoardKey);
            main.RefreshWorklogBar();

            Assert.True(workbooks.IsVisible);
            Assert.True(heading.IsVisible);
            Assert.False(DevicesOf(main).IsVisible);
        });
    }

    // The bar is as tall with no drop-down as with them: the first workbook or device made must not move the tabs, and
    // the schematic fitted to them, down.
    [Fact]
    public void The_first_drop_down_coming_moves_nothing_below_the_bar()
    {
        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();

            try
            {
                main.ApplyWorklogBarVisibility();
                main.Show();
                Dispatcher.UIThread.RunJobs();

                Control tabs = main.GetControl<Control>("MainTabControl");
                double before = tabs.Bounds.Top;
                Assert.False(WorkbooksOf(main).IsVisible);

                AddRepair("No video", deviceId: null);
                main.RefreshWorklogBar();
                Dispatcher.UIThread.RunJobs();

                Assert.True(WorkbooksOf(main).IsVisible);
                Assert.Equal(before, tabs.Bounds.Top);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // ###########################################################################################
    // An open repair on a board CRT no longer lists cannot be shown: the device is worked on through
    // its repair on the board it has now instead (code review, 2026-10-29: the pick did nothing).
    // ###########################################################################################
    [Fact]
    public void A_device_whose_open_repair_is_on_a_board_gone_is_worked_on_on_its_board()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());
        Assert.False(WorklogManager.IsWorkbookStatusOpen(WorklogManager.GetAllWorkbooks().Single(each => each.Id == closed.Id).Status));
        AddRepair("No video", breadbin.Id, "Nowhere|" + Guid.NewGuid().ToString("N"));
        AddRepair("Dead keyboard", null);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();

            Pick(main, breadbin);

            Assert.Equal(closed.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
        });
    }

    // ###########################################################################################
    // The closed picker takes no pick from the arrow keys (code review, 2026-10-29): each step was a
    // board load. The list is opened to pick from.
    // ###########################################################################################
    [Fact]
    public void The_closed_picker_takes_no_pick_from_the_arrow_keys()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddDevice("C64C silver");
        AddRepair("No video", breadbin.Id);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = DevicesOf(main);
            int picksBefore = main.DevicePicksForTests;

            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Down, Source = picker });
            Dispatcher.UIThread.RunJobs();

            Assert.Null(Assert.IsType<BarDeviceChoice>(picker.SelectedItem).Device);
            Assert.Equal(picksBefore, main.DevicePicksForTests);
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
    // The device the drop-down already shows, chosen again from the open list: its open repair is
    // worked on - the plan's "Choosing a machine selects its board and its open repair" - when the
    // bar holds one of its closed repairs (code review, 2026-10-29: no selection change, no pick).
    // ###########################################################################################
    [Fact]
    public void Choosing_the_device_shown_again_works_on_its_open_repair()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord open = AddRepair("No video", breadbin.Id);
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());
        UserSettings.WorklogBarDeviceId = breadbin.Id;

        UiTest.Run(() =>
        {
            // The bar holds the newest repair - the closed one - and the drop-down Breadbin #2.
            CRT.Main main = BuildMain();
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            ComboBox picker = DevicesOf(main);
            picker.IsDropDownOpen = true;
            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter, Source = picker });
            picker.IsDropDownOpen = false;

            Dispatcher.UIThread.RunJobs();

            Assert.Equal(open.Id, UserSettings.GetActiveWorkbookId(BoardKey));
        });
    }

    // ###########################################################################################
    // The same, by a CLICK on the device's item in the open list - the popup clicked as a person
    // clicks it (code review, 2026-10-29: the list closed before the click was counted, so a click
    // did nothing; only Enter worked).
    // ###########################################################################################
    [Fact]
    public void Clicking_the_device_shown_again_works_on_its_open_repair()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        WorkbookRecord open = AddRepair("No video", breadbin.Id);
        WorkbookRecord closed = AddRepair("PSU replaced", breadbin.Id);
        WorklogManager.AddEntry(closed.Id, "Sheet 1", new Avalonia.Rect(0, 0, 1, 1), "PSU", "", "Issue", "Closed", Array.Empty<string>());
        UserSettings.WorklogBarDeviceId = breadbin.Id;

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            main.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                ComboBox picker = DevicesOf(main);
                picker.IsDropDownOpen = true;
                Dispatcher.UIThread.RunJobs();

                ComboBoxItem item = picker.GetRealizedContainers().OfType<ComboBoxItem>()
                    .Single(container => container.Content is BarDeviceChoice { Device: { } device } && device.Id == breadbin.Id);
                TopLevel popup = TopLevel.GetTopLevel(item)!;
                Point at = item.TranslatePoint(new Point(item.Bounds.Width / 2, item.Bounds.Height / 2), popup)!.Value;

                popup.MouseDown(at, MouseButton.Left);
                popup.MouseUp(at, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

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
    // the drop-downs take as the board already shown - draws the bar again, never naming another repair
    // than the one "Add worklog" writes into (code review, 2026-10-29). The device stays picked: CRT works
    // only in its workbooks (owner decision, 2026-11-06), and none is the board shown's, so none is worked
    // in - "Add worklog" hidden, the workbook on no device not written into unseen.
    // ###########################################################################################
    [Fact]
    public void A_pick_whose_board_switch_lands_nowhere_works_in_no_workbook()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id, BoardKey.ToUpperInvariant());
        AddRepair("Dead keyboard", null);

        try
        {
            ListBoards(BoardKey);

            UiTest.Run(() =>
            {
                CRT.Main main = BuildMain();
                main.ApplyWorklogBarVisibility();

                Pick(main, breadbin);

                Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
                Assert.Null(WorkbooksOf(main).SelectedItem);
                Assert.False(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // ONE pick for one choice: the list opened on Breadbin #2, a refresh moved the drop-down meanwhile
    // (a workbook of another device activated - that device picked), and Enter on Breadbin #2 both changes
    // the selection back and closes the list - it is worked on once, not twice (code review, 2026-10-29).
    // ###########################################################################################
    [Fact]
    public void Enter_after_a_refresh_moved_the_picker_makes_one_pick()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        OwnedDeviceRecord silver = AddDevice("C64C silver");
        AddRepair("No video", breadbin.Id);
        WorkbookRecord silverRepair = AddRepair("Dead keyboard", silver.Id);
        AddRepair("PSU replaced", breadbin.Id);
        UserSettings.WorklogBarDeviceId = breadbin.Id;

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();
            ComboBox picker = DevicesOf(main);
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            picker.IsDropDownOpen = true;

            // Silver's workbook activated: the pick follows it (OwnedDeviceRepairs.PickFollowing).
            main.ActivateWorkbook(BoardKey, silverRepair.Id);
            Assert.Equal(silver.Id, main.DeviceWorkedOn!.Id);

            int picksBefore = main.DevicePicksForTests;
            picker.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter, Source = picker });
            picker.SelectedItem = ChoicesOf(main).Single(choice => choice.Device?.Id == breadbin.Id);
            picker.IsDropDownOpen = false;

            Dispatcher.UIThread.RunJobs();

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
    // "Add workbook" in the bar preselects the device picked, else the one the bar's workbook is on -
    // only on its own board.
    // ###########################################################################################
    [Fact]
    public void Add_workbook_in_the_bar_preselects_the_device_worked_on()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2");
        AddRepair("No video", breadbin.Id);

        UiTest.Run(() =>
        {
            CRT.Main main = BuildMain();

            // All devices shown, the bar's workbook on Breadbin #2.
            Assert.Null(main.DeviceWorkedOn);
            Assert.Equal(breadbin.Id, main.DeviceForNewWorkbook(BoardKey));
            Assert.Null(main.DeviceForNewWorkbook(OtherBoardKey));

            AddRepair("Dead keyboard", null);
            main.RefreshWorklogBar();

            Assert.Null(main.DeviceForNewWorkbook(BoardKey));

            Pick(main, breadbin);
            Assert.Equal(breadbin.Id, main.DeviceForNewWorkbook(BoardKey));
        });
    }
}
