using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.LogicalTree;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "MY DEVICES" STANDS IN THE LEFT PANEL WHILE THE "My Devices" TAB IS SHOWN, THE WORKBOOKS IN THE TAB (owner requests,
// 2026-11-06: the lists "in the left-side panel, where the component category and component list is", then "please
// again have [the workbooks] visible where they were previously, but do keep the 'My devices' in the left-side panel")
// - Main.DeviceList.cs, TabWorkbooks.ShowDevicesIn; and what both list, the Configuration tab's scope (MyDevicesScope).
// Through the real window: what is shown where is its layout.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainDeviceListTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public MainDeviceListTests()
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

    private static OwnedDeviceRecord AddDevice(string name, string boardKey) =>
        OwnedDeviceStore.Create(new OwnedDeviceDetails(name, boardKey, "PAL", string.Empty, string.Empty, OwnedDeviceStatus.Working))!;

    private static void OnTab(CRT.Main main, string tabItem)
    {
        main.GetControl<TabControl>("MainTabControl").SelectedItem = main.GetControl<TabItem>(tabItem);
        Dispatcher.UIThread.RunJobs();
        main.UpdateLayout();
    }

    private static void WithMain(Action<CRT.Main> body)
    {
        UiTest.Run(() =>
        {
            var main = new CRT.Main { BoardKeyOverrideForTests = BoardKey, Width = 1400, Height = 900 };
            main.ApplyWorklogBarVisibility();
            main.Show();

            try
            {
                Dispatcher.UIThread.RunJobs();
                body(main);
            }
            finally
            {
                main.IsApplicationShuttingDownForTests = true;
                main.Close();
            }
        });
    }

    // The device cards' names, and the workbook cards' titles, as listed.
    private static List<string> DeviceNames(CRT.Main main) =>
        main.TabWorkbooks.DeviceCardsForTests
            .SelectMany(card => card.GetLogicalDescendants().OfType<TextBlock>().Take(1))
            .Select(name => name.Text ?? string.Empty)
            .ToList();

    private static List<int> WorkbookIds(CRT.Main main) =>
        main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children
            .Select(card => Assert.IsType<WorkbookRecord>(card.Tag).Id)
            .ToList();

    // ###########################################################################################
    // "My devices"' frame is the workbook list's line, and follows the theme as that one does (code review, 2026-11-07: its
    // color was copied once, as Main was built, and stayed the old theme's after a switch). The theme is switched on the
    // window - never on the shared Application (UiTest).
    // ###########################################################################################
    [Fact]
    public void The_device_lists_frame_follows_the_theme()
    {
        WithMain(main =>
        {
            Color ColorIn(Avalonia.Styling.ThemeVariant variant)
            {
                Assert.True(Application.Current!.TryGetResource("Workbooks_Separator", variant, out object? brush));
                return Assert.IsAssignableFrom<ISolidColorBrush>(brush).Color;
            }

            Color light = ColorIn(Avalonia.Styling.ThemeVariant.Light);
            Color dark = ColorIn(Avalonia.Styling.ThemeVariant.Dark);
            Assert.NotEqual(light, dark);

            Border frame = main.TabWorkbooks.GetControl<Border>("DeviceListBorder");
            Color Shown() => Assert.IsAssignableFrom<ISolidColorBrush>(frame.BorderBrush).Color;

            main.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Light;
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(light, Shown());

            main.RequestedThemeVariant = Avalonia.Styling.ThemeVariant.Dark;
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(dark, Shown());
        });
    }

    // ###########################################################################################
    // On the tab "My devices" stands in the left panel where the category and component lists were, those lists and
    // the buttons under them gone - Hardware and Board stay - with no "Add device" of its own (owner request: "no need
    // for it"; the bar has one). The workbooks are in the tab's own column, with its splitter. Every other tab has the
    // lists back, the Contribute tab its own.
    // ###########################################################################################
    [Fact]
    public void On_the_My_Devices_tab_my_devices_takes_the_component_lists_place_and_the_workbooks_stay_in_the_tab()
    {
        WithMain(main =>
        {
            Border devices = main.TabWorkbooks.GetControl<Border>("DeviceListBorder");
            Border workbooks = main.TabWorkbooks.GetControl<Border>("WorkbookListBorder");
            Border leftPanel = main.GetControl<Border>("LeftPanel");
            Control host = main.GetControl<Control>("DeviceListHost");
            Control components = main.GetControl<Control>("ComponentListsGrid");
            Control actions = main.GetControl<Control>("ComponentActionsPanel");

            // Handed over at construction, whatever tab is shown; the workbooks never are.
            Assert.Contains(leftPanel, devices.GetLogicalAncestors());
            Assert.Contains(main.TabWorkbooks, workbooks.GetLogicalAncestors());
            Assert.DoesNotContain(leftPanel, workbooks.GetLogicalAncestors());
            Assert.DoesNotContain(devices.GetLogicalDescendants().OfType<Button>(), button => button.Content as string == "Add device");
            Assert.False(main.TabWorkbooks.GetControl<Border>("DeviceListSeparator").IsVisible);

            OnTab(main, "SchematicsTabItem");
            Assert.False(host.IsVisible);
            Assert.True(components.IsVisible);
            Assert.True(actions.IsVisible);

            OnTab(main, "WorkbooksTabItem");
            Assert.True(devices.IsEffectivelyVisible);
            Assert.False(components.IsVisible);
            Assert.False(actions.IsVisible);
            Assert.True(main.GetControl<ComboBox>("BoardComboBox").IsEffectivelyVisible);
            Assert.True(workbooks.IsEffectivelyVisible);
            Assert.True(workbooks.Bounds.Width >= 199, $"the workbook column is {workbooks.Bounds.Width} wide");
            Assert.True(main.TabWorkbooks.GetControl<GridSplitter>("OuterSplitter").IsEffectivelyVisible);

            // "My devices" fills the panel below Board.
            ComboBox board = main.GetControl<ComboBox>("BoardComboBox");
            double boardBottom = board.TranslatePoint(new Point(0, board.Bounds.Height), main)!.Value.Y;
            Assert.True(devices.TranslatePoint(default, main)!.Value.Y >= boardBottom);
            Assert.True(devices.Bounds.Height > 300, $"the list is only {devices.Bounds.Height} high");

            OnTab(main, "ContributeTabItem");
            Assert.False(host.IsVisible);
            Assert.True(components.IsVisible);
            Assert.True(main.GetControl<ListBox>("ContributeComponentListBox").IsVisible);
        });
    }

    // ###########################################################################################
    // The device cards keep their look in the left panel - the tab's styles do not reach there, the list's own do -
    // and the workbooks' search marks what it found on the cards in the tab, as it always did.
    // ###########################################################################################
    [Fact]
    public void The_cards_keep_their_look_and_the_search_still_marks_the_workbooks()
    {
        AddDevice("Breadbin #2", BoardKey);
        WorklogManager.CreateWorkbook(BoardKey, "Dead keyboard", "");
        WorklogManager.CreateWorkbook(BoardKey, "No video", "");

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            Border deviceCard = Assert.Single(main.TabWorkbooks.DeviceCardsForTests);
            Assert.Contains("WorkbooksJobCard", deviceCard.Classes);
            Assert.Equal(new Thickness(7), deviceCard.Padding);

            main.TabWorkbooks.GetControl<TextBox>("FindRepairTextBox").Text = "video";
            main.TabWorkbooks.RefreshWorkbooks();
            Dispatcher.UIThread.RunJobs();
            main.UpdateLayout();

            Control only = Assert.Single(main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children);
            List<string> marked = only.GetSelfAndVisualDescendants()
                .OfType<TextBlock>()
                .Where(text => text.Inlines != null)
                .SelectMany(text => text.Inlines!.OfType<Run>())
                .Where(run => run.Background != null)
                .Select(run => run.Text ?? string.Empty)
                .ToList();
            Assert.Contains("video", marked);
        });
    }

    // ###########################################################################################
    // THE SCOPE NARROWS BOTH LISTS ALIKE (owner request, 2026-11-06: "Show all devices", "Show only devices for selected
    // hardware" in the middle, "Show only devices for selected board"; asked, "Follow the same scope" for the
    // workbooks): every device and workbook; those of any board of the hardware shown; those of the board shown -
    // and with devices, none in the scope, the list says so in its words.
    // ###########################################################################################
    [Theory]
    [InlineData(MyDevicesScope.AllBoards, new[] { "C128D", "Silver", "Breadbin #2" }, 3)]
    [InlineData(MyDevicesScope.CurrentHardware, new[] { "Silver", "Breadbin #2" }, 2)]
    [InlineData(MyDevicesScope.CurrentBoard, new[] { "Breadbin #2" }, 1)]
    public void The_scope_narrows_the_devices_and_the_workbooks_alike(string scope, string[] devices, int workbooks)
    {
        AddDevice("Breadbin #2", BoardKey);
        AddDevice("Silver", "Commodore 64|250469");
        AddDevice("C128D", "Commodore 128|310378");
        WorklogManager.CreateWorkbook(BoardKey, "No video", "");
        WorklogManager.CreateWorkbook("Commodore 64|250469", "Dead keyboard", "");
        WorklogManager.CreateWorkbook("Commodore 128|310378", "No sound", "");
        UserSettings.WorkbooksScope = scope;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            Assert.Equal(devices.OrderBy(name => name), DeviceNames(main).OrderBy(name => name));
            Assert.Equal(workbooks, WorkbookIds(main).Count);
            Assert.False(main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText").IsVisible);
        });
    }

    // A press and release at the middle of `control`, as a person clicks it - laid out first, as on screen: a control
    // made by a refresh has no place until then.
    private static void Click(CRT.Main main, Control control)
    {
        Dispatcher.UIThread.RunJobs();
        main.UpdateLayout();

        Point at = control.TranslatePoint(new Point(control.Bounds.Width / 2, control.Bounds.Height / 2), main)!.Value;
        main.MouseDown(at, Avalonia.Input.MouseButton.Left);
        main.MouseUp(at, Avalonia.Input.MouseButton.Left);
        Dispatcher.UIThread.RunJobs();
        main.UpdateLayout();
    }

    private static Border CardOf(CRT.Main main, string name) =>
        main.TabWorkbooks.DeviceCardsForTests.OfType<Border>()
            .Single(card => card.Tag is OwnedDeviceRecord device && device.Name == name);

    private static bool IsMarked(Control card) => card.Classes.Contains("Selected");

    // ###########################################################################################
    // "My devices" is a headline over its panel, as "Hardware" and "Board" are over theirs - the same size, not bold,
    // outside the frame (owner request, 2026-11-06: "make it a headline for the panel below (and it should not be
    // bold)") - and "Clear" stands at the foot of the frame, in sight however long the list is ("should always be
    // visible even if scrolling down in list").
    // ###########################################################################################
    [Fact]
    public void My_devices_is_a_headline_over_its_panel_and_Clear_stays_in_sight_at_its_foot()
    {
        for (int index = 0; index < 30; index++)
        {
            AddDevice($"Device {index:00}", BoardKey);
        }

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            TextBlock heading = main.TabWorkbooks.GetControl<TextBlock>("DeviceListHeadingText");
            TextBlock hardware = main.GetVisualDescendants().OfType<TextBlock>().First(text => text.Text == "Hardware");
            Border frame = main.TabWorkbooks.GetControl<Border>("DeviceListBorder");

            Assert.Equal("My devices", heading.Text);
            Assert.Equal(FontWeight.Normal, heading.FontWeight);
            Assert.Equal(hardware.FontSize, heading.FontSize);
            Assert.DoesNotContain(frame, heading.GetVisualAncestors());
            Assert.True(heading.TranslatePoint(new Point(0, heading.Bounds.Height), main)!.Value.Y <= frame.TranslatePoint(default, main)!.Value.Y);

            // "Clear" UNDER the frame since 2026-11-07, level with the workbook list's own "Clear" (owner request: "move
            // the 'Clear' button outside of the 'My devices' panel and show it below the panel, and then vertically
            // align the two buttons") - and in sight however the list is scrolled.
            Button clear = main.TabWorkbooks.GetControl<Button>("ClearDeviceButton");
            Button clearWorkbook = main.TabWorkbooks.GetControl<Button>("ClearWorkbookButton");
            ScrollViewer list = main.TabWorkbooks.GetControl<ScrollViewer>("DeviceListScroller");
            Assert.DoesNotContain(frame, clear.GetVisualAncestors());
            Assert.True(clear.IsEffectivelyVisible);
            Assert.True(list.Extent.Height > list.Viewport.Height, "thirty devices do not fill the list");

            double frameBottom = frame.TranslatePoint(new Point(0, frame.Bounds.Height), main)!.Value.Y;
            double clearTop = clear.TranslatePoint(default, main)!.Value.Y;
            double clearBottom = clear.TranslatePoint(new Point(0, clear.Bounds.Height), main)!.Value.Y;
            Assert.True(clearTop >= frameBottom - 0.5, $"\"Clear\" starts at {clearTop}, inside the frame ending at {frameBottom}");
            Assert.Equal(clearBottom, clearWorkbook.TranslatePoint(new Point(0, clearWorkbook.Bounds.Height), main)!.Value.Y, 1);
            Assert.Equal(clear.Bounds.Height, clearWorkbook.Bounds.Height, 1);

            list.Offset = new Vector(0, list.Extent.Height);
            main.UpdateLayout();
            Assert.Equal(clearBottom, clear.TranslatePoint(new Point(0, clear.Bounds.Height), main)!.Value.Y, 1);
        });
    }

    // ###########################################################################################
    // A DEVICE CLICKED IN "My devices" IS THE DEVICE PICKED, the bar's "Device" drop-down with it (owner request,
    // 2026-11-06: "When in 'My Device' and I have selected a device, then I do expect that it will only show the
    // workbooks for this device in the 2 workbooks lists"): the tab's workbooks and the bar's are its own, its card
    // marked, its page shown. "Clear" lets go - every workbook again, nothing marked, the workbook back - and is
    // greyed out with nothing to let go of.
    // ###########################################################################################
    [Fact]
    public void A_device_clicked_is_picked_in_the_bar_too_and_Clear_shows_every_workbook_again()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2", BoardKey);
        AddDevice("Silver", BoardKey);
        WorkbookRecord own = WorklogManager.CreateWorkbook(BoardKey, "No video", "", breadbin.Id)!;
        WorkbookRecord other = WorklogManager.CreateWorkbook(BoardKey, "Dead keyboard", "")!;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");
            Button clear = main.TabWorkbooks.GetControl<Button>("ClearDeviceButton");
            Assert.False(clear.IsEnabled);
            Assert.Equal([other.Id, own.Id], WorkbookIds(main));

            Click(main, CardOf(main, "Breadbin #2"));

            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
            Assert.Equal(breadbin.Id, UserSettings.WorklogBarDeviceId);
            Assert.Equal([own.Id], WorkbookIds(main));
            Assert.Equal([own.Id], main.GetControl<ComboBox>("WorklogJobBox").Items.OfType<WorkbookRecord>().Select(workbook => workbook.Id));
            Assert.True(IsMarked(CardOf(main, "Breadbin #2")));
            Assert.False(IsMarked(CardOf(main, "Silver")));
            Assert.True(main.TabWorkbooks.IsDevicePageShown);
            Assert.True(clear.IsEnabled);

            Click(main, clear);

            Assert.Null(main.DeviceWorkedOn);
            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.Equal([other.Id, own.Id], WorkbookIds(main));
            Assert.False(IsMarked(CardOf(main, "Breadbin #2")));
            Assert.False(main.TabWorkbooks.IsDevicePageShown);
            Assert.False(clear.IsEnabled);
        });
    }

    // ###########################################################################################
    // A WORKBOOK CLICKED PICKS THE DEVICE IT IS ON (owner request, 2026-11-06: "when I in the 'Workbooks' list selects a
    // workbook associated to a device, then I want that device selected in both the drop-down and in the 'My devices'
    // panel"); one on no device, all devices.
    // ###########################################################################################
    [Fact]
    public void A_workbook_clicked_picks_its_device_in_the_bar_and_in_My_devices()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2", BoardKey);
        WorkbookRecord own = WorklogManager.CreateWorkbook(BoardKey, "No video", "", breadbin.Id)!;
        WorkbookRecord onNone = WorklogManager.CreateWorkbook(BoardKey, "Dead keyboard", "")!;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");
            Assert.Null(main.DeviceWorkedOn);

            Control WorkbookCard(int id) => main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children
                .Single(card => card.Tag is WorkbookRecord workbook && workbook.Id == id);

            Click(main, WorkbookCard(own.Id));

            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
            Assert.True(IsMarked(CardOf(main, "Breadbin #2")));
            Assert.Equal(own.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal([own.Id], WorkbookIds(main));

            main.TabWorkbooks.ClearDevicePick();
            Click(main, WorkbookCard(onNone.Id));

            Assert.Null(main.DeviceWorkedOn);
            Assert.False(IsMarked(CardOf(main, "Breadbin #2")));
            Assert.Equal(onNone.Id, UserSettings.GetActiveWorkbookId(BoardKey));
        });
    }

    // ###########################################################################################
    // ANOTHER DEVICE CLICKED WHILE ONE IS PICKED PICKS THAT ONE (owner report, 2026-11-06: "when I then select another
    // device in the list, then it deselects and actually does the exact same thing as 'Clear'") - whether it has
    // workbooks or none.
    // ###########################################################################################
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Another_device_clicked_while_one_is_picked_picks_that_one(bool secondHasWorkbooks)
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2", BoardKey);
        OwnedDeviceRecord silver = AddDevice("Silver", BoardKey);
        WorklogManager.CreateWorkbook(BoardKey, "No video", "", breadbin.Id);
        WorkbookRecord? silverRepair = secondHasWorkbooks ? WorklogManager.CreateWorkbook(BoardKey, "Dead SID", "", silver.Id) : null;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            Click(main, CardOf(main, "Breadbin #2"));
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);

            Click(main, CardOf(main, "Silver"));

            Assert.Equal(silver.Id, main.DeviceWorkedOn?.Id);
            Assert.Equal(silver.Id, UserSettings.WorklogBarDeviceId);
            Assert.True(IsMarked(CardOf(main, "Silver")));
            Assert.False(IsMarked(CardOf(main, "Breadbin #2")));
            Assert.True(main.TabWorkbooks.IsDevicePageShown);
            Assert.Equal(silver.Id, main.TabWorkbooks.SelectedDeviceIdForTests);
            Assert.Equal(silverRepair is null ? [] : [silverRepair.Id], WorkbookIds(main));
        });
    }

    // ###########################################################################################
    // ... AND ONE ON ANOTHER HARDWARE TOO, whose board CRT switches to (the cause of the report): switching hardware
    // replaces the boards' list first, so for a moment NO board is shown, and the worklog bar drawn then took the
    // device for one of "another board" and let go of it - the very thing "Clear" does. Through the real drop-downs,
    // as StartAsync starts them.
    // ###########################################################################################
    [Fact]
    public async Task A_device_on_another_hardware_clicked_stays_picked_through_the_board_switch()
    {
        const string C128 = "Commodore 128|310378";
        const string C64 = "Commodore 64|250407";

        string dataRoot = this.thisWorkspace.Path_("Data");
        DataManager.LoadFrom(dataRoot, MainFiles.Write(
            dataRoot,
            new MainFiles.Board("Commodore/C128/310378", "Commodore 128", "310378"),
            new MainFiles.Board("Commodore/C64/250407", "Commodore 64", "250407")));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
        UserSettings.SetLastHardware("Commodore 128");
        UserSettings.WorkbooksScope = MyDevicesScope.AllBoards;

        OwnedDeviceRecord c128 = AddDevice("C128", C128);
        OwnedDeviceRecord c64 = AddDevice("C64 Breadbin", C64);

        try
        {
            await UiTest.RunAsync(async () =>
            {
                var main = new CRT.Main { Width = 1400, Height = 900 };
                main.ApplyWorklogBarVisibility();
                main.Show();

                try
                {
                    main.PopulateHardwareDropDown();
                    await Until(() => main.GetCurrentBoardKey() == C128, () => $"the board shown is [{main.GetCurrentBoardKey()}]");
                    OnTab(main, "WorkbooksTabItem");

                    main.TabWorkbooks.SelectDevice(c128.Id);
                    Assert.Equal(c128.Id, main.DeviceWorkedOn!.Id);

                    main.TabWorkbooks.SelectDevice(c64.Id);
                    await Until(() => main.GetCurrentBoardKey() == C64, () => $"the board shown is [{main.GetCurrentBoardKey()}]");
                    await Until(() => main.DeviceWorkedOn is not null, () => "no device is picked");

                    Assert.Equal(c64.Id, main.DeviceWorkedOn!.Id);
                    Assert.Equal(c64.Id, UserSettings.WorklogBarDeviceId);
                    Assert.True(main.TabWorkbooks.IsDevicePageShown);
                    Assert.Equal(c64.Id, main.TabWorkbooks.SelectedDeviceIdForTests);
                }
                finally
                {
                    main.IsApplicationShuttingDownForTests = true;
                    main.Close();
                }
            });
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
            DraftManager.LoadFrom(string.Empty);
        }
    }

    private static async Task Until(Func<bool> condition, Func<string> what)
    {
        for (int attempt = 0; attempt < 500 && !condition(); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.True(condition(), what());
    }

    // ###########################################################################################
    // The device picked clicked again shows its page and nothing more: the workbook of it worked in since is kept,
    // not swapped for the repair a pick works on (its newest open one).
    // ###########################################################################################
    [Fact]
    public void The_device_picked_clicked_again_shows_its_page_and_keeps_the_workbook_worked_in()
    {
        OwnedDeviceRecord breadbin = AddDevice("Breadbin #2", BoardKey);
        WorkbookRecord older = WorklogManager.CreateWorkbook(BoardKey, "No video", "", breadbin.Id)!;
        WorkbookRecord newest = WorklogManager.CreateWorkbook(BoardKey, "No sound", "", breadbin.Id)!;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            Click(main, CardOf(main, "Breadbin #2"));
            Assert.Equal(newest.Id, UserSettings.GetActiveWorkbookId(BoardKey));

            Click(main, main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children
                .Single(card => card.Tag is WorkbookRecord workbook && workbook.Id == older.Id));
            Assert.Equal(older.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.False(main.TabWorkbooks.IsDevicePageShown);

            Click(main, CardOf(main, "Breadbin #2"));

            Assert.True(main.TabWorkbooks.IsDevicePageShown);
            Assert.Equal(older.Id, UserSettings.GetActiveWorkbookId(BoardKey));
            Assert.Equal(breadbin.Id, main.DeviceWorkedOn!.Id);
        });
    }

    // ###########################################################################################
    // A device's card (owner requests, 2026-11-06): its name and status, then its serial number or unique identifier
    // with how many workbooks it has at that line's end - an outlined badge, as the bar's "Device" drop-down counts
    // them ("the workbook counter is written in the second line, and not after the status") - then its hardware and
    // board.
    // ###########################################################################################
    [Fact]
    public void A_device_card_has_its_status_then_its_serial_and_count_then_its_board()
    {
        OwnedDeviceRecord breadbin = OwnedDeviceStore.Create(new OwnedDeviceDetails("Breadbin #2", BoardKey, "PAL", "CA123456", string.Empty, OwnedDeviceStatus.Working))!;
        AddDevice("Silver", BoardKey);
        WorklogManager.CreateWorkbook(BoardKey, "No video", "", breadbin.Id);
        WorklogManager.CreateWorkbook("Commodore 64|250425", "PSU", "", breadbin.Id);

        WithMain(main =>
        {
            main.RefreshWorklogBar();

            // The card's rows: its lines on the left, its badges on the right, a row each (2026-11-07).
            Grid RowsOf(string name) => (Grid)CardOf(main, name).Child!;

            string TextIn(Control control) =>
                control.GetSelfAndLogicalDescendants().OfType<TextBlock>().Select(text => text.Text ?? string.Empty).LastOrDefault() ?? string.Empty;

            Control At(Grid rows, int row, int column) => rows.Children.Single(child => Grid.GetRow(child) == row && Grid.GetColumn(child) == column);

            Grid rows = RowsOf("Breadbin #2");
            Assert.Equal(3, rows.RowDefinitions.Count);
            Assert.Equal("Working", TextIn(At(rows, 0, 1)));
            Assert.Equal("CA123456", TextIn(At(rows, 1, 0)));
            Assert.Equal("2", TextIn(At(rows, 1, 1)));
            Assert.Equal("Commodore 64 - 250407", TextIn(At(rows, 2, 0)));

            // No serial: the board on the second line, beside the count - no blank line.
            Grid silver = RowsOf("Silver");
            Assert.Equal(2, silver.RowDefinitions.Count);
            Assert.Equal("Commodore 64 - 250407", TextIn(At(silver, 1, 0)));
            Assert.Equal("0", TextIn(At(silver, 1, 1)));
        });
    }

    // ###########################################################################################
    // "Find a previous repair" searches the devices too (owner request, 2026-11-06: "searching a serial/unique in the
    // filter in 'My Devices' should filter and highlight info in this list also"): only the devices it finds are
    // listed, what it found marked on their cards; finding none says so.
    // ###########################################################################################
    [Fact]
    public void The_search_lists_only_the_devices_it_finds_and_marks_what_it_found()
    {
        OwnedDeviceStore.Create(new OwnedDeviceDetails("Breadbin #2", BoardKey, "PAL", "CA123456", string.Empty, OwnedDeviceStatus.Working));
        OwnedDeviceStore.Create(new OwnedDeviceDetails("Silver", BoardKey, "PAL", "SV998877", string.Empty, OwnedDeviceStatus.Working));

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");
            TextBox search = main.TabWorkbooks.GetControl<TextBox>("FindRepairTextBox");

            search.Text = "ca1234";
            main.TabWorkbooks.RefreshWorkbooks();
            main.UpdateLayout();

            Border only = Assert.Single(main.TabWorkbooks.DeviceCardsForTests);
            Assert.Equal("Breadbin #2", Assert.IsType<OwnedDeviceRecord>(only.Tag).Name);
            List<string> marked = only.GetLogicalDescendants()
                .OfType<TextBlock>()
                .Where(text => text.Inlines != null)
                .SelectMany(text => text.Inlines!.OfType<Run>())
                .Where(run => run.Background != null)
                .Select(run => run.Text ?? string.Empty)
                .ToList();
            Assert.Equal(["CA1234"], marked);

            search.Text = "no such device";
            main.TabWorkbooks.RefreshWorkbooks();

            Assert.Empty(main.TabWorkbooks.DeviceCardsForTests);
            Assert.Equal(OwnedDeviceWording.NoDevicesMatch, main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText").Text);
            Assert.True(main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText").IsVisible);

            search.Text = string.Empty;
            main.TabWorkbooks.RefreshWorkbooks();
            Assert.Equal(2, main.TabWorkbooks.DeviceCardsForTests.Count);
        });
    }

    // ###########################################################################################
    // "CLEAR" UNDER THE WORKBOOK LIST (owner request, 2026-11-06; asked, "No workbook worked in"): no workbook card
    // marked, none worked in on the board - the bar's "Workbook" drop-down empty, saying so, "Add worklog" hidden -
    // remembered, and greyed out with nothing to let go of. A workbook clicked is worked in again.
    // ###########################################################################################
    [Fact]
    public void Clear_under_the_workbooks_leaves_no_workbook_worked_in_till_one_is_picked()
    {
        WorkbookRecord older = WorklogManager.CreateWorkbook(BoardKey, "No video", "")!;
        WorkbookRecord newest = WorklogManager.CreateWorkbook(BoardKey, "Dead keyboard", "")!;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");
            Button clear = main.TabWorkbooks.GetControl<Button>("ClearWorkbookButton");
            ComboBox bar = main.GetControl<ComboBox>("WorklogJobBox");
            Assert.True(clear.IsEnabled);
            Assert.Equal(newest.Id, Assert.IsType<WorkbookRecord>(bar.SelectedItem).Id);

            Click(main, clear);

            Assert.Null(bar.SelectedItem);
            Assert.Equal(CRT.Main.NoWorkbookPickedText, bar.PlaceholderText);
            Assert.True(bar.IsEnabled);
            Assert.False(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);
            Assert.Null(CRT.Main.ResolveActiveWorkbookForBoard(BoardKey));
            Assert.DoesNotContain(main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children, card => IsMarked(card));
            Assert.Equal("No workbook selected", main.TabWorkbooks.GetControl<TextBlock>("WorkbookHeaderTitleText").Text);
            Assert.False(clear.IsEnabled);

            // Remembered: drawn again, still none.
            main.RefreshWorklogBar();
            Assert.Null(bar.SelectedItem);

            Click(main, main.TabWorkbooks.GetControl<StackPanel>("WorkbookListPanel").Children
                .Single(card => card.Tag is WorkbookRecord workbook && workbook.Id == older.Id));

            Assert.Equal(older.Id, Assert.IsType<WorkbookRecord>(bar.SelectedItem).Id);
            Assert.True(main.GetControl<Button>("WorklogAddEntryButton").IsVisible);
            Assert.True(clear.IsEnabled);
        });
    }

    [Fact]
    public void With_devices_but_none_in_the_scope_the_list_says_so()
    {
        AddDevice("C128D", "Commodore 128|310378");
        UserSettings.WorkbooksScope = MyDevicesScope.CurrentHardware;

        WithMain(main =>
        {
            main.RefreshWorklogBar();
            OnTab(main, "WorkbooksTabItem");

            TextBlock none = main.TabWorkbooks.GetControl<TextBlock>("NoDevicesText");
            Assert.Empty(DeviceNames(main));
            Assert.True(none.IsVisible);
            Assert.Equal("No devices for the selected hardware.", none.Text);
        });
    }
}
