using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;
using Handlers.Theming;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// "My devices" on the Workbooks tab (CRT 4.0 step 4, TabWorkbooks.Devices.cs): the user's own
// computers above the workbooks, a device's page - the workbooks associated to it with their totals,
// its notes and photos - in place of the workbook, and "On my device" under a workbook's title. Also
// the two forms it opens: "Add device" and "Create new workbook" with its device.
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

    // The device picked is a setting (UserSettings.WorklogBarDeviceId) a card clicked writes - none at the
    // start of a test and none left behind: device ids start at 1 in every test's own folder, so a pick left
    // over would narrow the next test's workbooks to its device #1.
    public WorkbooksDevicesTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbooks-" + Guid.NewGuid().ToString("N")));
        UserSettings.WorklogBarDeviceId = null;
    }

    public void Dispose()
    {
        UserSettings.WorklogBarDeviceId = null;
        UserSettings.DeviceSorting = OwnedDeviceSorting.ByName;
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
        OwnedDeviceStore.Create(new OwnedDeviceDetails(name, boardKey ?? this.thisBoardKey, "PAL", "CA123456", "Flea market", OwnedDeviceStatus.Working))!;

    // ###########################################################################################
    // "MY DEVICES" COUNTS WHAT IT LISTS, IN A BADGE AFTER ITS HEADLINE (owner request, 2026-11-07: 'show a device
    // counter after "My devices (X)" headline, so I can see the amount of devices in the list - show as badge') - the
    // devices listed, so a search narrows the count with the list.
    // ###########################################################################################
    [Fact]
    public void The_list_counts_its_devices_in_a_badge_after_its_headline()
    {
        this.AddDevice("Breadbin #2");
        this.AddDevice("C64C silver");

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();

            Border badge = Assert.IsType<Border>(tab.GetControl<Decorator>("DeviceCountHost").Child);
            Assert.Equal(["2"], Texts(badge));

            tab.GetControl<TextBox>("FindRepairTextBox").Text = "silver";
            tab.RefreshWorkbooks();

            Assert.Equal(["1"], Texts(Assert.IsType<Border>(tab.GetControl<Decorator>("DeviceCountHost").Child)));
        });
    }

    // ###########################################################################################
    // A CARD'S LINES AND BADGES (owner request, 2026-11-07: "if the device has no serial/unique, then it should not show
    // a blank line - then instead move up the hardware/board ... on the 3rd line, to the right as the status and the
    // counter, show the stars (if any)"): name, serial, board down the left; status, count, stars down the right.
    // ###########################################################################################
    private static Grid LinesOf(DeviceListCard card) =>
        card.Child is Grid { ColumnDefinitions.Count: 2 } grid && grid.Children[0] is Grid body && body.RowDefinitions.Count > 0
            ? body
            : (Grid)card.Child!;

    [Fact]
    public void A_card_without_a_serial_has_no_blank_line_and_its_stars_stand_third_on_the_right()
    {
        OwnedDeviceRecord full = this.AddDevice("Breadbin #2");
        OwnedDeviceStore.Update(full.Id, new OwnedDeviceDetails("Breadbin #2", this.thisBoardKey, "PAL", "CA123456", "", OwnedDeviceStatus.Working, 3));
        OwnedDeviceRecord bare = OwnedDeviceStore.Create(new OwnedDeviceDetails("C64C silver", this.thisBoardKey, "", "", "", OwnedDeviceStatus.Faulty))!;
        OwnedDeviceRecord starred = OwnedDeviceStore.Create(new OwnedDeviceDetails("SX-64", this.thisBoardKey, "", "", "", OwnedDeviceStatus.Working, 2))!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            DeviceListCard Card(OwnedDeviceRecord device) => tab.DeviceCardsForTests.Single(card => card.Device.Id == device.Id);

            // Serial and stars: three lines, the stars the third badge.
            Grid lines = LinesOf(Card(full));
            Assert.Equal(3, lines.RowDefinitions.Count);
            Assert.Equal(["Breadbin #2", "CA123456"], lines.Children.OfType<TextBlock>().Where(text => Grid.GetColumn(text) == 0).Take(2).Select(text => text.Text));
            StackPanel stars = lines.Children.OfType<StackPanel>().Single();
            Assert.Equal((2, 1), (Grid.GetRow(stars), Grid.GetColumn(stars)));

            // No serial, no stars: two lines, the board moved up beside the count.
            Grid twoLines = LinesOf(Card(bare));
            Assert.Equal(2, twoLines.RowDefinitions.Count);
            Assert.DoesNotContain(twoLines.Children.OfType<TextBlock>(), text => Grid.GetColumn(text) == 0 && string.IsNullOrEmpty(text.Text));
            Assert.Empty(twoLines.Children.OfType<StackPanel>());

            // No serial, with stars: the stars keep their place, third on the right.
            Grid withStars = LinesOf(Card(starred));
            Assert.Equal(2, Grid.GetRow(withStars.Children.OfType<StackPanel>().Single()));
            Assert.Equal(2, withStars.Children.OfType<TextBlock>().Count(text => Grid.GetColumn(text) == 0));
        });
    }

    // ###########################################################################################
    // SORTED BY HAND (owner request, 2026-11-07: "When 'Sort manually in list' then when I am in 'My Devices' tab, then
    // I should be able to drag each device up/down in the list"): each card has a handle, a card dragged by it is
    // dropped where the slot shows and the order is kept - and the cards follow every "Sorting of devices" choice.
    // ###########################################################################################
    [Fact]
    public void The_list_follows_the_sorting_chosen_and_sorted_by_hand_a_card_dragged_keeps_its_place()
    {
        this.AddDevice("Breadbin #2");
        this.AddDevice("Amiga 500", "Amiga 500|Rev 6A");
        this.AddDevice("C64C silver");

        UiTest.Run(() =>
        {
            UserSettings.WorkbooksScope = MyDevicesScope.AllBoards;

            try
            {
                TabWorkbooks tab = this.BuildTab();
                var window = new Window { Content = tab, Width = 1200, Height = 900 };
                window.Show();

                List<string> Names() => tab.DeviceCardsForTests.Select(card => card.Device.Name).ToList();
                bool HasHandles() => tab.DeviceCardsForTests.All(card => card.GetLogicalDescendants().OfType<Border>().Any(border => border.Classes.Contains("DeviceDragHandle")));

                Assert.Equal(["Amiga 500", "Breadbin #2", "C64C silver"], Names());
                Assert.DoesNotContain(tab.DeviceCardsForTests, card => card.GetLogicalDescendants().OfType<Border>().Any(border => border.Classes.Contains("DeviceDragHandle")));

                UserSettings.DeviceSorting = OwnedDeviceSorting.Manual;
                tab.RefreshWorkbooks();
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                // Never placed, in the order they were added - each with its handle.
                Assert.Equal(["Breadbin #2", "Amiga 500", "C64C silver"], Names());
                Assert.True(HasHandles());

                // The last card dragged by its handle to the top.
                DeviceListCard last = tab.DeviceCardsForTests[2];
                Border handle = last.GetLogicalDescendants().OfType<Border>().Single(border => border.Classes.Contains("DeviceDragHandle"));
                DeviceListCard first = tab.DeviceCardsForTests[0];
                Point start = handle.TranslatePoint(new Point(handle.Bounds.Width / 2, handle.Bounds.Height / 2), window)!.Value;
                Point top = first.TranslatePoint(new Point(first.Bounds.Width / 2, 2), window)!.Value;

                window.MouseDown(start, Avalonia.Input.MouseButton.Left);
                window.MouseMove(new Point(start.X, start.Y - 10), Avalonia.Input.RawInputModifiers.LeftMouseButton);
                Assert.True(tab.DeviceDragForTests!.IsDragging);
                window.MouseMove(new Point(start.X, top.Y), Avalonia.Input.RawInputModifiers.LeftMouseButton);
                window.MouseUp(new Point(start.X, top.Y), Avalonia.Input.MouseButton.Left);
                Dispatcher.UIThread.RunJobs();

                Assert.Equal(["C64C silver", "Breadbin #2", "Amiga 500"], Names());
                Assert.Equal(
                    ["C64C silver", "Breadbin #2", "Amiga 500"],
                    OwnedDeviceSorting.Sort(OwnedDeviceStore.GetAll(), OwnedDeviceSorting.Manual, key => (key, string.Empty)).Select(device => device.Name));

                // By hardware, board and name: the Amiga first, its hardware's name.
                UserSettings.DeviceSorting = OwnedDeviceSorting.ByHardwareBoardName;
                tab.RefreshWorkbooks();
                Assert.Equal("Amiga 500", Names()[0]);
                Assert.False(HasHandles());

                window.Close();
            }
            finally
            {
                UserSettings.WorkbooksScope = MyDevicesScope.CurrentBoard;
            }
        });
    }

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

    // The Configuration tab's scope set for a test's body, and put back however it ends - the setting is
    // process-wide (MyDevicesScope; "Show only devices for selected board" by default).
    private static void InScope(string scope, Action body)
    {
        string saved = UserSettings.WorkbooksScope;

        try
        {
            UserSettings.WorkbooksScope = scope;
            body();
        }
        finally
        {
            UserSettings.WorkbooksScope = saved;
        }
    }

    // ###########################################################################################
    // "My devices" says what it is for until there is one - naming the bar's "Add device", as the list has
    // none of its own since 2026-11-06 (owner request: "no need for it") - then lists the devices the
    // Configuration tab's scope has: every one, or the board shown's (owner request, 2026-11-06).
    // ###########################################################################################
    [Fact]
    public void My_devices_says_what_it_is_for_until_there_is_one_then_lists_the_devices_of_its_scope()
    {
        UiTest.Run(() => InScope(MyDevicesScope.AllBoards, () =>
        {
            TabWorkbooks tab = this.BuildTab();

            Assert.True(tab.GetControl<TextBlock>("NoDevicesText").IsVisible);
            Assert.Equal(OwnedDeviceWording.NoDevicesYet, Text(tab, "NoDevicesText"));
            Assert.Null(tab.FindControl<Button>("AddDeviceButton"));

            this.AddDevice("Breadbin #2");
            this.AddDevice("Plus/4", "Commodore Plus/4|PCB");
            tab.RefreshWorkbooks();

            Assert.False(tab.GetControl<TextBlock>("NoDevicesText").IsVisible);
            IReadOnlyList<DeviceListCard> list = tab.DeviceCardsForTests;
            Assert.Equal(2, list.Count);
            Assert.Contains("Breadbin #2", Texts(list[0]));
            Assert.Contains("Working", Texts(list[0]));

            UserSettings.WorkbooksScope = MyDevicesScope.CurrentBoard;
            tab.RefreshWorkbooks();

            Assert.Contains("Breadbin #2", Texts(Assert.Single(list)));
        }));
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
            Assert.Equal("Flea market", TextLinkRenderer.GetLinkText(tab.GetControl<TextBlock>("DeviceDescriptionText")));

            // The owner's own words for the list (2026-11-06), and no list of changes beside it.
            Assert.Equal("Workbooks associated to this device", Text(tab, "DeviceRepairsHeadingText"));
            Assert.DoesNotContain(Texts((Control)tab.GetControl<ScrollViewer>("DevicePageScroller").Content!), text => text.StartsWith("Changes made to it", StringComparison.Ordinal));

            Assert.StartsWith("2 workbooks", tab.GetControl<TextBlock>("DeviceRepairsSummaryText").Inlines!.Text, StringComparison.Ordinal);
            StackPanel repairs = tab.GetControl<StackPanel>("DeviceRepairsPanel");
            Assert.Equal(["PSU replaced", "No video"], repairs.Children.Select(row => ((WorkbookRecord)((Control)row).Tag!).Title));
            Assert.Contains(Texts((Control)repairs.Children[0]), text => text.StartsWith("On Commodore 64", StringComparison.Ordinal));
            Assert.DoesNotContain(Texts((Control)repairs.Children[1]), text => text.StartsWith("On Commodore 64", StringComparison.Ordinal));

            // The device's card is the one marked - no workbook card beside it.
            Assert.Contains("Selected", ((Control)tab.DeviceCardsForTests[0]).Classes);
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
            Assert.DoesNotContain("Selected", ((Control)tab.DeviceCardsForTests[0]).Classes);
        });
    }

    // ###########################################################################################
    // The description is several lines (owner request, 2026-11-06) - and the page keeps them, as it was
    // typed, with no label; its condition stands beside its status in stars.
    // ###########################################################################################
    [Fact]
    public void The_description_shows_on_the_page_in_its_own_lines_and_the_condition_in_stars()
    {
        OwnedDeviceRecord device = OwnedDeviceStore.Create(new OwnedDeviceDetails(
            "Breadbin #2", this.thisBoardKey, "", "", "A flea market in Odense\r\nCame with its box", "", Condition: 4))!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            TextBlock description = tab.GetControl<TextBlock>("DeviceDescriptionText");
            Assert.Equal("A flea market in Odense\nCame with its box", TextLinkRenderer.GetLinkText(description));
            Assert.Equal(TextWrapping.Wrap, description.TextWrapping);
            Assert.True(description.IsVisible);

            Panel stars = Assert.IsAssignableFrom<Panel>(tab.GetControl<Decorator>("DeviceConditionHost").Child);
            Assert.Equal(5, stars.Children.Count);
            Assert.Equal(4, stars.Children.Count(star => star.Classes.Contains("Lit")));
        });
    }

    // ###########################################################################################
    // "Add device" (the worklog bar's since 2026-11-06) adds a device with NO workbook (owner request,
    // 2026-11-06: "it is not logical that I need to create a workbook to create a new device"): the
    // form on any hardware and board - the board shown preselected, nothing locked - and the device it
    // adds listed with its page open. Listed in "Show all devices": it is on another board.
    // ###########################################################################################
    [Fact]
    public async Task Add_device_adds_one_on_any_board_and_opens_its_page()
    {
        string savedScope = UserSettings.WorkbooksScope;
        UserSettings.WorkbooksScope = MyDevicesScope.AllBoards;

        try
        {
            await this.AddDeviceOnAnyBoardAsync();
        }
        finally
        {
            UserSettings.WorkbooksScope = savedScope;
        }
    }

    private async Task AddDeviceOnAnyBoardAsync()
    {
        string shownBoard = this.thisBoardKey["Commodore 64|".Length..];
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Commodore 64", BoardName = shownBoard },
            new() { HardwareName = "Commodore 128", BoardName = "310378" },
        ];

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.BoardsOverrideForTests = boards;
            tab.BoardRegionsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>(BoardRegions.PalAndNtsc);

            OwnedDeviceWindow? shown = null;
            tab.ShowDeviceFormOverrideForTests = form =>
            {
                shown = form;
                form.Show();
                Assert.True(form.GetControl<ComboBox>("HardwareComboBox").IsEnabled);
                Assert.True(form.GetControl<ComboBox>("BoardComboBox").IsEnabled);
                Assert.Equal(this.thisBoardKey, form.Details().BoardKey);

                form.GetControl<ComboBox>("HardwareComboBox").SelectedItem = "Commodore 128";
                form.GetControl<TextBox>("NameTextBox").Text = "C128D";
                form.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

                return Task.FromResult(form.Saved);
            };

            await tab.AddDeviceAsync();

            Assert.NotNull(shown);
            OwnedDeviceRecord added = Assert.Single(OwnedDeviceStore.GetAll());
            Assert.Equal(("C128D", "Commodore 128|310378"), (added.Name, added.BoardKey));
            Assert.Empty(WorklogManager.GetAllWorkbooks());

            Assert.Equal(added.Id, tab.SelectedDeviceIdForTests);
            Assert.True(tab.IsDevicePageShown);
            Assert.Equal("C128D", Text(tab, "DeviceNameText"));
            Assert.Contains("C128D", Texts((Control)Assert.Single(tab.DeviceCardsForTests)));
        });
    }

    // ###########################################################################################
    // *** PHOTOS A NEW DEVICE COULD NOT TAKE IN ARE SAID ON ITS PAGE (code review, 2026-11-08). *** The note was set after
    // the page was drawn - inside the photos' section, which the page hides for a device with no photos - so with none
    // copied in, it was never seen.
    // ###########################################################################################
    [Fact]
    public async Task Photos_a_new_device_could_not_take_in_are_said_on_its_page()
    {
        string gone = this.thisWorkspace.Path_("Picked", "gone.jpg");

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.BoardsOverrideForTests = [new() { HardwareName = "Commodore 64", BoardName = this.thisBoardKey["Commodore 64|".Length..] }];
            tab.BoardRegionsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>([]);
            tab.ShowDeviceFormOverrideForTests = async form =>
            {
                form.AskPhotoOverrideForTests = _ => Task.FromResult<WorklogAddPhotoWindow.PhotoResult?>(new(gone, "The front"));
                await form.AddPhotoAsync();
                form.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
                form.Show();
                form.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                await form.SavingForTests;

                return form.Saved;
            };

            Assert.True(await tab.AddDeviceAsync());

            Assert.True(tab.IsDevicePageShown);
            Assert.True(tab.GetControl<StackPanel>("DevicePhotosSection").IsVisible, "The photos' section, the note in it, is hidden");
            TextBlock note = tab.GetControl<TextBlock>("DevicePhotosNoteText");
            Assert.True(note.IsVisible);
            Assert.Equal(OwnedDeviceWording.PhotosAdded(0, 1), note.Text);
        });
    }

    // ###########################################################################################
    // *** A COMMENT'S AND A PHOTO'S COMMENT ARE SHOWN WITH THE INDENT THAT SAYS WHAT THEIR FIRST LINE IS (code review,
    // 2026-11-08). *** Stored right ("  - replaced C14", a bullet nested once), the page trimmed them as it showed them,
    // and a nested bullet read as one of the first level.
    // ###########################################################################################
    [Fact]
    public void The_page_shows_a_comments_first_line_at_the_level_it_was_saved_at()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddComment(device.Id, "  - replaced C14\n- U8", new DateTime(2026, 11, 8, 9, 0, 0));
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("front.jpg"), "  - the front")]);

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            List<string> shown = tab.GetControl<StackPanel>("DeviceCommentsPanel").GetLogicalDescendants().OfType<TextBlock>()
                .Concat(tab.GetControl<WrapPanel>("DevicePhotosPanel").GetLogicalDescendants().OfType<TextBlock>())
                .Select(TextLinkRenderer.GetLinkText)
                .OfType<string>()
                .ToList();

            Assert.Contains("  - replaced C14\n- U8", shown);
            Assert.Contains("  - the front", shown);
        });
    }

    // Called off, nothing is added and the page stays as it was.
    [Fact]
    public async Task Add_device_called_off_adds_nothing()
    {
        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.BoardsOverrideForTests = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];
            tab.BoardRegionsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>([]);
            tab.ShowDeviceFormOverrideForTests = _ => Task.FromResult<OwnedDeviceRecord?>(null);

            await tab.AddDeviceAsync();

            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.False(tab.IsDevicePageShown);
        });
    }

    // ###########################################################################################
    // THE PAGE IS ONLY READ (owner request, 2026-11-07: "it should not allow for direct edits or 'Add X', as this should
    // be a view-only place. If anything needs to be added or edited, the user should go to the 'Edit device' button"):
    // no box to type in and no button but its own four - "Edit device", "Delete device" and the two exports - and no
    // "New repair" ('remove the text ""New repair" starts one on its board." and then also the "New repair" button').
    // ###########################################################################################
    [Fact]
    public void The_page_is_only_read_its_only_buttons_edit_delete_and_export()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddComment(device.Id, "Recapped", new DateTime(2026, 11, 7, 9, 0, 0));
        OwnedDeviceStore.AddLink(device.Id, "Service manual", "https://zimmers.net");
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("front.jpg"), "The front")]);

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            var window = new Window { Content = tab, Width = 1200, Height = 1400 };
            window.Show();

            try
            {
                tab.SelectDevice(device.Id);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                ScrollViewer page = tab.GetControl<ScrollViewer>("DevicePageScroller");
                Assert.DoesNotContain(page.GetVisualDescendants().OfType<TextBox>(), box => box.IsEffectivelyVisible);
                Assert.Equal(
                    ["EditDeviceButton", "DeleteDeviceButton", "ExportHistoryButton", "ExportHistoryZipButton"],
                    page.GetVisualDescendants().OfType<Button>().Where(button => button.IsEffectivelyVisible).Select(button => button.Name));
                Assert.Null(tab.FindControl<Button>("NewRepairButton"));
                Assert.Null(tab.FindControl<Button>("AddDevicePhotosButton"));

                // The notes went with their box - the description says it (OwnedDeviceStore.FoldNotes).
                Assert.Null(tab.FindControl<TextBlock>("DeviceNotesHeadingText"));
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // A HEADLINE ONLY OVER SOMETHING (owner request, 2026-11-07: "Only show headlines if it has data, which means if it
    // has no comments, then do not show the headline with comments and then 'No comments added'") - while "no workbooks
    // associated" is still said ("It is fine though, that it shows that no workbooks are associated").
    // ###########################################################################################
    [Fact]
    public void A_part_with_nothing_in_it_has_no_headline_but_no_workbooks_is_said()
    {
        OwnedDeviceRecord device = this.AddDevice();

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            Assert.False(tab.GetControl<StackPanel>("DeviceLinksSection").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("DeviceCommentsSection").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("DevicePhotosSection").IsVisible);
            Assert.True(tab.GetControl<TextBlock>("NoDeviceRepairsText").IsVisible);
            Assert.Equal(OwnedDeviceWording.NoRepairsYet, Text(tab, "NoDeviceRepairsText"));

            OwnedDeviceStore.AddComment(device.Id, "Recapped", new DateTime(2026, 11, 7, 9, 0, 0));
            tab.RefreshWorkbooks();

            Assert.True(tab.GetControl<StackPanel>("DeviceCommentsSection").IsVisible);
            Assert.False(tab.GetControl<StackPanel>("DeviceLinksSection").IsVisible);
        });
    }

    // ###########################################################################################
    // Each photo is decoded once, however often the tab is drawn again (code review, 2026-10-28) - and the device form
    // showing it reads no second copy (code review, 2026-11-08: each kept its own, DevicePhotoThumbnails since).
    // ###########################################################################################
    [Fact]
    public void A_photo_on_the_page_is_decoded_once()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("front.jpg"), "The front")]);

        UiTest.Run(() =>
        {
            int readsBefore = DevicePhotoThumbnails.ReadsMadeForTests;
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            tab.RefreshWorkbooks();
            tab.RefreshWorkbooks();

            var form = new OwnedDeviceWindow();
            form.InitializeForEdit(OwnedDeviceStore.Get(device.Id)!, [new() { HardwareName = "Commodore 64", BoardName = "250407" }], _ => Task.FromResult<IReadOnlyList<string>>([]));

            Assert.Equal(1, DevicePhotoThumbnails.ReadsMadeForTests - readsBefore);
            Assert.Equal("Photos (1)", Text(tab, "DevicePhotosHeadingText"));
            Control tile = (Control)Assert.Single(tab.GetControl<WrapPanel>("DevicePhotosPanel").Children);
            Assert.Empty(tile.GetLogicalDescendants().OfType<Button>());
        });
    }

    // ###########################################################################################
    // A photo on the page clicked opens it whole in the worklog's viewer, which zooms with the wheel (owner request,
    // 2026-11-06: "when I click a photo in here and in the worklog, that I then can zoom in/out with scrollwheel").
    // ###########################################################################################
    [Fact]
    public void A_photo_on_the_page_clicked_opens_the_viewer_on_it()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("front.jpg"), "The front")]);

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            WorklogPhotoViewerWindow? shown = null;
            tab.ShowPhotoViewerOverrideForTests = viewer => shown = viewer;
            var window = new Window { Content = tab, Width = 1200, Height = 900 };
            window.Show();

            try
            {
                tab.SelectDevice(device.Id);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                Image photo = tab.GetControl<WrapPanel>("DevicePhotosPanel").GetVisualDescendants().OfType<Image>().Single();
                Point at = photo.TranslatePoint(new Point(photo.Bounds.Width / 2, photo.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, Avalonia.Input.MouseButton.Left);
                window.MouseUp(at, Avalonia.Input.MouseButton.Left);

                Assert.NotNull(shown);
                Assert.Equal("front.jpg", shown!.Title);
                Assert.Equal("The front", shown.GetControl<TextBlock>("CommentText").Text);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // THE PAGE, A HEADLINE PER PART (owner request, 2026-11-06: '"Photos" should sit just above the photos'): Links of
    // interest, Comments, Photos in that order (Notes went on 2026-11-07, owner request); a link
    // clicked opens its address; each photo in a panel of its own, its comment under it ("I want each photo to have its
    // own panel and with its text displayed below").
    // ###########################################################################################
    [Fact]
    public void The_page_has_a_headline_per_part_its_links_open_and_each_photo_has_its_panel()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddLink(device.Id, "Service manual", "https://zimmers.net");
        OwnedDeviceStore.AddComment(device.Id, "Recapped", new DateTime(2026, 11, 7, 9, 0, 0));
        OwnedDeviceStore.AddPhotos(device.Id, [new OwnedDevicePhotoSource(this.Picture("front.jpg"), "The front")]);

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            var opened = new List<string>();
            tab.OpenLinkOverrideForTests = opened.Add;
            var window = new Window { Content = tab, Width = 1200, Height = 1400 };
            window.Show();

            try
            {
                tab.SelectDevice(device.Id);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                double Y(string name) => tab.GetControl<TextBlock>(name).TranslatePoint(default, window)!.Value.Y;
                Assert.Equal("Links of interest (1)", Text(tab, "DeviceLinksHeadingText"));
                Assert.Equal("Photos (1)", Text(tab, "DevicePhotosHeadingText"));
                Assert.True(Y("DeviceLinksHeadingText") < Y("DeviceCommentsHeadingText"));
                Assert.True(Y("DeviceCommentsHeadingText") < Y("DevicePhotosHeadingText"));

                Control link = Assert.Single(tab.GetControl<StackPanel>("DeviceLinksPanel").Children);
                Assert.Contains("Service manual", Texts(link));
                Point at = link.TranslatePoint(new Point(link.Bounds.Width / 2, link.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, Avalonia.Input.MouseButton.Left);
                window.MouseUp(at, Avalonia.Input.MouseButton.Left);
                Assert.Equal(["https://zimmers.net"], opened);

                Border panel = Assert.IsType<Border>(Assert.Single(tab.GetControl<WrapPanel>("DevicePhotosPanel").Children));
                Image photo = panel.GetVisualDescendants().OfType<Image>().Single();
                TextBlock comment = panel.GetVisualDescendants().OfType<TextBlock>().First(text => TextLinkRenderer.GetLinkText(text) == "The front");
                Assert.True(comment.TranslatePoint(default, window)!.Value.Y > photo.TranslatePoint(default, window)!.Value.Y + photo.Bounds.Height - 1);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // *** A LINK TYPED WITHOUT ITS SCHEME OPENS ALL THE SAME (code review, 2026-11-08). *** A link is stored as typed, so
    // "www.c64-wiki.com" is as likely as "https://...": handed over as typed, it was tried as a file and refused, and the
    // error card said it could not be opened - while the exported PDF and the Contribute tab opened it (LinkTarget, the one
    // rule).
    // ###########################################################################################
    [Fact]
    public void A_device_link_typed_without_its_scheme_opens_as_a_web_address()
    {
        OwnedDeviceRecord device = this.AddDevice();
        OwnedDeviceStore.AddLink(device.Id, "Wiki", "www.c64-wiki.com");

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            var opened = new List<string>();
            tab.OpenLinkOverrideForTests = opened.Add;
            var window = new Window { Content = tab, Width = 1200, Height = 1400 };
            window.Show();

            try
            {
                tab.SelectDevice(device.Id);
                Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();

                Control link = Assert.Single(tab.GetControl<StackPanel>("DeviceLinksPanel").Children);
                Point at = link.TranslatePoint(new Point(link.Bounds.Width / 2, link.Bounds.Height / 2), window)!.Value;
                window.MouseDown(at, Avalonia.Input.MouseButton.Left);
                window.MouseUp(at, Avalonia.Input.MouseButton.Left);

                Assert.Equal(["https://www.c64-wiki.com"], opened);
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // The page lists the device's comments, newest first and dated, as the form does - none, no headline; "Edit device"
    // called off after a comment was added there (saved at once - owner decision, 2026-11-06) still shows it.
    // ###########################################################################################
    [Fact]
    public async Task The_page_lists_the_comments_newest_first_and_one_added_in_a_form_called_off_shows()
    {
        OwnedDeviceRecord device = this.AddDevice();

        await UiTest.RunAsync(async () =>
        {
            TabWorkbooks tab = this.BuildTab();
            tab.SelectDevice(device.Id);

            Assert.False(tab.GetControl<StackPanel>("DeviceCommentsSection").IsVisible);

            OwnedDeviceStore.AddComment(device.Id, "Bought at a flea market", new DateTime(2026, 10, 1, 9, 5, 0));
            tab.BoardsOverrideForTests = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];
            tab.BoardRegionsOverrideForTests = _ => Task.FromResult<IReadOnlyList<string>>([]);
            tab.ShowDeviceFormOverrideForTests = async form =>
            {
                form.AskCommentOverrideForTests = _ => Task.FromResult<string?>("Recapped");
                await form.AddCommentAsync();
                return null;
            };

            await tab.EditSelectedDeviceAsync();

            Assert.Equal("Comments (2)", Text(tab, "DeviceCommentsHeadingText"));
            Assert.True(tab.GetControl<StackPanel>("DeviceCommentsSection").IsVisible);
            List<List<string>> rows = tab.GetControl<StackPanel>("DeviceCommentsPanel").Children
                .Select(row => row.GetLogicalDescendants().OfType<TextBlock>()
                    .Select(text => text.Text ?? TextLinkRenderer.GetLinkText(text) ?? string.Empty).ToList())
                .ToList();
            Assert.Equal("Recapped", rows[0][1]);
            Assert.Equal(["2026-10-01 09:05", "Bought at a flea market"], rows[1]);
        });
    }

    // ###########################################################################################
    // A device clicked is the device PICKED (owner request, 2026-11-06: "it will only show the workbooks for this
    // device"): the workbook list holds its workbooks only, on any board - one on another board naming its board -
    // and "Clear" shows every workbook again. A workbook clicked picks the device it is on, one on none all devices.
    // ###########################################################################################
    [Fact]
    public void A_device_clicked_narrows_the_workbooks_to_its_own_and_Clear_shows_them_all_again()
    {
        OwnedDeviceRecord device = this.AddDevice();
        WorkbookRecord own = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "", device.Id)!;
        WorkbookRecord elsewhere = WorklogManager.CreateWorkbook(this.thisOtherBoardKey, "PSU replaced", "", device.Id)!;
        WorkbookRecord onNone = WorklogManager.CreateWorkbook(this.thisBoardKey, "Dead keyboard", "")!;

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            StackPanel workbooks = tab.GetControl<StackPanel>("WorkbookListPanel");
            List<int> Listed() => workbooks.Children.Select(card => ((WorkbookRecord)card.Tag!).Id).ToList();

            Assert.Equal([onNone.Id, own.Id], Listed());

            tab.SelectDevice(device.Id);

            Assert.Equal(device.Id, UserSettings.WorklogBarDeviceId);
            Assert.Equal([elsewhere.Id, own.Id], Listed());
            Assert.Contains(Texts((Control)workbooks.Children[0]), text => text.Contains("250425", StringComparison.Ordinal));
            Assert.DoesNotContain(Texts((Control)workbooks.Children[1]), text => text.Contains("250407", StringComparison.Ordinal));
            Assert.True(tab.GetControl<Button>("ClearDeviceButton").IsEnabled);

            tab.ClearDevicePick();

            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.False(tab.IsDevicePageShown);
            Assert.Equal([onNone.Id, own.Id], Listed());
            Assert.False(tab.GetControl<Button>("ClearDeviceButton").IsEnabled);

            tab.SelectWorkbookForTests(own.Id);
            Assert.Equal(device.Id, UserSettings.WorklogBarDeviceId);
            Assert.Contains("Selected", ((Control)tab.DeviceCardsForTests[0]).Classes);

            tab.SelectWorkbookForTests(elsewhere.Id, this.thisOtherBoardKey);
            Assert.Equal(device.Id, UserSettings.WorklogBarDeviceId);

            tab.ClearDevicePick();
            tab.SelectWorkbookForTests(onNone.Id);
            Assert.Null(UserSettings.WorklogBarDeviceId);
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

    // ###########################################################################################
    // *** A WORKBOOK PUT ON ANOTHER DEVICE TAKES THE DEVICE PICKED ALONG (code review, 2026-11-08). *** CRT works only in
    // the device picked's workbooks, so the workbook shown, put on another device - or on none - dropped out of the list
    // and the bar, and the device's other workbook became the one worked in. The pick follows the workbook, as it does
    // for a workbook activated off the device picked (OwnedDeviceRepairs.PickFollowing).
    // ###########################################################################################
    [Fact]
    public void A_workbook_put_on_another_device_takes_the_device_picked_along()
    {
        OwnedDeviceRecord first = this.AddDevice("Breadbin #2");
        OwnedDeviceRecord second = this.AddDevice("C64C silver");
        WorkbookRecord older = WorklogManager.CreateWorkbook(this.thisBoardKey, "Bad RAM", "")!;
        WorkbookRecord moved = WorklogManager.CreateWorkbook(this.thisBoardKey, "No video", "")!;
        WorklogManager.SetWorkbookDevice(older.Id, first.Id);
        WorklogManager.SetWorkbookDevice(moved.Id, first.Id);

        UiTest.Run(() =>
        {
            TabWorkbooks tab = this.BuildTab();
            string Header() => tab.GetControl<TextBlock>("WorkbookHeaderTitleText").Text ?? string.Empty;

            tab.SelectWorkbookForTests(moved.Id);
            Assert.Equal(first.Id, UserSettings.WorklogBarDeviceId);
            Assert.Contains("No video", Header(), StringComparison.Ordinal);

            tab.PutShownWorkbookOnDevice(second.Id);

            Assert.Equal(second.Id, UserSettings.WorklogBarDeviceId);
            Assert.Contains("No video", Header(), StringComparison.Ordinal);
            Assert.Equal(moved.Id, UserSettings.GetActiveWorkbookId(this.thisBoardKey));

            // On no device: all devices.
            tab.PutShownWorkbookOnDevice(null);

            Assert.Null(UserSettings.WorklogBarDeviceId);
            Assert.Contains("No video", Header(), StringComparison.Ordinal);
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
            Assert.Equal("Add workbook", window.Title);

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
            Assert.Equal("Add workbook", window.Title);

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
    // "Add device": the board shown preselected among the boards CRT knows, region and status "Unknown",
    // each first in its list (owner request, 2026-11-06: "Not set" and "Working" until then, and "Not mine
    // any more" offered no more) - and what is saved is what the form shows.
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
            window.Initialize(boards, "Commodore 64|250425", boardRegions: _ => Task.FromResult<IReadOnlyList<string>>(BoardRegions.PalAndNtsc));
            window.Show();

            Assert.Equal("Commodore 64", window.GetControl<ComboBox>("HardwareComboBox").SelectedItem);
            Assert.Equal(["250407", "250425"], window.GetControl<ComboBox>("BoardComboBox").Items.Cast<string>());
            Assert.Equal("250425", window.GetControl<ComboBox>("BoardComboBox").SelectedItem);
            Assert.True(window.GetControl<StackPanel>("RegionPanel").IsVisible);
            Assert.Equal(["Unknown", "PAL", "NTSC"], window.GetControl<ComboBox>("RegionComboBox").Items.Cast<string>());
            Assert.Equal("Unknown", window.GetControl<ComboBox>("RegionComboBox").SelectedItem);
            Assert.Equal(["Unknown", "Working", "Faulty", "In repair", "For parts"], window.GetControl<ComboBox>("StatusComboBox").Items.Cast<string>());
            Assert.Equal("Unknown", window.GetControl<ComboBox>("StatusComboBox").SelectedItem);
            Assert.False(window.GetControl<Button>("SaveButton").IsEnabled);

            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            window.GetControl<ComboBox>("RegionComboBox").SelectedItem = "PAL";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            OwnedDeviceRecord saved = Assert.Single(OwnedDeviceStore.GetAll());
            Assert.Equal(("Breadbin #2", "Commodore 64|250425", "PAL", OwnedDeviceStatus.Unknown), (saved.Name, saved.BoardKey, saved.Region, saved.Status));
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
            window.Initialize(boards, "Commodore 64|250407", boardRegions: _ => Task.FromResult<IReadOnlyList<string>>([]));
            window.Show();
            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";

            Button cancel = window.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, "Cancel"));
            cancel.RaiseEvent(Enter());

            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.Null(window.Saved);

            window.GetControl<TextBox>("NameTextBox").RaiseEvent(Enter());

            Assert.Single(OwnedDeviceStore.GetAll());
        });
    }

    private static Avalonia.Input.KeyEventArgs Enter(Avalonia.Input.KeyModifiers modifiers = Avalonia.Input.KeyModifiers.None) =>
        new() { RoutedEvent = Avalonia.Input.InputElement.KeyDownEvent, Key = Avalonia.Input.Key.Enter, KeyModifiers = modifiers };

    // ###########################################################################################
    // "Description" is a box of several lines (owner requests, 2026-11-06: "Please have that as a
    // textbox", then 'change text "Where did it come from" to "Description"'): it wraps, Enter in it saves
    // NOTHING - it is the box's new line - and Ctrl+Enter saves, as in "Add workbook"'s note. Its lines
    // reach the device as typed.
    // ###########################################################################################
    [Fact]
    public void The_description_takes_several_lines_and_Enter_there_is_no_save()
    {
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407", boardRegions: _ => Task.FromResult<IReadOnlyList<string>>([]));
            window.Show();

            TextBox description = window.GetControl<TextBox>("DescriptionTextBox");
            Assert.True(description.AcceptsReturn);
            Assert.Equal(TextWrapping.Wrap, description.TextWrapping);
            Assert.True(description.MinHeight >= 3 * description.FontSize * 1.2, $"The box is {description.MinHeight}px - not three lines");

            // It grows with its text (owner request, 2026-11-07: "auto-expand to fit all text").
            description.Text = string.Join("\n", Enumerable.Range(1, 12).Select(line => $"Line {line}"));
            Dispatcher.UIThread.RunJobs();
            double grown = description.Bounds.Height;
            Assert.True(grown > description.MinHeight + 3 * description.FontSize, $"The box is {grown}px for twelve lines");
            description.Text = string.Empty;
            Assert.Contains("Description", window.GetVisualDescendants().OfType<TextBlock>().Select(text => text.Text));
            Assert.DoesNotContain("Where it came from", window.GetVisualDescendants().OfType<TextBlock>().Select(text => text.Text));

            window.GetControl<TextBox>("NameTextBox").Text = "Breadbin #2";
            description.Text = "A flea market in Odense\nCame with its box";

            description.RaiseEvent(Enter());
            Assert.Empty(OwnedDeviceStore.GetAll());
            Assert.Null(window.Saved);

            description.RaiseEvent(Enter(Avalonia.Input.KeyModifiers.Control));

            Assert.Equal("A flea market in Odense\nCame with its box", Assert.Single(OwnedDeviceStore.GetAll()).Description);
        });
    }

    // ###########################################################################################
    // REGION FOLLOWS THE BOARD (owner request, 2026-11-06: "currently I always see the drop-down of a
    // region, even if that board does not have the region concept"): shown only for a board that comes
    // in PAL and NTSC, as the board picked changes - and a region picked is not saved for a board the
    // box went away for.
    // ###########################################################################################
    [Fact]
    public void Region_shows_only_while_the_board_picked_comes_in_PAL_and_NTSC()
    {
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Commodore 64", BoardName = "250407" },
            new() { HardwareName = "Commodore 64", BoardName = "KU-14194HB" },
            new() { HardwareName = "ZX Spectrum", BoardName = "Issue 2" },
        ];

        var asked = new List<string>();
        var regional = new HashSet<string>(["Commodore 64|250407"], StringComparer.OrdinalIgnoreCase);

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407", boardRegions: boardKey =>
            {
                asked.Add(boardKey);
                return Task.FromResult<IReadOnlyList<string>>(regional.Contains(boardKey) ? BoardRegions.PalAndNtsc : []);
            });
            window.Show();

            StackPanel region = window.GetControl<StackPanel>("RegionPanel");
            ComboBox board = window.GetControl<ComboBox>("BoardComboBox");
            Assert.True(region.IsVisible);

            window.GetControl<ComboBox>("RegionComboBox").SelectedItem = "PAL";
            board.SelectedItem = "KU-14194HB";
            Assert.False(region.IsVisible);
            Assert.Equal(string.Empty, window.Details().Region);

            board.SelectedItem = "250407";
            Assert.True(region.IsVisible);

            // Another hardware: its first board, which has none.
            window.GetControl<ComboBox>("HardwareComboBox").SelectedItem = "ZX Spectrum";
            Assert.False(region.IsVisible);

            // Each board asked once.
            Assert.Equal(["Commodore 64|250407", "Commodore 64|KU-14194HB", "ZX Spectrum|Issue 2"], asked);

            window.GetControl<TextBox>("NameTextBox").Text = "Spectrum";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.Equal(string.Empty, Assert.Single(OwnedDeviceStore.GetAll()).Region);
        });
    }

    // ###########################################################################################
    // A board's answer still being read leaves the form as it is - no flicker - and an answer for a
    // board no longer picked is kept for later, never shown for the board picked now.
    // ###########################################################################################
    [Fact]
    public async Task Region_waits_for_the_boards_answer_and_takes_none_for_another_board()
    {
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Commodore 64", BoardName = "250407" },
            new() { HardwareName = "Commodore 64", BoardName = "KU-14194HB" },
        ];

        var answers = new Dictionary<string, TaskCompletionSource<IReadOnlyList<string>>>(StringComparer.OrdinalIgnoreCase)
        {
            ["Commodore 64|250407"] = new(),
            ["Commodore 64|KU-14194HB"] = new(),
        };

        await UiTest.RunAsync(async () =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407", boardRegions: boardKey => answers[boardKey].Task);
            window.Show();

            StackPanel region = window.GetControl<StackPanel>("RegionPanel");
            Assert.False(region.IsVisible);

            Task first = window.RegionsFollowedForTests;
            answers["Commodore 64|250407"].SetResult(BoardRegions.PalAndNtsc);
            await first;
            Assert.True(region.IsVisible);

            // The next board not answered yet: shown as it was.
            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "KU-14194HB";
            Assert.True(region.IsVisible);
            Task second = window.RegionsFollowedForTests;

            // Back on the first board, answered already - then the second's answer comes, for a board
            // not picked: nothing changes.
            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "250407";
            answers["Commodore 64|KU-14194HB"].SetResult([]);
            await second;
            Assert.True(region.IsVisible);

            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "KU-14194HB";
            Assert.False(region.IsVisible);
        });
    }

    // A board whose question fails has no Region box - the form goes on.
    [Fact]
    public void A_board_whose_regions_cannot_be_read_shows_no_region()
    {
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Commodore 64|250407", boardRegions: _ => Task.FromException<IReadOnlyList<string>>(new IOException("held by another program")));
            window.Show();

            Assert.False(window.GetControl<StackPanel>("RegionPanel").IsVisible);
        });
    }

    // ###########################################################################################
    // A region a device has on a board WITHOUT regions (saved before the box followed the board) is
    // kept - not shown, never wiped by opening and saving the form.
    // ###########################################################################################
    [Fact]
    public void A_region_stored_on_a_board_without_regions_is_kept_and_not_shown()
    {
        OwnedDeviceRecord device = this.AddDevice("Spectrum", "ZX Spectrum|Issue 2");
        HardwareBoardEntry[] boards = [new() { HardwareName = "ZX Spectrum", BoardName = "Issue 2" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.InitializeForEdit(device, boards, _ => Task.FromResult<IReadOnlyList<string>>([]));
            window.Show();

            Assert.False(window.GetControl<StackPanel>("RegionPanel").IsVisible);

            window.GetControl<TextBox>("SerialTextBox").Text = "S-1";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            OwnedDeviceRecord saved = OwnedDeviceStore.Get(device.Id)!;
            Assert.Equal(("PAL", "S-1"), (saved.Region, saved.SerialNumber));
        });
    }

    // ###########################################################################################
    // A REGION PICKED FOR ONE BOARD IS NOT CARRIED TO ANOTHER (code review, 2026-11-09: "Japan" picked on a board in EU and
    // Japan stayed picked on a PAL/NTSC board, listed after its regions, and was saved so). On a board without it nothing
    // is picked; back on a board with it, it is picked once more only when picked there again.
    // ###########################################################################################
    [Fact]
    public void A_region_picked_for_one_board_is_not_carried_to_a_board_without_it()
    {
        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Bang & Olufsen", BoardName = "Beocenter" },
            new() { HardwareName = "Bang & Olufsen", BoardName = "Beomaster" },
        ];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.Initialize(boards, "Bang & Olufsen|Beocenter", boardRegions: boardKey =>
                Task.FromResult<IReadOnlyList<string>>(boardKey.EndsWith("Beocenter", StringComparison.Ordinal) ? ["EU", "Japan"] : BoardRegions.PalAndNtsc));
            window.Show();

            ComboBox region = window.GetControl<ComboBox>("RegionComboBox");
            region.SelectedItem = "Japan";

            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "Beomaster";

            Assert.Equal(["Unknown", "PAL", "NTSC"], region.Items.Cast<string>());
            Assert.Equal("Unknown", region.SelectedItem);
            Assert.Equal(string.Empty, window.Details().Region);
        });
    }

    // The device's own region comes back on its own board - and on another, only where that board has it.
    [Fact]
    public void A_devices_own_region_comes_back_on_its_own_board()
    {
        OwnedDeviceRecord device = OwnedDeviceStore.Create(
            new OwnedDeviceDetails("Beocenter", "Bang & Olufsen|Beocenter", "Japan", "S-1", "Flea market", OwnedDeviceStatus.Working))!;

        HardwareBoardEntry[] boards =
        [
            new() { HardwareName = "Bang & Olufsen", BoardName = "Beocenter" },
            new() { HardwareName = "Bang & Olufsen", BoardName = "Beomaster" },
        ];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.InitializeForEdit(device, boards, boardKey =>
                Task.FromResult<IReadOnlyList<string>>(boardKey.EndsWith("Beocenter", StringComparison.Ordinal) ? ["EU", "Japan"] : BoardRegions.PalAndNtsc));
            window.Show();

            ComboBox region = window.GetControl<ComboBox>("RegionComboBox");
            ComboBox board = window.GetControl<ComboBox>("BoardComboBox");
            Assert.Equal("Japan", region.SelectedItem);

            board.SelectedItem = "Beomaster";
            Assert.Equal("Unknown", region.SelectedItem);

            // Nor is it offered there (second code review, 2026-11-09: listed after PAL and NTSC, it could be picked and saved).
            Assert.Equal(["Unknown", "PAL", "NTSC"], region.Items.Cast<string>());

            board.SelectedItem = "Beocenter";
            Assert.Equal("Japan", region.SelectedItem);
        });
    }

    // ###########################################################################################
    // The regions a board comes in, asked of it as CRT views it - read from its board file, by BoardRegions' rule: PAL and
    // NTSC, or names of its own (2026-11-09). A board CRT does not list has none.
    // ###########################################################################################
    [Fact]
    public async Task A_board_says_which_regions_it_comes_in_from_its_board_file()
    {
        string regionalKey = "Commodore 64|250407 " + Guid.NewGuid().ToString("N");
        string ownKey = "Bang & Olufsen|Beocenter " + Guid.NewGuid().ToString("N");
        string plainKey = "ZX Spectrum|Issue 2 " + Guid.NewGuid().ToString("N");

        try
        {
            PublishedBoardTree.Load(this.thisWorkspace.Path_("Regional"), regionalKey, new BoardData { Regions = [.. BoardRegions.PalAndNtsc] });
            IReadOnlyList<string> regional = await OwnedDeviceWindow.BoardRegionsAsViewedAsync(regionalKey);
            IReadOnlyList<string> unlisted = await OwnedDeviceWindow.BoardRegionsAsViewedAsync(plainKey);

            PublishedBoardTree.Load(this.thisWorkspace.Path_("Own"), ownKey, new BoardData { Regions = ["EU", "US"] });
            IReadOnlyList<string> own = await OwnedDeviceWindow.BoardRegionsAsViewedAsync(ownKey);

            PublishedBoardTree.Load(this.thisWorkspace.Path_("Plain"), plainKey, new BoardData());
            IReadOnlyList<string> plain = await OwnedDeviceWindow.BoardRegionsAsViewedAsync(plainKey);

            Assert.Equal(["PAL", "NTSC"], regional);
            Assert.Equal(["EU", "US"], own);
            Assert.Empty(unlisted);
            Assert.Empty(plain);
        }
        finally
        {
            PublishedBoardTree.Unload(this.thisWorkspace.Path_("Empty"));
        }
    }

    // ###########################################################################################
    // Edited onto another board: only the board changes - nothing is recorded beside it (owner request,
    // 2026-11-06: "remove the 'Changes made to it', as this is exactly what the workbooks will
    // cover"). A board CRT no longer lists stays offered.
    // ###########################################################################################
    [Fact]
    public void Edit_device_keeps_a_board_CRT_no_longer_lists_and_saves_the_board_picked()
    {
        OwnedDeviceRecord device = this.AddDevice("Breadbin #2", "Commodore 64|250466");
        HardwareBoardEntry[] boards = [new() { HardwareName = "Commodore 64", BoardName = "250407" }];

        UiTest.Run(() =>
        {
            var window = new OwnedDeviceWindow();
            window.InitializeForEdit(device, boards, _ => Task.FromResult<IReadOnlyList<string>>(BoardRegions.PalAndNtsc));
            window.Show();

            Assert.Equal("Save device", window.GetControl<Button>("SaveButton").Content);
            Assert.Equal(["250407", "250466"], window.GetControl<ComboBox>("BoardComboBox").Items.Cast<string>());
            Assert.Equal("250466", window.GetControl<ComboBox>("BoardComboBox").SelectedItem);
            Assert.Equal("Flea market", window.GetControl<TextBox>("DescriptionTextBox").Text);
            Assert.Equal("PAL", window.GetControl<ComboBox>("RegionComboBox").SelectedItem);

            window.GetControl<ComboBox>("BoardComboBox").SelectedItem = "250407";
            window.GetControl<Button>("SaveButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

            OwnedDeviceRecord saved = OwnedDeviceStore.Get(device.Id)!;
            Assert.Equal(("Commodore 64|250407", "PAL"), (saved.BoardKey, saved.Region));
            Assert.Null(saved.Unknown);
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

        // The board shown's picture lies under the data root the other board is loaded from below, and
        // its row names it there, as in CRT (WorkbooksBoardPreviewTests' header).
        string shownPicture = this.thisWorkspace.Path_("Data", "Shown", "shown.png");
        Directory.CreateDirectory(Path.GetDirectoryName(shownPicture)!);
        File.WriteAllBytes(shownPicture, [1, 2, 3]);

        var shownBoard = new BoardData { Schematics = [new BoardSchematicEntry { SchematicName = "Sheet 1", SchematicImageFile = "Shown/shown.png" }] };
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
