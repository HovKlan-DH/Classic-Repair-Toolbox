using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// THE LISTS ON THE FAR LEFT ARE THE CONTRIBUTE TAB'S SELECTOR WHILE IT IS SHOWN (Main.ContributeLists.cs -
// owner requests, 2026-11-03: "have those lists work as selector for the component. Only one component should
// be possible to select at a time", and "should select all in the 'Category' list ... but only the first
// time"). Lists of its own in that place since 2026-11-04, so these read the ones a person sees on the tab shown.
// Main's StartAsync is never called (MainWindowTests' header).
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class MainContributeListsTests : IDisposable
{
    private const string BoardKey = "Commodore 64|250407";

    private readonly TempWorkspace thisWorkspace = new();

    public MainContributeListsTests()
    {
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
        DraftManager.LoadFrom(this.thisWorkspace.Path_("Drafts"));
    }

    public void Dispose()
    {
        DraftManager.LoadFrom(string.Empty);
        WorklogManager.LoadFrom(this.thisWorkspace.Path_("Workbook-after"));
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    private static BoardData Board() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" },
            new ComponentEntry { BoardLabel = "U2", FriendlyName = "SID", Category = "IC" },
            new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
        ],
    };

    // A window on a board, the lists as a load leaves them: both categories ticked, U1 selected, on the
    // tab it opened on.
    private static CRT.Main WindowOnBoard()
    {
        var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };

        window.UseBoardForTests(
            new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" },
            Board());

        ListShown(window).Selection.Select(0);
        Dispatcher.UIThread.RunJobs();

        return window;
    }

    private static ListBox ListShown(CRT.Main window) =>
        window.ContributeListsShownForTests ? window.ContributeComponentListBox : window.ComponentFilterListBox;

    private static ListBox CategoriesShown(CRT.Main window) =>
        window.ContributeListsShownForTests ? window.ContributeCategoryListBox : window.CategoryFilterListBox;

    private static TextBox BoxShown(CRT.Main window) =>
        window.ContributeListsShownForTests ? window.ContributeSearchTextBox : window.ComponentSearchTextBox;

    private static List<string> Listed(CRT.Main window) =>
        ListShown(window).Items.OfType<ComponentListItem>().Select(item => item.BoardLabel).ToList();

    private static List<string> Selected(CRT.Main window) =>
        (ListShown(window).SelectedItems?.Cast<ComponentListItem>() ?? []).Select(item => item.BoardLabel).ToList();

    private static List<string> SelectedCategories(CRT.Main window) =>
        (CategoriesShown(window).SelectedItems?.Cast<string>() ?? []).ToList();

    private static void ShowContribute(CRT.Main window)
    {
        window.MainTabControl.SelectedItem = window.ContributeTabItem;
        Dispatcher.UIThread.RunJobs();
    }

    private static void ShowSchematics(CRT.Main window)
    {
        window.MainTabControl.SelectedItem = window.SchematicsTabItem;
        Dispatcher.UIThread.RunJobs();
    }

    // ###########################################################################################
    // ONE SELECTION FOR BOTH (owner request, 2026-11-08: "always keep the same selection for the left-side panel "Category"
    // and "Component" lists, meaning it should not change selections when shifting between e.g. "Schematics" tab and
    // "Contribute" tab" - an empty list on this tab until then). What the other tabs select is picked here, several at a
    // time with Ctrl or Shift; what is picked here is what they select once the tab is left. "Mark all" stays off here.
    // ###########################################################################################
    [Fact]
    public void Entering_the_tab_keeps_the_selection_and_a_pick_there_is_the_selection_back_on_the_other_tabs()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            BoxShown(window).Text = string.Empty;

            Assert.Equal(["U1"], Selected(window));
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));

            ShowContribute(window);

            Assert.True(window.ContributeListsShownForTests);
            Assert.Equal(["U1"], Selected(window));
            Assert.Equal(["U1"], window.TabDrafts.PickedComponentsForTests);
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));
            Assert.Equal(SelectionMode.Multiple, ListShown(window).SelectionMode);
            Assert.False(window.GetControl<Button>("MarkAllComponentsButton").IsEnabled);
            Assert.Equal(["U1", "U2", "C1"], Listed(window));

            window.TabDrafts.PickComponentsForTests("U2", "C1");
            Dispatcher.UIThread.RunJobs();

            ShowSchematics(window);

            Assert.False(window.ContributeListsShownForTests);
            Assert.Equal(["U2", "C1"], Selected(window));
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));
            Assert.Equal(SelectionMode.Multiple, ListShown(window).SelectionMode);
            Assert.True(window.GetControl<Button>("MarkAllComponentsButton").IsEnabled);

            ShowContribute(window);

            Assert.Equal(["U2", "C1"], Selected(window));
        });
    }

    // ###########################################################################################
    // BACK ON THE TAB THE LISTS STAND WHERE THEY STOOD (owner request, 2026-11-04: "navigate to the exact same
    // position in the 'Contribute' tab, if toggling between the main tabs"). The board read again while away -
    // what any change made here leads to as the tab is left - empties the other tabs' list, and put this one
    // back at the top while they were one list. A shown window: nothing scrolls in one that was never laid out.
    // ###########################################################################################
    [Fact]
    public void Back_on_the_tab_the_component_list_stands_where_it_stood_with_the_board_read_again_meanwhile()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [.. Enumerable.Range(1, 200).Select(i => new ComponentEntry { BoardLabel = $"U{i}", FriendlyName = $"Part {i}", Category = "IC" })],
            };

            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(entry, board);

            window.Show();
            Dispatcher.UIThread.RunJobs();
            ShowContribute(window);
            Dispatcher.UIThread.RunJobs();

            ScrollViewer scroller = ListShown(window).GetVisualDescendants().OfType<ScrollViewer>().First();
            Assert.True(scroller.Extent.Height > scroller.Viewport.Height + 500, $"The list should be long: {scroller.Extent.Height} in {scroller.Viewport.Height}");

            scroller.Offset = new Avalonia.Vector(0, 500);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(500, scroller.Offset.Y, 1);

            ShowSchematics(window);
            Assert.False(window.ContributeListsShownForTests);

            // The load: the lists let go of their rows, a layout pass runs with none, and they come again.
            window.LetGoOfListsForLoad();
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            window.UseBoardForTests(entry, board);
            Dispatcher.UIThread.RunJobs();

            ShowContribute(window);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(500, scroller.Offset.Y, 1);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** A BOARD READ AGAIN SELECTS ITS ROWS IN ONE GO, AND THE LIST STAYS WHERE IT STOOD (owner report,
    // 2026-11-05: "it slowly populated the list of components in the left-side, and then marked all
    // components", at "Align KiCad data" and again on going to the Schematics tab). *** What was selected
    // comes back as the board is read again, row by row: each row asked the list to be scrolled to it, posted,
    // a layout pass apiece, so with every row selected ("Mark all", or all a word found) the list was seen
    // running down to its end, once per read. A shown window: nothing scrolls in one never laid out.
    // ###########################################################################################
    [Fact]
    public void A_board_read_again_selects_its_rows_in_one_go_and_the_list_stays_where_it_stood()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [.. Enumerable.Range(1, 300).Select(i => new ComponentEntry { BoardLabel = $"U{i}", FriendlyName = $"Part {i}", Category = "IC" })],
            };

            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(entry, board);

            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.ComponentFilterListBox.SelectAll();
            Dispatcher.UIThread.RunJobs();

            ScrollViewer scroller = window.ComponentFilterListBox.GetVisualDescendants().OfType<ScrollViewer>().First();

            scroller.Offset = new Avalonia.Vector(0, 300);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(300, scroller.Offset.Y, 1);

            // Over to the Contribute tab, which holds the selection, and back to the Schematics tab as "Align KiCad
            // data" goes - the board read again with the list on screen, the first load after the tab was left.
            ShowContribute(window);
            ShowSchematics(window);

            int moves = 0;

            scroller.PropertyChanged += (_, change) =>
            {
                if (change.Property == ScrollViewer.OffsetProperty)
                {
                    moves++;
                }
            };

            window.LetGoOfListsForLoad();
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            // The rows stay on screen until the new ones come - as they went until 2026-11-05, an empty list was seen.
            Assert.Equal(300, window.ComponentFilterListBox.ItemCount);

            window.UseBoardForTests(entry, board);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(300, window.ComponentFilterListBox.SelectedItems!.Count);

            // Where it stood, never walked anywhere.
            Assert.Equal(0, moves);
            Assert.Equal(300, scroller.Offset.Y, 1);

            // And read again behind the Contribute tab: the same, once back.
            ShowContribute(window);
            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, board);
            Dispatcher.UIThread.RunJobs();
            ShowSchematics(window);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(300, window.ComponentFilterListBox.SelectedItems!.Count);
            Assert.Equal(0, moves);
            Assert.Equal(300, scroller.Offset.Y, 1);

            window.Close();
        });
    }

    // ###########################################################################################
    // THE ROWS KEPT AS A BOARD IS READ AGAIN TAKE NO CLICK UNTIL THE NEW ONES COME (code review, 2026-11-05). A
    // category ticked or a row selected then looked as if it had worked and was gone as the load ended, without
    // a word. They look the same meanwhile - greyed out, the flicker the kept rows remove would be back. Real
    // presses, as nothing else proves where one lands.
    // ###########################################################################################
    [Fact]
    public void The_rows_kept_as_a_board_is_read_again_take_no_click_until_the_new_ones_come()
    {
        UiTest.Run(() =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(entry, Board());

            window.Show();
            Dispatcher.UIThread.RunJobs();

            void Press(ListBox list, int row)
            {
                Control container = list.ContainerFromIndex(row)!;
                Avalonia.Point centre = container.TranslatePoint(new Avalonia.Point(container.Bounds.Width / 2, container.Bounds.Height / 2), window)!.Value;

                window.MouseMove(centre);
                window.MouseDown(centre, MouseButton.Left);
                window.MouseUp(centre, MouseButton.Left);
                Dispatcher.UIThread.RunJobs();
            }

            window.LoadBeginsForTests();
            Assert.Equal(3, window.ComponentFilterListBox.ItemCount);

            Press(window.ComponentFilterListBox, 0);
            Assert.Empty(Selected(window));

            Press(window.CategoryFilterListBox, 0);
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));

            // "Mark all" is no more the moment for it: it would select rows about to go.
            window.GetControl<Button>("MarkAllComponentsButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(Selected(window));

            // The new rows take a click again.
            window.UseBoardForTests(entry, Board());
            Dispatcher.UIThread.RunJobs();

            Press(window.ComponentFilterListBox, 0);
            Assert.Equal(["U1"], Selected(window));

            window.Close();
        });
    }

    // However a load ends - here with no board chosen to read, as for one that fails - the lists take clicks again.
    [Fact]
    public void A_load_that_ends_with_no_new_rows_leaves_the_lists_taking_clicks()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            window.ReloadCurrentBoardFromDisk(string.Empty);
            Dispatcher.UIThread.RunJobs();

            Assert.True(window.ComponentFilterListBox.IsHitTestVisible);
            Assert.True(window.CategoryFilterListBox.IsHitTestVisible);
        });
    }

    // Another board's load lets go of the rows at once: they would be read as the board chosen.
    [Fact]
    public void Another_boards_load_lets_go_of_the_rows_at_once()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            Assert.Equal(3, window.ComponentFilterListBox.ItemCount);

            window.BoardKeyOverrideForTests = "Commodore 64|250425";
            window.LetGoOfListsForLoad();

            Assert.Equal(0, window.ComponentFilterListBox.ItemCount);
        });
    }

    // And the other tabs' list, which going to this tab and back has no reason to touch.
    [Fact]
    public void Back_from_the_tab_the_other_tabs_list_stands_where_it_stood()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [.. Enumerable.Range(1, 200).Select(i => new ComponentEntry { BoardLabel = $"U{i}", FriendlyName = $"Part {i}", Category = "IC" })],
            };

            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" }, board);

            window.Show();
            Dispatcher.UIThread.RunJobs();

            ScrollViewer scroller = ListShown(window).GetVisualDescendants().OfType<ScrollViewer>().First();

            scroller.Offset = new Avalonia.Vector(0, 300);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(300, scroller.Offset.Y, 1);

            ShowContribute(window);
            Dispatcher.UIThread.RunJobs();
            ShowSchematics(window);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(300, scroller.Offset.Y, 1);

            window.Close();
        });
    }

    // ###########################################################################################
    // *** THE LIST STANDS WHERE THE ONE SHOWN BEFORE STOOD (owner request, 2026-11-09: "Toggling between tabs ... it should
    // not change position in the lists"). *** The Contribute tab's list opens where the other tabs' stood, and the other
    // tabs' opens where it was left on the Contribute tab - the same row at the top. A shown window: nothing scrolls in
    // one never laid out.
    // ###########################################################################################
    [Fact]
    public void Each_list_opens_where_the_one_shown_before_it_stood()
    {
        UiTest.Run(() =>
        {
            var board = new BoardData
            {
                Components = [.. Enumerable.Range(1, 200).Select(i => new ComponentEntry { BoardLabel = $"U{i}", FriendlyName = $"Part {i}", Category = "IC" })],
            };

            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" }, board);

            window.Show();
            Dispatcher.UIThread.RunJobs();

            ScrollViewer others = window.ComponentFilterListBox.GetVisualDescendants().OfType<ScrollViewer>().First();
            others.Offset = new Avalonia.Vector(0, 300);
            Dispatcher.UIThread.RunJobs();

            ShowContribute(window);
            Dispatcher.UIThread.RunJobs();

            ScrollViewer contribute = window.ContributeComponentListBox.GetVisualDescendants().OfType<ScrollViewer>().First();
            Assert.Equal(300, contribute.Offset.Y, 1);

            contribute.Offset = new Avalonia.Vector(0, 900);
            Dispatcher.UIThread.RunJobs();

            ShowSchematics(window);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(900, others.Offset.Y, 1);

            window.Close();
        });
    }

    // ###########################################################################################
    // LISTS OF ITS OWN IN THAT PLACE (code review, 2026-11-04): nothing done on the tab reaches the other tabs'
    // lists, so there is nothing to bring back. Lent, one pair held both, and every switch swapped it.
    // ###########################################################################################
    [Fact]
    public void The_tab_shows_lists_of_its_own_and_nothing_done_on_them_reaches_the_other_tabs_lists()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            object? rows = window.ComponentFilterListBox.ItemsSource;

            ShowContribute(window);

            Assert.False(window.ComponentFilterListBox.IsVisible);
            Assert.False(window.CategoryFilterListBox.IsVisible);
            Assert.False(window.ComponentSearchTextBox.IsVisible);
            Assert.True(window.ContributeComponentListBox.IsVisible);
            Assert.True(window.ContributeCategoryListBox.IsVisible);
            Assert.True(window.ContributeSearchTextBox.IsVisible);

            window.ContributeCategoryListBox.Selection.Deselect(0);
            Dispatcher.UIThread.RunJobs();

            window.ContributeSearchTextBox.Text = "filter";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            window.TabDrafts.PickComponentForTests("C1");
            Dispatcher.UIThread.RunJobs();

            Assert.Same(rows, window.ComponentFilterListBox.ItemsSource);
            Assert.Equal("U1", Assert.Single(window.ComponentFilterListBox.SelectedItems!.Cast<ComponentListItem>()).BoardLabel);
            Assert.Equal(["IC", "Capacitor"], window.CategoryFilterListBox.SelectedItems!.Cast<string>());
            Assert.True(string.IsNullOrEmpty(window.ComponentSearchTextBox.Text));

            ShowSchematics(window);

            Assert.True(window.ComponentFilterListBox.IsVisible);
            Assert.False(window.ContributeComponentListBox.IsVisible);
            Assert.False(window.ContributeSearchTextBox.IsVisible);
        });
    }

    // To the person typing it is one filter box, as it always stood there: the keyboard stays in it.
    [Fact]
    public void The_keyboard_in_the_filter_box_stays_in_it_across_the_tabs()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            window.Show();
            Dispatcher.UIThread.RunJobs();

            window.ComponentSearchTextBox.Focus();
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.ComponentSearchTextBox.IsFocused);

            ShowContribute(window);

            Assert.True(window.ContributeSearchTextBox.IsFocused);

            ShowSchematics(window);

            Assert.True(window.ComponentSearchTextBox.IsFocused);

            window.Close();
        });
    }

    // ###########################################################################################
    // A CATEGORY GOING AS THE LISTS ARE READ AGAIN IS NOBODY'S CHOICE. It is a selection changed while the list is
    // still changing: taken for a pick, what was ticked then was kept as the tick for good - a category just
    // come, not ticked yet, unticked by nobody - and the lists were asked for from inside their own change.
    // ###########################################################################################
    [Fact]
    public void A_category_going_as_the_lists_are_read_again_is_nobodys_choice()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);

            window.UseBoardForTests(entry, new BoardData
            {
                Components =
                [
                    new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" },
                    new ComponentEntry { BoardLabel = "D1", FriendlyName = "Rectifier", Category = "Diode" },
                ],
            });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["IC", "Diode"], window.TabDrafts.CategoryItems);
            Assert.Equal(["Diode", "IC"], SelectedCategories(window).Order());
            Assert.Equal(["U1", "D1"], Listed(window));
        });
    }

    // ###########################################################################################
    // A LOAD'S LAST STEP LIGHTS UP WHAT THE LIST HAS SELECTED BY THEN (code review, 2026-11-04): decided as the list
    // was shown, and on whether the tab was shown then, so leaving the tab while the thumbnails decoded left a
    // selection in the list with nothing lit on the picture until another was chosen.
    // ###########################################################################################
    [Fact]
    public void Leaving_the_tab_during_a_load_lights_up_what_stayed_selected_as_it_ends()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);

            // A change made there has the board read again, and the tab is left before it is in.
            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());
            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));
            Assert.Empty(window.TabSchematicsControl.highlightIndexBySchematic);

            // The rectangles come with the load's last step.
            window.TabSchematicsControl.highlightRectsBySchematicAndLabel = new(StringComparer.OrdinalIgnoreCase)
            {
                ["Top"] = new(StringComparer.OrdinalIgnoreCase) { ["U1"] = [new Rect(1, 1, 5, 5)] },
            };

            window.ShowHighlightsOfSelectionAfterLoad();

            Assert.True(window.TabSchematicsControl.highlightIndexBySchematic.ContainsKey("Top"));
        });
    }

    // ###########################################################################################
    // A COMPONENT RENAMED ON THE TAB STAYS SELECTED AS THE BOARD IS READ AGAIN (code review, 2026-11-09): what is handed back
    // is held by its row's key, which holds its names - read again with the new name, no row had it, U1 was neither selected
    // nor lit up, and coming back to the tab closed its pane. A key no row has is found by its label and region.
    // ###########################################################################################
    [Fact]
    public void A_component_renamed_on_the_tab_stays_selected_as_the_board_is_read_again()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);
            Assert.Equal(["U1"], Selected(window));

            ShowSchematics(window);

            // The board read again with U1's friendly name changed there.
            BoardData renamed = Board();
            renamed.Components[0] = new ComponentEntry { BoardLabel = "U1", FriendlyName = "Processor", Category = "IC" };
            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, renamed);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U1"], Selected(window));

            // Back on the tab it is still the pick.
            ShowContribute(window);
            Assert.Equal(["U1"], window.TabDrafts.PickedComponentsForTests);
        });
    }

    // ###########################################################################################
    // "MARK ALL" ON THE OTHER TABS IS TOLD BY THEIR LIST ITSELF, and opens nothing on the tab (third code review, 2026-11-09):
    // compared with the tab's own list, which lets through a component with no category theirs does not, it missed - and the
    // pane opened on all of them as one.
    // ###########################################################################################
    [Fact]
    public async Task Mark_all_on_the_other_tabs_is_told_by_their_list_and_opens_nothing_on_the_tab()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };
            BoardData board = Board();
            board.Components.Add(new ComponentEntry { BoardLabel = "R9", FriendlyName = "Pull-up" });

            window.UseBoardForTests(
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" },
                board);

            window.MarkAllComponentsButton.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();
            Assert.DoesNotContain("R9", Selected(window));

            ShowContribute(window);
            await window.TabDrafts.ShowEntryAsync(ContributeEntry.ComponentBasics);

            Assert.False(window.TabDrafts.IsComponentOpen);
            Assert.Contains("Every component in the list is selected", window.TabDrafts.NoComponentTextForTests, StringComparison.Ordinal);
        });
    }

    // Every row selected without "Mark all" - picked with Ctrl, or found by the filter box - is no "Mark all": the tab opens
    // them together as it opens any several (code review, 2026-11-09: it spoke of a "Mark all" nobody pressed).
    [Fact]
    public async Task Every_row_selected_without_Mark_all_is_opened_as_any_several()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnItsDraft(Board());

            ShowContribute(window);
            await window.TabDrafts.ShowEntryAsync(ContributeEntry.ComponentBasics);
            ShowSchematics(window);

            ListShown(window).SelectAll();
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(3, Selected(window).Count);

            ShowContribute(window);

            for (int attempt = 0; attempt < 200 && !window.TabDrafts.ComponentPaneForTests.IsSeveral; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.True(window.TabDrafts.IsComponentOpen);
            Assert.Equal(3, window.TabDrafts.ComponentPaneForTests.Labels.Count);
        });
    }

    // ###########################################################################################
    // THE REGION SWITCHED ON THE TAB LEAVES WHAT IS HELD, as a load does: under a word the other tabs' list
    // selected all it found, a component unticked among them too, with nobody having chosen anything.
    // ###########################################################################################
    [Fact]
    public async Task The_region_switched_on_the_tab_keeps_what_the_other_tabs_had_selected_under_a_word()
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnBoard();

            BoxShown(window).Text = "u";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            ListShown(window).Selection.Deselect(1);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U1"], Selected(window));

            ShowContribute(window);

            await SwitchRegionAsync(window, BoardRegions.Ntsc);

            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));
        });
    }

    // There and back: a row of one region has a key of its own, so the other region has nothing to select - going
    // by what the list had selected then, there was nothing left to bring back either.
    [Fact]
    public async Task The_region_switched_there_and_back_on_the_tab_brings_the_selection_back()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };

            window.PickRegion(BoardRegions.Pal);

            window.UseBoardForTests(
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" },
                new BoardData
                {
                    Components =
                    [
                        new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC", Region = "PAL" },
                        new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC", Region = "NTSC" },
                        new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
                    ],
                });

            ListShown(window).Selection.Select(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U1"], Selected(window));

            ShowContribute(window);

            await SwitchRegionAsync(window, BoardRegions.Ntsc);
            await SwitchRegionAsync(window, BoardRegions.Pal);

            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));
        });
    }

    // The region picked, and the other tabs' list having followed it - which takes a read of the rectangles first.
    // ###########################################################################################
    // *** A COMPONENT SELECTED STAYS SELECTED AS THE REGION IS SWITCHED (owner report, 2026-11-13: "when I (in "Schematics"
    // or "Overview" tab... and others?) have selected one component in the component list, and I then change region, then
    // it deselects that component and show all components"). *** A component told apart per region has a row - and a key -
    // of its own in each region, and the list made again for the other region found no row with the key selected. It is
    // the same component: its row there is selected. One of every region keeps its key, and stayed selected all along.
    // ###########################################################################################
    [Fact]
    public async Task A_component_of_each_region_stays_selected_as_the_region_is_switched()
    {
        await UiTest.RunAsync(async () =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };

            window.PickRegion(BoardRegions.Pal);

            window.UseBoardForTests(
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" },
                new BoardData
                {
                    Components =
                    [
                        new ComponentEntry { BoardLabel = "C116", FriendlyName = "Filter", TechnicalNameOrValue = "10uF", Category = "Capacitor", Region = "PAL" },
                        new ComponentEntry { BoardLabel = "C116", FriendlyName = "Filter", TechnicalNameOrValue = "47uF", Category = "Capacitor", Region = "NTSC" },
                        new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
                    ],
                });

            ShowSchematics(window);

            ListShown(window).Selection.Select(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["C116"], Selected(window));

            await SwitchRegionAsync(window, BoardRegions.Ntsc);

            Assert.Equal(["C116"], Selected(window));
            Assert.Equal("NTSC", Assert.Single(ListShown(window).SelectedItems!.Cast<ComponentListItem>()).SelectionKey.Split('\u001F')[^1]);

            await SwitchRegionAsync(window, BoardRegions.Pal);

            Assert.Equal(["C116"], Selected(window));
        });
    }

    private static async Task SwitchRegionAsync(CRT.Main window, string region)
    {
        object? rows = window.ComponentFilterListBox.ItemsSource;

        window.PickRegion(region);

        for (int attempt = 0; attempt < 200 && ReferenceEquals(rows, window.ComponentFilterListBox.ItemsSource); attempt++)
        {
            Dispatcher.UIThread.RunJobs();
            await Task.Delay(10);
        }

        Assert.False(ReferenceEquals(rows, window.ComponentFilterListBox.ItemsSource), "The other tabs' list never followed the region");
    }

    // ###########################################################################################
    // "OPEN" ON A PROBLEM OF THE OTHER REGION SWITCHES ALL OF CRT, as the buttons under the list do - they say
    // so, and the picture and the "Resources" tab follow - not the tab alone (TabDraftsProblemsTests has the page).
    // ###########################################################################################
    [Fact]
    public async Task Open_on_a_problem_of_the_other_region_switches_the_region_of_all_of_CRT()
    {
        var board = new BoardData
        {
            Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
            Components =
            [
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "VIC", Category = "IC", Region = "PAL", Package = "DIP-40" },
                new ComponentEntry { BoardLabel = "U7", FriendlyName = "VIC", Category = "IC", Region = "NTSC", Package = "DIP-40" },
            ],
            ComponentHighlights = [new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U7", X = "1", Y = "1", Width = "5", Height = "5" }],
            ComponentImages = [new ComponentImageEntry { BoardLabel = "U7", Region = "NTSC", Pin = "99", Name = "Clock", File = "Scope baseline/a.png" }],
        };

        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnItsDraft(board);

            window.PickRegion(BoardRegions.Pal);
            ShowContribute(window);

            await window.TabDrafts.ShowEntryAsync(ContributeEntry.ComponentBaselines);

            window.TabDrafts.EntryProblemRowsForTests.Single(item => item.Message.Contains("pin 99", StringComparison.Ordinal)).OpenCommand.Execute(null);

            for (int attempt = 0; attempt < 200 && !window.TabDrafts.IsComponentOpen; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal("99", Assert.Single(window.TabDrafts.ComponentPaneForTests.BaselinesForTests).Pin);
            Assert.Equal(BoardRegions.Ntsc, window.LocalRegion);
            Assert.Contains("active", window.GetControl<Button>("NtscRegionButton").Classes);
            Assert.DoesNotContain("active", window.GetControl<Button>("PalRegionButton").Classes);
        });
    }

    // ###########################################################################################
    // THE REGION SHOWN GOES WITH THE REGION DELETED (code review, 2026-11-09): "Board: Regions" deleting the region CRT showed
    // left every button unlit and the list built for a region the board no longer had. The one a board load would show is
    // shown instead, its button lit - and the tab's own region with it.
    // ###########################################################################################
    [Fact]
    public async Task Deleting_the_region_shown_shows_another_of_the_boards()
    {
        var board = new BoardData
        {
            Regions = ["PAL", "NTSC", "Japan"],
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
        };

        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnItsDraft(board);

            window.PickRegion(BoardRegions.Pal);
            ShowContribute(window);

            await window.TabDrafts.ShowEntryAsync(ContributeEntry.Regions);
            window.TabDrafts.ComponentPaneForTests.DeleteRegion("PAL");

            for (int attempt = 0; attempt < 200 && window.LocalRegion == BoardRegions.Pal; attempt++)
            {
                Dispatcher.UIThread.RunJobs();
                await Task.Delay(10);
            }

            Assert.Equal(BoardRegions.Ntsc, window.LocalRegion);
            Assert.False(window.GetControl<Button>("PalRegionButton").IsVisible);
            Assert.Contains("active", window.GetControl<Button>("NtscRegionButton").Classes);
            Assert.Equal(BoardRegions.Ntsc, window.TabDrafts.ComponentPaneForTests.RegionShownForTests);
        });
    }

    // ###########################################################################################
    // THE REGION SHOWN FOLLOWS THE BOARD READ AS THE TAB IS ENTERED AND LEFT (third code review, 2026-11-09): the draft the tab
    // shows and the board on screen can come in other regions - "view boards as officially published" - and entering the tab
    // left the region one only the board on screen had, every button unlit.
    // ###########################################################################################
    [Fact]
    public async Task The_region_shown_follows_the_board_read_as_the_tab_is_entered_and_left()
    {
        var draft = new BoardData
        {
            Regions = ["EU", "US"],
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
        };

        var published = new BoardData
        {
            Regions = ["PAL", "NTSC"],
            Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU", Category = "IC" }],
        };

        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnItsDraft(draft);
            window.UseBoardForTests(window.TabDrafts.HardwareBoardsOverrideForTests![0], published);
            window.PickRegion(BoardRegions.Pal);

            // The draft's own part open: the tab reads the draft.
            ShowContribute(window);
            await window.TabDrafts.ShowEntryAsync(ContributeEntry.Regions);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal("EU", window.LocalRegion);

            // Left: the board on screen's.
            ShowSchematics(window);
            Assert.Equal(BoardRegions.Pal, window.LocalRegion);
            Assert.Contains("active", window.GetControl<Button>("PalRegionButton").Classes);

            // Entered again: the draft's once more.
            ShowContribute(window);
            Assert.Equal("EU", window.LocalRegion);
            Assert.False(window.GetControl<Button>("PalRegionButton").IsVisible);
        });
    }

    // ###########################################################################################
    // ON THE TAB, EACH COMPONENT'S ROW SAYS WHAT IS WRONG WITH IT (owner request, 2026-11-04: "this should
    // also show on each component in the component list in the left-side") - in sight however long its name
    // is, which is what they are looked for by. On no other tab: the list is that tab's again.
    // ###########################################################################################
    [Fact]
    public void On_the_tab_a_components_row_shows_its_badges_in_sight_and_on_the_others_none()
    {
        const string boardFileName = "Commodore/C64/250407/board.json";

        var board = new BoardData
        {
            Components =
            [
                new ComponentEntry
                {
                    BoardLabel = "U1",
                    FriendlyName = "A very long friendly name that will not fit in the list on the far left at all",
                    TechnicalNameOrValue = "And a long technical name after it too",
                    Category = "IC",
                },
                new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "IC" },
            ],
            ComponentHighlights =
            [
                new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" },
                new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "C1", X = "1", Y = "1", Width = "5", Height = "5" },
            ],
            Schematics = [new BoardSchematicEntry { SchematicName = "Top", SchematicImageFile = "top.png" }],
            ComponentLinks = [new ComponentLinkEntry { BoardLabel = "U1", Name = "Bad", Url = "ftp://example.com" }],
        };

        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, boardFileName);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, boardFileName), new DraftMarker { BoardKey = boardFileName });

        UiTest.Run(() =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = boardFileName };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.TabDrafts.HardwareBoardsOverrideForTests = [entry];
            window.TabDrafts.PublishedBoardOverrideForTests = _ => null;
            window.TabDrafts.RefreshDrafts();
            window.UseBoardForTests(entry, board);

            window.Show();
            ShowContribute(window);
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            List<Border> Badges(int row) => ListShown(window).ContainerFromIndex(row)!
                .GetVisualDescendants().OfType<Border>()
                .Where(border => border.Classes.Contains("NavBadge") && border.IsEffectivelyVisible)
                .ToList();

            Border error = Assert.Single(Badges(0));
            Assert.Contains("error", error.Classes);

            double right = error.TranslatePoint(new Point(error.Bounds.Width, 0), ListShown(window))!.Value.X;

            Assert.True(right <= ListShown(window).Bounds.Width, $"The badge ends at {right} in a list {ListShown(window).Bounds.Width} wide");

            Assert.Empty(Badges(1));

            // Another tab's list says nothing of the draft's checks.
            ShowSchematics(window);
            window.UpdateLayout();
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(Badges(0));

            window.Close();
        });
    }

    // One selection for both (owner request, 2026-11-08): a category unticked on the other tabs is unticked here too -
    // every category ticked the first time here until then.
    [Fact]
    public void A_category_unticked_on_the_other_tabs_is_unticked_on_this_one_too()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            CategoriesShown(window).Selection.Deselect(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["Capacitor"], SelectedCategories(window));

            ShowContribute(window);

            Assert.Equal(["Capacitor"], SelectedCategories(window));
            Assert.Equal(["C1"], Listed(window));

            ShowSchematics(window);

            Assert.Equal(["Capacitor"], SelectedCategories(window));
        });
    }

    // ###########################################################################################
    // A CATEGORY UNTICKED HERE IS UNTICKED ON THE OTHER TABS - the board's filter, as a tick there is (owner request,
    // 2026-11-08: one selection for both; never the board's filter until then) - AND THE WORD TYPED HERE IS THEIRS TOO
    // (owner request, 2026-11-09: "it should be the exact same lists"; each box its own until then): it narrows their list
    // as it narrows this one, selecting nothing - a component picked here that it hides is selected nowhere there. Back
    // here with nothing done there, this tab's lists are as they were, that pick included.
    // ###########################################################################################
    [Fact]
    public void A_category_unticked_and_a_word_typed_on_the_tab_are_the_other_tabs_too()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            CategoriesShown(window).Selection.Deselect(CategoriesShown(window).Items.Cast<string>().ToList().IndexOf("IC"));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["C1"], Listed(window));

            BoxShown(window).Text = "filter";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            ShowSchematics(window);

            Assert.Equal(["Capacitor"], SelectedCategories(window));
            Assert.Equal(["Capacitor"], UserSettings.GetSelectedCategories(MainContributeListsTests.BoardKey));
            Assert.Equal(["C1"], Listed(window));
            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Empty(Selected(window));
            Assert.False(window.IsComponentSearchPendingForTests);

            ShowContribute(window);

            Assert.Equal(["Capacitor"], SelectedCategories(window));
            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Equal(["C1"], Listed(window));
            Assert.Equal(["U1"], window.TabDrafts.PickedComponentsForTests);
        });
    }

    // ###########################################################################################
    // A NEW BOARD RENAMED UNDER "Board: Name" KEEPS WHAT ITS LISTS SAY (code review, 2026-11-04). Everything
    // else this machine files under the board's name follows it (BoardKeyMove): the categories unticked and the
    // word typed started from nothing under the new name, and what the other tabs had selected did not come back.
    // ###########################################################################################
    [Fact]
    public void A_board_renamed_keeps_what_its_lists_say_here_and_what_the_other_tabs_had_selected()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            CategoriesShown(window).Selection.Deselect(CategoriesShown(window).Items.Cast<string>().ToList().IndexOf("IC"));
            Dispatcher.UIThread.RunJobs();

            BoxShown(window).Text = "filter";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            Assert.Equal(["C1"], Listed(window));

            const string renamedKey = "Commodore 64|999999";

            window.CarryListsToRenamedBoard(MainContributeListsTests.BoardKey, renamedKey);
            window.BoardKeyOverrideForTests = renamedKey;

            window.UseBoardForTests(
                new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "999999", ExcelDataFile = "Commodore/C64/999999/board.json" },
                Board());
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Capacitor"], SelectedCategories(window));
            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Equal(["C1"], Listed(window));

            // One selection for both (2026-11-08): what was unticked here is the renamed board's on the other tabs too.
            ShowSchematics(window);

            Assert.Equal(["Capacitor"], SelectedCategories(window));
        });
    }

    // ###########################################################################################
    // A word typed just before leaving the tab, its timer still waiting: it is taken then, and goes across with the rest
    // (owner request, 2026-11-09: "the exact same lists") - and no timer runs after the lists are given back (code review,
    // 2026-11-04: it fired after they were back, replaced them and let go of the selection just brought back).
    // ###########################################################################################
    [Fact]
    public void A_word_typed_just_before_leaving_the_tab_is_taken_and_goes_across_with_no_timer_left()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            BoxShown(window).Text = "filter";
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.IsComponentSearchPendingForTests);

            ShowSchematics(window);

            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["C1"], Listed(window));
            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Empty(Selected(window));

            ShowContribute(window);

            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Equal(["C1"], Listed(window));
        });
    }

    // ###########################################################################################
    // A word in the box, however long ago it was typed: what is given back stays. Avalonia posts TextChanged, so setting
    // the box behind a flag started the filter's timer all the same, and 200 ms later it replaced the lists and let go of
    // the selection (code review, 2026-11-04; the owner's "Leaving takes them back" did not hold for 200 ms) - and on the
    // other tabs a word applied selects all it finds. So the box the word is given to starts no timer.
    // ###########################################################################################
    [Fact]
    public void Leaving_the_tab_with_a_word_in_the_filter_starts_no_filter_on_the_other_tabs()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            Assert.Equal(["U1"], Selected(window));

            ShowContribute(window);
            Assert.False(window.IsComponentSearchPendingForTests);

            BoxShown(window).Text = "filter";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();
            window.TabDrafts.PickComponentsForTests("C1");
            Dispatcher.UIThread.RunJobs();

            ShowSchematics(window);
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["C1"], Selected(window));

            // And back again, its word with it: nothing waits there either.
            ShowContribute(window);

            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["C1"], Listed(window));
        });
    }

    // ###########################################################################################
    // A WORD TYPED AS THE BOARD IS READ AGAIN STAYS IN THE BOX (code review, 2026-11-04). A load ending wrote
    // what the tab kept for the board over it: the last keystrokes, their timer still waiting. The word gone, and
    // the list not narrowed by it. A load has nothing to do with this box, so one whose timer fires meanwhile
    // is taken as ever.
    // ###########################################################################################
    [Fact]
    public void A_word_typed_as_the_board_is_read_again_stays_in_the_box_and_narrows_the_list()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);

            // Its timer still waiting as the load ends.
            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.IsComponentSearchPendingForTests);

            window.ShowBoardInContributeTab(entry);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("sid", BoxShown(window).Text);
            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["U2"], Listed(window));

            // Its timer fired during the load.
            BoxShown(window).Text = "filter";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();
            Assert.Equal(["C1"], Listed(window));

            window.ShowBoardInContributeTab(entry);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Equal(["C1"], Listed(window));

            // Kept for the board, as a word its timer takes is.
            ShowSchematics(window);
            ShowContribute(window);

            Assert.Equal("filter", BoxShown(window).Text);
            Assert.Equal(["C1"], Listed(window));
        });
    }

    // What the box was only left holding is nobody's typing: the board shown has its own word, and gets it back.
    [Fact]
    public void Another_board_shown_has_its_own_word_in_the_box()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            window.BoardKeyOverrideForTests = "Commodore 64|250425";
            window.ShowBoardInContributeTab(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250425", ExcelDataFile = "Commodore/C64/250425/board.json" });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(string.Empty, BoxShown(window).Text);

            window.BoardKeyOverrideForTests = MainContributeListsTests.BoardKey;
            window.ShowBoardInContributeTab(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("sid", BoxShown(window).Text);
            Assert.False(window.IsComponentSearchPendingForTests);
        });
    }

    // ###########################################################################################
    // "LEAVING TAKES THEM BACK AS THEY WERE", A WORD IN THE BOX OR NOT (code review, 2026-11-04). The word selected
    // everything it finds when the lists came back, so a component unticked among its matches was selected and
    // lit up again, and what was held to come back was never asked - across the load that follows too.
    // ###########################################################################################
    [Fact]
    public void Coming_back_with_a_word_in_the_filter_brings_back_what_was_selected_not_all_it_finds()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            BoxShown(window).Text = "u";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            Assert.Equal(["U1", "U2"], Listed(window));
            Assert.Equal(["U1", "U2"], Selected(window).Order());

            ListShown(window).Selection.Deselect(1);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["U1"], Selected(window));

            ShowContribute(window);
            ShowSchematics(window);

            Assert.Equal("u", BoxShown(window).Text);
            Assert.Equal(["U1", "U2"], Listed(window));
            Assert.Equal(["U1"], Selected(window));

            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());

            Assert.Equal(["U1"], Selected(window));
        });
    }

    // Nothing selected is what was selected, too.
    [Fact]
    public void Coming_back_with_a_word_in_the_filter_and_nothing_selected_selects_nothing()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            BoxShown(window).Text = "u";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            ListShown(window).Selection.Clear();
            Dispatcher.UIThread.RunJobs();

            ShowContribute(window);
            ShowSchematics(window);

            Assert.Equal(["U1", "U2"], Listed(window));
            Assert.Empty(Selected(window));
        });
    }

    // A word typed since is somebody's choice, as a selection made since is: what it finds is selected, then and
    // after a load - nothing held from before comes back in its place.
    [Fact]
    public void A_word_typed_after_coming_back_lets_go_of_the_selection_held()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);
            ShowSchematics(window);
            Assert.Equal(["U1"], Selected(window));

            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();
            Assert.Equal(["U2"], Selected(window));

            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());

            Assert.Equal(["U2"], Selected(window));
        });
    }

    // Typed and taken back before the filter ran: the lists show that already, so nothing runs - and what is
    // selected stays. A word that does differ still waits for its timer.
    [Fact]
    public void A_word_typed_and_taken_back_before_the_filter_ran_changes_nothing()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            BoxShown(window).Text = "cpu";
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.IsComponentSearchPendingForTests);

            BoxShown(window).Text = string.Empty;
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["U1"], Selected(window));

            BoxShown(window).Text = "cpu";
            Dispatcher.UIThread.RunJobs();
            Assert.True(window.IsComponentSearchPendingForTests);

            window.FlushComponentSearchForTests();
            Assert.Equal(["U1"], Listed(window));

            // "Clear" sets the box itself: nothing is left waiting to run after it.
            window.GetControl<Button>("ClearComponentsButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.False(window.IsComponentSearchPendingForTests);
            Assert.Equal(["U1", "U2", "C1"], Listed(window));
        });
    }

    // The component picked stays too, whichever tab was shown meanwhile.
    [Fact]
    public void The_component_picked_on_the_tab_is_still_picked_after_another_tab()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            window.TabDrafts.PickComponentForTests("U2");
            Dispatcher.UIThread.RunJobs();

            ShowSchematics(window);
            ShowContribute(window);

            Assert.Equal(["U2"], Selected(window));
            Assert.Equal("U2", window.TabDrafts.PickedComponentForTests);
        });
    }

    // Another board has lists of its own, as they were left on it.
    [Fact]
    public void Each_board_keeps_its_own_ticks()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);
            CategoriesShown(window).Selection.Deselect(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["Capacitor"], SelectedCategories(window));

            window.BoardKeyOverrideForTests = "Commodore 64|250425";
            window.ShowBoardInContributeTab(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250425", ExcelDataFile = "Commodore/C64/250425/board.json" });
            Dispatcher.UIThread.RunJobs();

            // In the order they were ticked, which is no order of the list's.
            Assert.Equal(["Capacitor", "IC"], SelectedCategories(window).Order());

            window.BoardKeyOverrideForTests = MainContributeListsTests.BoardKey;
            window.ShowBoardInContributeTab(new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" });
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["Capacitor"], SelectedCategories(window));
        });
    }

    // ###########################################################################################
    // WHILE ANOTHER BOARD IS BEING READ THE TAB'S LISTS STILL SHOW THE BOARD BEFORE (code review, 2026-11-05). What is
    // typed meanwhile is typed for the board chosen, and taking it went by those lists: their ticks were filed as that
    // board's, a category unticked on the one board turning up unticked on the other, and what was unticked on the
    // other gone.
    // ###########################################################################################
    [Fact]
    public void A_word_typed_while_another_board_is_read_leaves_the_ticks_of_both_boards_alone()
    {
        UiTest.Run(() =>
        {
            const string otherKey = "Commodore 64|250425";

            var other = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250425", ExcelDataFile = "Commodore/C64/250425/board.json" };
            var first = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            // "Capacitor" unticked on the other board, and "IC" on this one.
            window.BoardKeyOverrideForTests = otherKey;
            window.ShowBoardInContributeTab(other);
            Dispatcher.UIThread.RunJobs();
            CategoriesShown(window).Selection.Deselect(1);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["IC"], SelectedCategories(window));

            window.BoardKeyOverrideForTests = MainContributeListsTests.BoardKey;
            window.ShowBoardInContributeTab(first);
            Dispatcher.UIThread.RunJobs();
            CategoriesShown(window).Selection.Deselect(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Equal(["Capacitor"], SelectedCategories(window));

            // The other board chosen, and a word typed for it before it has been read.
            window.BoardKeyOverrideForTests = otherKey;
            window.FollowBoardInContributeLists();
            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            window.ShowBoardInContributeTab(other);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal("sid", BoxShown(window).Text);
            Assert.Equal(["IC"], SelectedCategories(window));

            window.BoardKeyOverrideForTests = MainContributeListsTests.BoardKey;
            window.ShowBoardInContributeTab(first);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(string.Empty, BoxShown(window).Text);
            Assert.Equal(["Capacitor"], SelectedCategories(window));
        });
    }

    // The box narrows the list here, and selects nothing - everywhere else what it finds is selected.
    [Fact]
    public void The_filter_on_the_tab_narrows_and_selects_nothing()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            Assert.Equal(["U2"], Listed(window));
            Assert.Empty(Selected(window));
        });
    }

    // What was typed on the other tabs is the tab's word too - narrowing, what it found picked - and comes back with the rest.
    [Fact]
    public void A_filter_typed_before_the_tab_is_its_word_too_and_comes_back_after_it()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            BoxShown(window).Text = "cpu";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            ShowContribute(window);

            Assert.Equal("cpu", BoxShown(window).Text);
            Assert.Equal(["U1"], Listed(window));
            Assert.Equal(["U1"], window.TabDrafts.PickedComponentsForTests);

            ShowSchematics(window);

            Assert.Equal("cpu", BoxShown(window).Text);
            Assert.Equal(["U1"], Listed(window));
        });
    }

    // "Clear" goes back to how the tab starts: every category, nothing typed, no component picked.
    [Fact]
    public void Clear_on_the_tab_goes_back_to_how_the_tab_starts()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            CategoriesShown(window).Selection.Deselect(1);
            Dispatcher.UIThread.RunJobs();

            window.GetControl<Button>("ClearComponentsButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(string.Empty, BoxShown(window).Text);
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));
            Assert.Equal(["U1", "U2", "C1"], Listed(window));
            Assert.Equal(string.Empty, window.TabDrafts.PickedComponentForTests);
        });
    }

    // ###########################################################################################
    // WHAT THE PANE SAVED IS READ AGAIN AS THE TAB IS LEFT, AND A LOAD DROPS A SELECTION: what was selected
    // before the tab is held until that load has shown it - unless somebody has chosen since.
    // ###########################################################################################
    [Fact]
    public void The_selection_brought_back_survives_the_board_being_read_again_as_the_tab_is_left()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);
            ShowSchematics(window);

            // The lists as a load makes them, of rows made anew.
            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());

            Assert.Equal(["U1"], Selected(window));

            // Chosen by hand since: gone is gone, the next load has nothing to bring back.
            ShowContribute(window);
            ShowSchematics(window);
            ListShown(window).Selection.Clear();
            Dispatcher.UIThread.RunJobs();

            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());

            Assert.Empty(Selected(window));
        });
    }

    // ###########################################################################################
    // BACK ON THE TAB WHILE THAT LOAD RUNS, what the other tabs had selected is held all the same (code review,
    // 2026-11-04). The load has let go of the lists by then, so going by what they had selected left nothing to
    // bring back after the tab.
    // ###########################################################################################
    [Fact]
    public void Coming_back_to_the_tab_while_the_board_is_read_again_loses_no_selection()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);
            ShowSchematics(window);

            // What the pane saved is being read again, and CRT is back on the tab before it is in.
            window.LetGoOfListsForLoad();
            ShowContribute(window);
            window.UseBoardForTests(entry, Board());

            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));
        });
    }

    // ###########################################################################################
    // ENTERING THE TAB WHILE THE BOARD ON SCREEN IS READ AGAIN (code review, 2026-11-05). Its rows stay on the list until
    // the new ones come, selection and all, so there is something to hold - and that load would have dropped it for
    // good, with nothing held to bring back after the tab.
    // ###########################################################################################
    [Fact]
    public void Entering_the_tab_while_the_board_is_read_again_holds_what_its_list_still_has_selected()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            Assert.Equal(["U1"], Selected(window));

            window.LoadBeginsForTests();
            ShowContribute(window);
            window.UseBoardForTests(entry, Board());

            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));
        });
    }

    // Another board's load has taken the rows, so there is nothing to hold - and a word in the box is left to select
    // all it finds on that board, as it would have.
    [Fact]
    public void Entering_the_tab_while_another_board_is_read_holds_nothing()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250425", ExcelDataFile = "Commodore/C64/250425/board.json" };

            BoxShown(window).Text = "sid";
            Dispatcher.UIThread.RunJobs();
            window.FlushComponentSearchForTests();

            window.BoardKeyOverrideForTests = "Commodore 64|250425";
            window.LoadBeginsForTests();
            ShowContribute(window);
            window.UseBoardForTests(entry, Board());

            ShowSchematics(window);

            Assert.Equal(["U2"], Selected(window));
        });
    }

    // ###########################################################################################
    // A CATEGORY CHANGED AFTER COMING BACK IS SOMEBODY'S CHOICE, AS A SELECTION MADE AND A WORD TYPED ARE (code review,
    // 2026-11-05). U1 went from the list with its category and came back not selected, and what was still held from
    // before the tab selected it again at the next load, or as the region was switched.
    // ###########################################################################################
    [Fact]
    public void A_category_changed_after_coming_back_lets_go_of_the_selection_held()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            ShowContribute(window);
            ShowSchematics(window);

            Assert.Equal(["U1"], Selected(window));

            CategoriesShown(window).Selection.Deselect(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(Selected(window));

            CategoriesShown(window).Selection.Select(0);
            Dispatcher.UIThread.RunJobs();
            Assert.Empty(Selected(window));

            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());

            Assert.Empty(Selected(window));
        });
    }

    // ###########################################################################################
    // A BOARD READ AGAIN THAT DROPS THE SELECTION TAKES ITS HIGHLIGHTS WITH IT (code review, 2026-11-05). The rows stay
    // until the new ones come, so nothing said the selection had gone: what the Schematics tab goes by - the labels
    // highlighted, the blink timer - stood as it was until another component was chosen.
    // ###########################################################################################
    [Fact]
    public void A_selection_dropped_by_a_load_lets_go_of_its_highlights_as_the_load_ends()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };

            window.TabSchematicsControl.highlightRectsBySchematicAndLabel = new(StringComparer.OrdinalIgnoreCase)
            {
                ["Top"] = new(StringComparer.OrdinalIgnoreCase) { ["U1"] = [new Rect(1, 1, 5, 5)] },
            };

            window.ShowHighlightsOfSelectionAfterLoad();
            Assert.True(window.TabSchematicsControl.highlightIndexBySchematic.ContainsKey("Top"));

            window.LetGoOfListsForLoad();
            window.UseBoardForTests(entry, Board());
            Assert.Empty(Selected(window));

            window.ShowHighlightsOfSelectionAfterLoad();

            Assert.Empty(window.TabSchematicsControl.highlightIndexBySchematic);
        });
    }

    // ###########################################################################################
    // CTRL+CLICK AND SHIFT+CLICK PICK SEVERAL ON THE TAB ITSELF - real presses (owner report, 2026-11-08: "I cannot mark
    // multiple components when I am in "Contribute" tab ... it is not possible from within the tab itself"). The click's
    // Tapped, which opens again the one component clicked when no selection changes, took a Ctrl+click as that and put the
    // list back to the one clicked.
    // ###########################################################################################
    [Fact]
    public void Ctrl_click_and_Shift_click_pick_several_on_the_tab()
    {
        UiTest.Run(() =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(entry, Board());
            window.Show();
            Dispatcher.UIThread.RunJobs();
            ShowContribute(window);
            window.UpdateLayout();

            // Each press once the open the one before started has ended: it runs under CRT's "please wait" cover, which takes
            // every click meanwhile by design (BusyOverlay) - under the full suite's load a press came while it was still up.
            void Press(int row, RawInputModifiers modifiers = RawInputModifiers.None)
            {
                BusyOverlay overlay = BusyOverlay.For(window)!;

                for (int wait = 0; wait < 600 && overlay.IsBusy; wait++)
                {
                    Dispatcher.UIThread.RunJobs();
                    Thread.Sleep(5);
                }

                Assert.False(overlay.IsBusy, "An open never ended.");

                ListBox list = window.ContributeComponentListBox;
                Control container = list.ContainerFromIndex(row)!;
                Avalonia.Point centre = container.TranslatePoint(new Avalonia.Point(container.Bounds.Width / 2, container.Bounds.Height / 2), window)!.Value;

                window.MouseMove(centre, modifiers);
                window.MouseDown(centre, MouseButton.Left, modifiers);
                window.MouseUp(centre, MouseButton.Left, modifiers);
                Dispatcher.UIThread.RunJobs();
            }

            Press(0);
            Assert.Equal(["U1"], Selected(window));

            Press(2, RawInputModifiers.Control);
            Assert.Equal(["U1", "C1"], Selected(window));
            Assert.Equal(["U1", "C1"], window.TabDrafts.PickedComponentsForTests);

            // A plain click picks the one clicked alone again.
            Press(1);
            Assert.Equal(["U2"], Selected(window));
            Assert.Equal(["U2"], window.TabDrafts.PickedComponentsForTests);

            Press(2, RawInputModifiers.Shift);
            Assert.Equal(["U2", "C1"], Selected(window));
            Assert.Equal(["U2", "C1"], window.TabDrafts.PickedComponentsForTests);

            window.IsApplicationShuttingDownForTests = true;
            window.Close();
        });
    }

    // ###########################################################################################
    // THE OTHER TABS' COMPONENT LIST PICKS AS THE CONTRIBUTE TAB'S DOES (owner request, 2026-11-08: "in "Schematics" then
    // clicking one component only will TOGGLE that component on/off? I want it to work like "Contribute""): a click selects
    // the one clicked alone, Ctrl+click adds or takes one away, Shift+click a range. "Multiple,Toggle" since the first
    // commit (2026-03-07) - every click toggled. The categories stay ticks, a click turning one on or off.
    // ###########################################################################################
    [Fact]
    public void On_the_other_tabs_a_click_selects_one_and_Ctrl_and_Shift_add_as_on_the_Contribute_tab()
    {
        UiTest.Run(() =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

            window.UseBoardForTests(entry, Board());
            window.Show();
            Dispatcher.UIThread.RunJobs();
            window.UpdateLayout();

            void Press(ListBox list, int row, RawInputModifiers modifiers = RawInputModifiers.None)
            {
                Control container = list.ContainerFromIndex(row)!;
                Avalonia.Point centre = container.TranslatePoint(new Avalonia.Point(container.Bounds.Width / 2, container.Bounds.Height / 2), window)!.Value;

                window.MouseMove(centre, modifiers);
                window.MouseDown(centre, MouseButton.Left, modifiers);
                window.MouseUp(centre, MouseButton.Left, modifiers);
                Dispatcher.UIThread.RunJobs();
            }

            ListBox components = window.ComponentFilterListBox;

            Press(components, 0);
            Assert.Equal(["U1"], Selected(window));

            Press(components, 1);
            Assert.Equal(["U2"], Selected(window));

            Press(components, 2, RawInputModifiers.Control);
            Assert.Equal(["U2", "C1"], Selected(window));

            Press(components, 2, RawInputModifiers.Control);
            Assert.Equal(["U2"], Selected(window));

            Press(components, 0);
            Press(components, 2, RawInputModifiers.Shift);
            Assert.Equal(["U1", "U2", "C1"], Selected(window));

            // The categories are ticks: a click turns one off, and on again.
            Press(window.CategoryFilterListBox, 1);
            Assert.Equal(["IC"], SelectedCategories(window));
            Press(window.CategoryFilterListBox, 1);
            Assert.Equal(["IC", "Capacitor"], SelectedCategories(window));

            window.IsApplicationShuttingDownForTests = true;
            window.Close();
        });
    }

    // Several components selected on the tab are picked together, edited as one (owner request, 2026-11-08) - one at a
    // time until then.
    [Fact]
    public void Several_components_selected_on_the_tab_are_picked_together()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            ListShown(window).Selection.Select(0);
            ListShown(window).Selection.Select(1);
            Dispatcher.UIThread.RunJobs();

            Assert.Equal(["U1", "U2"], Selected(window));
            Assert.Equal(["U1", "U2"], window.TabDrafts.PickedComponentsForTests);
        });
    }

    // U1 has no category, so an error; C1 nothing wrong.
    private const string DraftedBoard = "Commodore/C64/250407/board.json";

    private static BoardData BoardWithAnIssue() => new()
    {
        Components =
        [
            new ComponentEntry { BoardLabel = "U1", FriendlyName = "CPU" },
            new ComponentEntry { BoardLabel = "C1", FriendlyName = "Filter", Category = "Capacitor" },
        ],
        Schematics = [new BoardSchematicEntry { SchematicName = "Top" }],
        ComponentHighlights =
        [
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" },
            new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "C1", X = "1", Y = "1", Width = "5", Height = "5" },
        ],
    };

    private static CRT.Main WindowOnItsDraft(BoardData board)
    {
        string boardFile = DraftFolderLayout.GetWorkbookPath(DraftManager.DraftsRoot, DraftedBoard);
        Directory.CreateDirectory(Path.GetDirectoryName(boardFile)!);
        CachedWorkbooks.Write(boardFile, board);
        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(DraftManager.DraftsRoot, DraftedBoard), new DraftMarker { BoardKey = DraftedBoard });

        var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = DraftedBoard };
        var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey, Width = 1300, Height = 800 };

        window.TabDrafts.HardwareBoardsOverrideForTests = [entry];
        window.TabDrafts.PublishedBoardOverrideForTests = _ => null;
        window.TabDrafts.RefreshDrafts();
        window.UseBoardForTests(entry, board);

        return window;
    }

    // ###########################################################################################
    // *** "UI" AND "TABLE" KEEP THE ENTRY IN CRT'S OWN WINDOW (owner report, 2026-11-07: "when I am in a menu and I
    // toggle between "UI" and "Table" view, then it should keep the same menu, and not jump to "Overview""). *** The
    // buttons pressed as a person presses them, on a new board's draft, with Main reading the board as it does.
    // ###########################################################################################
    [Theory]
    [InlineData("FilesNavButton")]
    [InlineData("CreditsNavButton")]
    [InlineData("RegionsNavButton")]
    [InlineData("BasicsNavButton")]
    public async Task UI_and_Table_keep_the_entry_shown(string navButton)
    {
        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnItsDraft(BoardWithAnIssue());
            window.Show();

            try
            {
                ShowContribute(window);

                TabDrafts tab = window.TabDrafts;

                static void Press(TabDrafts tab, string name) =>
                    tab.GetControl<Button>(name).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

                async Task Until(Func<bool> condition, string what)
                {
                    for (int attempt = 0; attempt < 300 && !condition(); attempt++)
                    {
                        Dispatcher.UIThread.RunJobs();
                        await Task.Delay(10);
                    }

                    Dispatcher.UIThread.RunJobs();
                    Assert.True(condition(), what);
                }

                if (navButton == "BasicsNavButton")
                {
                    ListShown(window).SelectedIndex = 0;
                    await Until(() => tab.IsComponentOpen, "No component opened.");
                }

                Press(tab, navButton);
                await Until(() => tab.SelectedNavEntryForTests == navButton, "The entry never opened.");

                for (int round = 0; round < 3; round++)
                {
                    Press(tab, "TableViewButton");
                    await Until(() => tab.GetControl<Button>("TableViewButton").Classes.Contains("Selected"), "Table never came.");
                    await Task.Delay(50);
                    Dispatcher.UIThread.RunJobs();
                    Assert.Equal(navButton, tab.SelectedNavEntryForTests);

                    // The last round with a search typed, as it was when the owner met it: a search finding rows on
                    // another sheet moves the table there, and the entry with it - the entry UI then keeps.
                    if (round == 2 && tab.IsTableOpen)
                    {
                        tab.GetControl<BoardTableEditor>("TableEditor").SearchText = "Filter";
                        Dispatcher.UIThread.RunJobs();
                        navButton = tab.SelectedNavEntryForTests!;
                        Assert.NotEqual("StatusNavButton", navButton);
                    }

                    Press(tab, "UiViewButton");
                    await Until(() => tab.GetControl<Button>("UiViewButton").Classes.Contains("Selected"), "UI never came.");
                    await Task.Delay(50);
                    Dispatcher.UIThread.RunJobs();
                    Assert.Equal(navButton, tab.SelectedNavEntryForTests);
                }
            }
            finally
            {
                window.Close();
            }
        });
    }

    // ###########################################################################################
    // "WITH ISSUES" IS THERE ON THE TAB ALONE, WHILE A COMPONENT HAS ONE (owner requests, 2026-11-04: "I only
    // want to see this checkbox when it at least has one component with an issue"; "the checkbox itself should
    // be at the right side of the text") - and takes no room from the list, which stands in one place whichever
    // tab is shown. A press anywhere on the row is a press on the box, as on the "Labels visible" panel.
    // ###########################################################################################
    [Fact]
    public void With_issues_is_there_on_the_tab_alone_its_box_after_its_text_and_moves_nothing()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnItsDraft(BoardWithAnIssue());

            window.Show();
            Dispatcher.UIThread.RunJobs();

            CheckBox box = window.ComponentsWithIssuesCheckBox;
            StackPanel row = window.GetControl<StackPanel>("ComponentsWithIssuesRow");

            Assert.False(row.IsVisible);

            double before = ListShown(window).TranslatePoint(default, window)!.Value.Y;

            ShowContribute(window);
            Dispatcher.UIThread.RunJobs();

            Assert.True(row.IsVisible);
            Assert.False(box.IsChecked);
            Assert.Equal(before, ListShown(window).TranslatePoint(default, window)!.Value.Y);

            // At the right of the heading, the box after its text. Asked of the controls, as the headless
            // platform's text has no real width.
            Assert.Equal(Avalonia.Layout.HorizontalAlignment.Right, row.HorizontalAlignment);
            Assert.Equal(1, Grid.GetColumn(row));
            Assert.IsType<TextBlock>(row.Children[0]);
            Assert.Same(box, Assert.IsType<Viewbox>(row.Children[1]).Child);

            // The word ticks it as the box does.
            Point word = row.TranslatePoint(new Point(2, row.Bounds.Height / 2), window)!.Value;

            window.MouseDown(word, MouseButton.Left);
            window.MouseUp(word, MouseButton.Left);
            Dispatcher.UIThread.RunJobs();

            Assert.True(box.IsChecked);
            Assert.Equal(["U1"], Listed(window));

            ShowSchematics(window);

            Assert.False(row.IsVisible);

            window.Close();
        });
    }

    // With nothing wrong on any component the box is not there, on the tab or anywhere.
    [Fact]
    public void With_no_issue_on_any_component_there_is_no_box()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            ShowContribute(window);

            Assert.True(window.ContributeListsShownForTests);
            Assert.False(window.GetControl<StackPanel>("ComponentsWithIssuesRow").IsVisible);
        });
    }

    // "Clear" is back to how the tab starts: nothing narrowed.
    [Fact]
    public void Clear_lets_go_of_the_box_too()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnItsDraft(BoardWithAnIssue());

            ShowContribute(window);

            CheckBox box = window.ComponentsWithIssuesCheckBox;

            Assert.True(box.IsVisible);

            box.IsChecked = true;
            Assert.Equal(["U1"], Listed(window));

            window.GetControl<Button>("ClearComponentsButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.False(box.IsChecked);
            Assert.True(box.IsVisible);
            Assert.Equal(["U1", "C1"], Listed(window));
        });
    }

    // ###########################################################################################
    // "CLEAR" WITH NO BOARD READ (code review, 2026-11-05) - one loading, or one that could not be read. It
    // filtered data that was not there, and an exception in a click handler ends CRT: App's handlers only log it.
    // ###########################################################################################
    [Fact]
    public void Clear_with_no_board_read_leaves_CRT_running()
    {
        UiTest.Run(() =>
        {
            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };

            window.GetControl<Button>("ClearComponentsButton").RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(window.ComponentFilterListBox.Items);
        });
    }

    // ###########################################################################################
    // *** WHAT STAYS SELECTED AS A BOARD IS READ AGAIN KEEPS ITS TRACES (code review, 2026-11-05). *** The list keeps
    // the selection and its rectangles light up, and then the board's KiCad data is read, which starts by letting
    // go of every net it has lit: back on the tab, the component stood selected with no copper and no pin 1.
    // ###########################################################################################
    [Fact]
    public async Task A_component_still_selected_lights_its_nets_once_the_KiCad_data_is_read()
    {
        string pcb = this.thisWorkspace.WriteFile("board.kicad_pcb", ClassicRepairToolbox.Tests.KiCadFixtures.Pcb);

        await UiTest.RunAsync(async () =>
        {
            CRT.Main window = WindowOnBoard();

            Assert.Equal(["U1"], Selected(window));

            await window.TabSchematicsControl.LoadKiCadProjectForCurrentBoardAsync([pcb]);

            Assert.Equal(["U1"], Selected(window));
            Assert.Contains("GND", window.TabSchematicsControl.SelectedKiCadNetNamesForTests);
        });
    }

    // A region picked just as the board is read again: what was worked out for the board before stays off the
    // one now shown (code review, 2026-11-05: its rectangles were put in, and only then was the load noticed).
    [Fact]
    public async Task A_region_picked_as_the_board_is_read_again_leaves_the_rectangles_of_the_board_before_out()
    {
        await UiTest.RunAsync(async () =>
        {
            var entry = new HardwareBoardEntry { HardwareName = "Commodore 64", BoardName = "250407", ExcelDataFile = "Commodore/C64/250407/board.json" };
            BoardData before = Board();

            before.ComponentHighlights.Add(new ComponentHighlightEntry { SchematicName = "Top", BoardLabel = "U1", X = "1", Y = "1", Width = "5", Height = "5" });

            var window = new CRT.Main { BoardKeyOverrideForTests = MainContributeListsTests.BoardKey };
            window.UseBoardForTests(entry, before);

            window.PickRegion(BoardRegions.Ntsc);
            window.UseBoardForTests(entry, Board());

            for (int i = 0; i < 20; i++)
            {
                await Task.Delay(10, TestContext.Current.CancellationToken);
                Dispatcher.UIThread.RunJobs();
            }

            Assert.Empty(window.TabSchematicsControl.highlightRectsBySchematicAndLabel);
        });
    }

    // A board read again keeps its rows until the new ones come - when there is no such board any more, none
    // come (code review, 2026-11-05: its rows stayed on the lists, with no board behind them).
    [Fact]
    public void A_board_no_longer_listed_takes_its_rows_off_the_lists()
    {
        UiTest.Run(() =>
        {
            CRT.Main window = WindowOnBoard();

            Assert.NotEmpty(window.ComponentFilterListBox.Items);

            window.HardwareComboBox.ItemsSource = new[] { "No such hardware" };
            window.HardwareComboBox.SelectedIndex = 0;
            window.BoardComboBox.ItemsSource = new[] { "No such board" };
            window.BoardComboBox.SelectedIndex = 0;
            Dispatcher.UIThread.RunJobs();

            Assert.Empty(window.ComponentFilterListBox.Items);
            Assert.Empty(window.CategoryFilterListBox.Items);
        });
    }
}
